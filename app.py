# app.py
# -*- coding: utf-8 -*-
"""
ジョブカン summary Excel → freee人事労務 取込用スケジュールCSV（生成のみ）

このバージョンは freee API 連携（打刻同期・有休/欠勤PUT）と OAuth 認証を
すべて削除した「CSV生成専用」版です。認証情報（.env / トークン）は不要です。

入力:  ジョブカン summary Excel（複数可・.xlsx）
出力:  freee_schedule_import.csv（固定列・UTF-8 BOM）

元コードとの差分:
  - freee API を呼ばないため認証不要
  - 「freee人事労務での表示名」カラムは空欄（取込時に無視される列のため影響なし）
  - freee未登録者を除外するフィルタが無いので、Excelにある従業員は全員出力される
"""

from __future__ import annotations

import io
import re
from datetime import datetime, date, timedelta
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


# =========================
# Config（ジョブカン summary レイアウト）
# =========================
ROW_STAFF_CODE = 4
COL_STAFF_CODE = 3       # C4
ROW_DAILY_HEADER = 10

COL_DATE = 1
COL_KINTAI_STATUS = 2
COL_HOLIDAY = 3
COL_SHIFT_START = 6
COL_SHIFT_END = 7
COL_CLOCK_IN = 8
COL_CLOCK_OUT = 9
COL_BREAK = 14
COL_PAID_LEAVE = 15

SCHEDULE_COLUMNS_FIXED = [
    "従業員番号",
    "freee人事労務での表示名（編集しても反映されません）",
    "日付",
    "勤務パターンコード",
    "勤務日種別",
    "出勤時刻",
    "退勤時刻",
    "休憩時間",
    "休憩開始1",
    "休憩終了1",
    "休憩開始2",
    "休憩終了2",
    "休憩開始3",
    "休憩終了3",
    "夜勤日種別",
]


# =========================
# Utils
# =========================
def safe_str(v: Any) -> str:
    return "" if v is None else str(v).strip()


def normalize_employee_number(s: Any) -> str:
    t = safe_str(s)
    if t == "":
        return ""
    trans = str.maketrans("０１２３４５６７８９", "0123456789")
    t = t.translate(trans).strip()
    t2 = t.lstrip("0")
    return t2 if t2 != "" else "0"


def parse_hhmm(value: Any) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%H:%M")
    if hasattr(value, "hour") and hasattr(value, "minute"):
        return f"{int(value.hour):02d}:{int(value.minute):02d}"

    s = str(value).strip()
    if s == "":
        return None

    m = re.match(r"^(\d{1,2}):(\d{1,2})$", s)
    if m:
        h = int(m.group(1)); mm = int(m.group(2))
        return f"{h:02d}:{mm:02d}"

    m = re.match(r"^(\d{2})(\d{2})$", s)
    if m:
        return f"{int(m.group(1)):02d}:{int(m.group(2)):02d}"

    raise ValueError(f"時刻の解釈に失敗: {value!r}")


def parse_yyyymm_from_a1(v: Any) -> Optional[Tuple[int, int]]:
    s = safe_str(v)
    m = re.match(r"^\s*(\d{4})年\s*(\d{1,2})月\s*$", s)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def parse_mmdd_youbi(s: str, year: int) -> date:
    m = re.match(r"^\s*(\d{1,2})/(\d{1,2})", s)
    if not m:
        raise ValueError(f"日付文字列の解釈に失敗: {s!r}")
    month = int(m.group(1)); day = int(m.group(2))
    return date(year, month, day)


def combine_date_time(d: date, hhmm: str) -> datetime:
    h, m = map(int, hhmm.split(":"))
    day_add = 0
    if h >= 24:
        day_add = h // 24
        h = h % 24
    base = datetime(d.year, d.month, d.day, h, m, 0)
    if day_add:
        base += timedelta(days=day_add)
    return base


def break_hhmm_to_minutes(hhmm: Optional[str]) -> int:
    if not hhmm:
        return 0
    h, m = map(int, hhmm.split(":"))
    return h * 60 + m


def minutes_to_hhmm(mins: int) -> str:
    if mins <= 0:
        return ""
    h = mins // 60
    m = mins % 60
    return f"{h:02d}:{m:02d}"


def parse_duration_to_minutes(value: Any) -> int:
    if value is None:
        return 0
    if isinstance(value, datetime):
        return int(value.hour) * 60 + int(value.minute)
    if hasattr(value, "hour") and hasattr(value, "minute"):
        return int(value.hour) * 60 + int(value.minute)
    if isinstance(value, (int, float)):
        if value <= 0:
            return 0
        return int(round(float(value) * 60))
    s = str(value).strip()
    if s == "":
        return 0
    m = re.match(r"^(\d{1,2}):(\d{1,2})$", s)
    if m:
        hh = int(m.group(1)); mm = int(m.group(2))
        if hh < 0 or mm < 0:
            return 0
        return hh * 60 + mm
    try:
        f = float(s)
        if f <= 0:
            return 0
        return int(round(f * 60))
    except Exception:
        return 0


def compute_shift_duration_minutes(wd: date, shift_start: Optional[str], shift_end: Optional[str]) -> int:
    if not shift_start or not shift_end:
        return 0
    sdt = combine_date_time(wd, shift_start)
    edt = combine_date_time(wd, shift_end)
    if edt <= sdt:
        edt += timedelta(days=1)
    mins = int(round((edt - sdt).total_seconds() / 60))
    return max(0, mins)


def df_to_csv_bytes_utf8sig(df: pd.DataFrame) -> bytes:
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8-sig")


def normalize_holiday_kbn(raw: Any) -> str:
    s = safe_str(raw)
    if s == "":
        return "所定労働日"
    s2 = s.replace("　", " ").replace(",", ", ").strip()
    if "法休" in s2:
        return "法定休日"
    if "公休" in s2:
        return "所定休日"
    return "所定労働日"


def is_zero_zero_shift(shift_start: Optional[str], shift_end: Optional[str]) -> bool:
    s = (shift_start or "").strip()
    e = (shift_end or "").strip()
    if s == "" and e == "":
        return True
    return (s in ("0:00", "00:00") and e in ("0:00", "00:00"))


# =========================
# Jobcan reader
# =========================
def find_sheet_staff_code(ws) -> str:
    return normalize_employee_number(ws.cell(ROW_STAFF_CODE, COL_STAFF_CODE).value)


def read_one_sheet_rows(ws, year: int, month: int) -> List[Dict[str, Any]]:
    if safe_str(ws.cell(ROW_DAILY_HEADER, COL_DATE).value) != "日付":
        return []

    rows: List[Dict[str, Any]] = []
    r = ROW_DAILY_HEADER + 1

    while r <= ws.max_row:
        s = safe_str(ws.cell(r, COL_DATE).value)
        if s == "" or s in ("合計", "小計", "総計"):
            break
        if not re.match(r"^\d{1,2}/\d{1,2}", s):
            break

        wd = parse_mmdd_youbi(s, year)
        if wd.month != month:
            r += 1
            continue

        kintai_status = safe_str(ws.cell(r, COL_KINTAI_STATUS).value)
        holiday_kbn = safe_str(ws.cell(r, COL_HOLIDAY).value)

        shift_start = parse_hhmm(ws.cell(r, COL_SHIFT_START).value)
        shift_end = parse_hhmm(ws.cell(r, COL_SHIFT_END).value)
        if shift_start == "00:00" and shift_end == "00:00":
            shift_start, shift_end = None, None

        clock_in = parse_hhmm(ws.cell(r, COL_CLOCK_IN).value)
        clock_out = parse_hhmm(ws.cell(r, COL_CLOCK_OUT).value)
        if clock_in == "00:00" and clock_out == "00:00":
            clock_in, clock_out = None, None

        break_hhmm = parse_hhmm(ws.cell(r, COL_BREAK).value)
        break_mins = break_hhmm_to_minutes(break_hhmm)

        paid_leave_raw = ws.cell(r, COL_PAID_LEAVE).value
        paid_leave_mins = parse_duration_to_minutes(paid_leave_raw)

        rows.append({
            "work_date": wd,
            "kintai_status": kintai_status,
            "holiday_kbn": holiday_kbn,
            "shift_start": shift_start,
            "shift_end": shift_end,
            "clock_in": clock_in,
            "clock_out": clock_out,
            "break_hhmm": break_hhmm,
            "break_mins": break_mins,
            "paid_leave_raw": paid_leave_raw,
            "paid_leave_mins": paid_leave_mins,
            "source_row": r,
        })
        r += 1

    return rows


def read_jobcan_excels(files: List[Tuple[str, bytes]]) -> Tuple[int, int, Dict[str, List[Dict[str, Any]]]]:
    staff_rows: Dict[str, List[Dict[str, Any]]] = {}
    detected_year: Optional[int] = None
    detected_month: Optional[int] = None
    detected_from: Optional[str] = None

    for fname, content in files:
        wb = load_workbook(io.BytesIO(content), data_only=True)
        for sname in wb.sheetnames:
            ws = wb[sname]

            ym = parse_yyyymm_from_a1(ws.cell(1, 1).value)
            if ym is None:
                continue
            y, m = ym

            if detected_year is None:
                detected_year, detected_month = y, m
                detected_from = f"{fname}:{sname}"
            else:
                if y != detected_year or m != detected_month:
                    raise ValueError(
                        f"Excelの年月が混在しています。最初={detected_year}年{detected_month}月（{detected_from}） / "
                        f"別={y}年{m}月（{fname}:{sname}）"
                    )

            staff_code = find_sheet_staff_code(ws)
            if staff_code == "":
                continue

            rows = read_one_sheet_rows(ws, detected_year, detected_month)
            if rows:
                staff_rows.setdefault(staff_code, []).extend(rows)

    if detected_year is None or detected_month is None:
        raise ValueError("年月を確定できませんでした。各シートのA1が 'YYYY年MM月' になっているか確認してください。")

    normalized: Dict[str, List[Dict[str, Any]]] = {}
    for code, rows in staff_rows.items():
        by_date: Dict[str, Dict[str, Any]] = {}
        for rec in rows:
            k = rec["work_date"].strftime("%Y-%m-%d")
            by_date[k] = rec
        normalized[code] = [by_date[k] for k in sorted(by_date.keys())]

    return detected_year, detected_month, normalized


# =========================
# Schedule CSV (fixed columns)
# =========================
def build_schedule_csv_fixed(
    staff_rows: Dict[str, List[Dict[str, Any]]],
) -> Tuple[pd.DataFrame, bytes]:
    """
    固定列CSVを生成する（freee API 不使用）。

    休憩時間のルール（優先順）:
      1) シフト開始・終了が 0:00/0:00（または空） -> 休憩は必ず空欄
      2) シフト時間が 6時間未満 -> 休憩なし（空欄）
      3) シフト時間が 9時間(540分) -> 休憩は "01:00" 固定
      4) シフト時間が 9時間(540分)を超える -> (シフト - 480分) が休憩
         例) 7:00-17:00(600分) -> 休憩120分(02:00)
      5) それ以外（360〜540の間） -> Excel由来の休憩時間（分）を採用（上限丸め）

    有休行（B列=有休）:
      - 勤務日種別は "所定労働日"
      - 出勤/退勤はシフト開始/終了
      - 休憩は「通常のシフト休憩ルール」で決める
    """
    cols = SCHEDULE_COLUMNS_FIXED[:]
    out_rows: List[Dict[str, Any]] = []

    def new_row() -> Dict[str, Any]:
        return {c: "" for c in cols}

    def calc_break_hhmm_for_shift(
        zero_shift: bool,
        wd: date,
        shift_start: Optional[str],
        shift_end: Optional[str],
        excel_break_mins: int,
    ) -> str:
        shift_mins = compute_shift_duration_minutes(wd, shift_start, shift_end)

        if zero_shift or shift_mins <= 0:
            return ""

        if shift_mins < 360:
            return ""

        if shift_mins == 540:
            return "01:00"

        if shift_mins > 540:
            desired_break = shift_mins - 480
            if desired_break >= shift_mins:
                return ""
            return minutes_to_hhmm(desired_break)

        if excel_break_mins <= 0:
            return ""
        eff = min(excel_break_mins, shift_mins)
        if eff >= shift_mins:
            return ""
        return minutes_to_hhmm(eff)

    for staff_code, rows in staff_rows.items():
        for rec in rows:
            wd: date = rec["work_date"]

            holiday_based = normalize_holiday_kbn(rec.get("holiday_kbn"))
            is_paid_leave = (safe_str(rec.get("kintai_status")) == "有休")
            zero_shift = is_zero_zero_shift(rec.get("shift_start"), rec.get("shift_end"))

            row = new_row()
            row["従業員番号"] = staff_code
            row["freee人事労務での表示名（編集しても反映されません）"] = ""
            row["日付"] = wd.strftime("%Y-%m-%d")
            row["勤務パターンコード"] = ""
            row["夜勤日種別"] = ""

            if is_paid_leave:
                row["勤務日種別"] = "所定労働日"
            else:
                row["勤務日種別"] = holiday_based

            if is_paid_leave:
                row["出勤時刻"] = rec.get("shift_start") or ""
                row["退勤時刻"] = rec.get("shift_end") or ""
                row["休憩時間"] = calc_break_hhmm_for_shift(
                    zero_shift,
                    wd,
                    rec.get("shift_start"),
                    rec.get("shift_end"),
                    int(rec.get("break_mins", 0) or 0),
                )
            else:
                if row["勤務日種別"] == "所定労働日":
                    row["出勤時刻"] = rec.get("shift_start") or ""
                    row["退勤時刻"] = rec.get("shift_end") or ""
                    row["休憩時間"] = calc_break_hhmm_for_shift(
                        zero_shift,
                        wd,
                        rec.get("shift_start"),
                        rec.get("shift_end"),
                        int(rec.get("break_mins", 0) or 0),
                    )
                else:
                    row["出勤時刻"] = ""
                    row["退勤時刻"] = ""
                    row["休憩時間"] = ""

            out_rows.append(row)

    out = pd.DataFrame(out_rows, columns=cols)
    csv_bytes = df_to_csv_bytes_utf8sig(out)
    return out, csv_bytes


# =========================
# Streamlit cache helpers
# =========================
@st.cache_data(show_spinner=False)
def read_jobcan_excels_cached(files: List[Tuple[str, bytes]]) -> Tuple[int, int, Dict[str, List[Dict[str, Any]]]]:
    return read_jobcan_excels(files)


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="ジョブカン勤怠 → freeeスケジュールCSV生成", layout="wide")
st.title("ジョブカン summary Excel → freee取込用スケジュールCSV（生成のみ）")

st.caption("CSV列は固定（テンプレ不要）です。freee認証は不要です。")
st.caption("休日区分は '祝日, 公休' / '祝日, 法休' などでも休日扱いします。")
st.caption("シフトが 0:00/0:00 の場合は休憩を必ずブランクにします。")

uploaded_excels = st.file_uploader(
    "ジョブカン summary Excel（複数可）",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded_excels:
    st.info("Excelファイルをアップロードしてください。")
    st.stop()

if st.button("CSVを生成", type="primary"):
    try:
        files = [(f.name, f.getvalue()) for f in uploaded_excels]
        year, month, staff_rows = read_jobcan_excels_cached(files)
        st.success(f"[OK] Excel読込完了：{year}年{month}月 / 対象従業員（シート数ベース）={len(staff_rows)}名")

        schedule_df, schedule_bytes = build_schedule_csv_fixed(staff_rows)
        st.success(f"[OK] スケジュールCSV生成：rows={len(schedule_df)}")

        st.session_state["schedule_bytes"] = schedule_bytes
        st.session_state["schedule_df"] = schedule_df
        st.session_state["year"] = year
        st.session_state["month"] = month

    except Exception as e:
        for k in ("schedule_bytes", "schedule_df", "year", "month"):
            st.session_state.pop(k, None)
        st.exception(e)

if st.session_state.get("schedule_bytes"):
    st.divider()
    st.write(f"生成済み: {st.session_state['year']}年{st.session_state['month']}月")
    st.dataframe(st.session_state["schedule_df"], use_container_width=True, height=480)
    st.download_button(
        label="freee_schedule_import.csv をダウンロード",
        data=st.session_state["schedule_bytes"],
        file_name="freee_schedule_import.csv",
        mime="text/csv",
    )
