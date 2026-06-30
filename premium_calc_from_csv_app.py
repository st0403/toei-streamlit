# app.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import math
import re
import unicodedata
from datetime import date
from typing import Optional

import pandas as pd
import streamlit as st

# =========================
# 設定
# =========================
STORE_NAME_DEFAULT = "コメダ珈琲店 千歳北信濃店"
PREMIUM_PER_HOUR = 100

RATES = {
    "所定内労働": 1.0,
    "法定内残業": 1.0,
    "時間外労働": 1.25,
    "法定休日労働": 1.35,
    "深夜労働": 0.25,
}

COL_EMP_NO = "従業員番号"
COL_NAME = "氏名"
COL_DATE = "日付"
COL_DEPT = "部門"
COL_CLOCK_IN = "出勤時刻"
COL_ATTENDANCE_TYPE = "勤怠種別"  # 有給休暇の判定に使用

TIME_COLS = {
    "所定内労働": "所定内労働時間",
    "法定内残業": "法定内残業時間",
    "時間外労働": "時間外労働時間",
    "法定休日労働": "法定休日労働時間",
    "深夜労働": "深夜労働時間",
}

# =========================
# 共通関数
# =========================
def normalize_text(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v)
    s = re.sub(r"[\u200B-\u200D\uFEFF]", "", s)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def hhmm_to_minutes(v) -> int:
    if pd.isna(v):
        return 0
    s = str(v)
    if ":" not in s:
        return 0
    h, m = s.split(":")
    return int(h) * 60 + int(m)

def safe_read_csv(upload) -> pd.DataFrame:
    raw = upload.getvalue()
    for enc in ("utf-8-sig", "cp932"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc)
        except:
            pass
    return pd.read_csv(io.BytesIO(raw), encoding="utf-8")

def _nth_weekday(year: int, month: int, weekday: int, n: int) -> date:
    """月内の第n番目のweekday（0=月）の日付を返す"""
    first = date(year, month, 1)
    # first の曜日を 0=月 に合わせる
    offset = (weekday - first.weekday()) % 7
    if n > 1:
        offset += 7 * (n - 1)
    return first.replace(day=1 + offset)


def _vernal_equinox_day(year: int) -> int:
    """春分の日の「日」（3月） 1980-2099 簡易式"""
    if year < 1980 or year > 2099:
        return 20
    return int(20.8431 + 0.242194 * (year - 1980) - (year - 1980) // 4)


def _autumnal_equinox_day(year: int) -> int:
    """秋分の日の「日」（9月） 1980-2099 簡易式"""
    if year < 1980 or year > 2099:
        return 23
    return int(23.2488 + 0.242194 * (year - 1980) - (year - 1980) // 4)


def _builtin_japanese_holidays(year: int) -> set[date]:
    """内閣府の国民の祝日ルールに基づき祝日を計算（jpholiday に依存しない）"""
    out: set[date] = set()
    # 固定日
    out.add(date(year, 1, 1))   # 元日
    out.add(date(year, 2, 11))  # 建国記念の日
    out.add(date(year, 2, 23))  # 天皇誕生日（2020〜）
    out.add(date(year, 4, 29))  # 昭和の日
    out.add(date(year, 5, 3))   # 憲法記念日
    out.add(date(year, 5, 4))   # みどりの日
    out.add(date(year, 5, 5))   # こどもの日
    out.add(date(year, 11, 3))  # 文化の日
    out.add(date(year, 11, 23)) # 勤労感謝の日
    # 山の日（2020年は8/10、それ以外は8/11）
    out.add(date(year, 8, 10 if year == 2020 else 11))
    # ハッピーマンデー等
    out.add(_nth_weekday(year, 1, 0, 2))   # 成人の日（1月第2月曜）
    out.add(_nth_weekday(year, 7, 0, 3))   # 海の日（7月第3月曜）
    out.add(_nth_weekday(year, 9, 0, 3))   # 敬老の日（9月第3月曜）
    out.add(_nth_weekday(year, 10, 0, 2))  # スポーツの日（10月第2月曜）
    # 春分の日・秋分の日
    out.add(date(year, 3, _vernal_equinox_day(year)))
    out.add(date(year, 9, _autumnal_equinox_day(year)))
    # 振替休日：日曜と重なった祝日の翌日（月曜）を追加
    for d in list(out):
        if d.weekday() == 6:  # 日曜
            next_mon = d.replace(day=d.day + 1)
            if next_mon not in out:
                out.add(next_mon)
    return out


@st.cache_data(show_spinner=False)
def build_jp_holiday_set(year: int) -> set[date]:
    """祝日集合。jpholiday があれば併用、なければ内閣府ルールの自前計算のみ"""
    builtin = _builtin_japanese_holidays(year)
    try:
        import jpholiday
        for m in range(1, 13):
            for d in range(1, 32):
                try:
                    dt = date(year, m, d)
                    if jpholiday.is_holiday(dt):
                        builtin.add(dt)
                except (ValueError, TypeError):
                    pass
    except Exception:
        pass
    return builtin


def parse_manual_holidays(text: str) -> list[date]:
    """カンマ/改行/空白区切りの日付文字列（YYYY-MM-DD）をパースして date のリストで返す"""
    if not text.strip():
        return []
    parts = re.split(r"[,\n\r\t ]+", text.strip())
    out: list[date] = []
    for p in parts:
        if not p:
            continue
        try:
            out.append(date.fromisoformat(p))
        except Exception:
            pass
    return out


# =========================
# 計算
# =========================
def compute(df, store_name, year, month, use_holiday, manual_holidays: set[date] | None = None):

    df = df.copy()
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")
    df = df[df[COL_DATE].notna()]

    df["_dept_norm"] = df[COL_DEPT].apply(normalize_text)
    store_norm = normalize_text(store_name)

    # ★ここが最大の修正点（完全一致 → 含む）
    df = df[df["_dept_norm"].str.contains(store_norm, na=False)]

    df["year"] = df[COL_DATE].dt.year
    df["month"] = df[COL_DATE].dt.month
    df = df[(df["year"]==year) & (df["month"]==month)]

    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    holidays = build_jp_holiday_set(year)
    if use_holiday and manual_holidays:
        holidays = holidays | set(manual_holidays)
    df["date_only"] = df[COL_DATE].dt.date
    df["is_holiday"] = df["date_only"].apply(
        lambda d: d.weekday()>=5 or (use_holiday and d in holidays)
    )

    df = df[df["is_holiday"]]

    for k, col in TIME_COLS.items():
        df[f"{k}_mins"] = df[col].apply(hhmm_to_minutes)

    # 有給（実労働0）の日は割増対象から除外
    df["_work_mins"] = df[[f"{k}_mins" for k in TIME_COLS]].sum(axis=1)
    df = df[df["_work_mins"] > 0].drop(columns=["_work_mins"])

    # 土日祝で有給休暇の日は割増対象から除外（勤怠種別で判定）
    if COL_ATTENDANCE_TYPE in df.columns:
        df["_attendance_norm"] = df[COL_ATTENDANCE_TYPE].apply(normalize_text)
        df = df[df["_attendance_norm"] != "有給休暇"].drop(columns=["_attendance_norm"])

    g = df.groupby([COL_EMP_NO, COL_NAME]).sum(numeric_only=True).reset_index()

    for k in TIME_COLS:
        rate = RATES[k]
        g[f"{k}加算額"] = g[f"{k}_mins"].apply(
            lambda m: math.ceil(m * PREMIUM_PER_HOUR * rate / 60)
        )

    add_cols = [f"{k}加算額" for k in TIME_COLS]
    g["合計"] = g[add_cols].sum(axis=1)

    return g, pd.DataFrame([{
        COL_EMP_NO:"合計",
        COL_NAME:store_name,
        "合計": int(g["合計"].sum())
    }])

# =========================
# UI
# =========================
st.set_page_config(layout="wide")
st.title("日次勤怠CSV → 土日祝+100円 自動計算")

up = st.file_uploader("日次勤怠CSV", type="csv")
use_holiday = st.checkbox("祝日も含める", value=True)

if not up:
    st.stop()

df = safe_read_csv(up)

dept_list = sorted(df[COL_DEPT].dropna().unique())
store = st.selectbox("対象部門（店舗）", dept_list)

ym = pd.to_datetime(df[COL_DATE], errors="coerce").dt.to_period("M").astype(str).dropna().unique().tolist()
ym_sel = st.selectbox("対象年月", sorted(ym))
y, m = ym_sel.split("-")
iy, im = int(y), int(m)

# 処理月の祝日表示・手入力で補完
manual_key = f"manual_holidays_{ym_sel}"
if manual_key not in st.session_state:
    st.session_state[manual_key] = ""
if manual_key + "_set" not in st.session_state:
    st.session_state[manual_key + "_set"] = set()

holidays_year = build_jp_holiday_set(iy)
month_holidays = sorted([d for d in holidays_year if d.year == iy and d.month == im])

with st.expander("処理月の祝日（表示・手入力で補完）", expanded=False):
    if month_holidays:
        hd_df = pd.DataFrame({
            "日付": [d.isoformat() for d in month_holidays],
            "曜日": [["月", "火", "水", "木", "金", "土", "日"][d.weekday()] for d in month_holidays],
        })
        st.dataframe(hd_df, use_container_width=True, height=240)
    else:
        st.info("この月は（自動判定上）祝日がありません。")

    text = st.text_area(
        "追加したい祝日（YYYY-MM-DD。カンマ/改行区切り）※当月のみ反映",
        value=st.session_state[manual_key],
        placeholder="例）2026-09-22\n2026-09-23",
        height=100,
    )
    col1, col2 = st.columns(2)
    with col1:
        if st.button("祝日を追加（当月のみ）", key=manual_key + "_add"):
            ds = parse_manual_holidays(text)
            valid = [d for d in ds if d.year == iy and d.month == im]
            invalid = len(ds) - len(valid)
            st.session_state[manual_key] = text
            st.session_state[manual_key + "_set"] = set(valid)
            if invalid:
                st.warning(f"当月以外 or 形式不正のため除外: {invalid}件")
            st.success(f"当月の追加祝日: {len(set(valid))}件を反映しました。")
    with col2:
        if st.button("当月の追加祝日をクリア", key=manual_key + "_clear"):
            st.session_state[manual_key] = ""
            st.session_state[manual_key + "_set"] = set()
            st.success("クリアしました。")

manual_holidays = st.session_state.get(manual_key + "_set", set())

# 漏れ検知ヒント（当月・CSV内の平日で休日候補に含まれていない日）
if use_holiday:
    try:
        dt_ser = pd.to_datetime(df[COL_DATE], errors="coerce")
        mask = (dt_ser.dt.year == iy) & (dt_ser.dt.month == im)
        month_dates = dt_ser.loc[mask].dt.date.unique()
        all_holidays = holidays_year | manual_holidays
        missed = sum(1 for d in month_dates if d.weekday() < 5 and d not in all_holidays)
        if missed > 0:
            st.caption(f"祝日漏れの可能性：{missed}日（平日で休日候補に含まれていない日がCSVにあります）")
    except Exception:
        pass

if st.button("計算する"):
    res, total = compute(df, store, iy, im, use_holiday, manual_holidays)

    if res.empty:
        st.error("対象データがありません（部門名 or 年月が一致していません）")
        st.stop()

    st.dataframe(res)
    st.dataframe(total)

    # 計算結果を1つのCSVにまとめてダウンロード（UTF-8 BOMでExcelでも文字化けしない）
    out = pd.concat([res, total], ignore_index=True)
    csv_bytes = out.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="計算結果をCSVでダウンロード",
        data=csv_bytes,
        file_name=f"土日祝加算額_{store}_{ym_sel}.csv",
        mime="text/csv",
    )
