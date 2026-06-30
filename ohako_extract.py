# ohako_extract.py
# -*- coding: utf-8 -*-
"""おはこフーズ 土日祝割増 & 有給休暇 計算ツール"""

import io
import math
from datetime import date, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(page_title="おはこフーズ 勤怠データ計算", layout="wide")
st.title("おはこフーズ 勤怠データ抽出・支給額計算ツール")

# ── ファイルアップロード ──────────────────────────────────────────────────────
c1, c2 = st.columns(2)
with c1:
    uploaded = st.file_uploader("① 日次勤怠CSV", type="csv", key="csv_upload")
with c2:
    wage_uploaded = st.file_uploader("② 従業員基本給CSV（時給テンプレート）", type="csv", key="wage_csv")

if uploaded is None or wage_uploaded is None:
    missing = []
    if uploaded is None:
        missing.append("日次勤怠CSV")
    if wage_uploaded is None:
        missing.append("従業員基本給CSV")
    st.info(f"{'・'.join(missing)} をアップロードしてください。")
    st.stop()

# ── データ読み込み ────────────────────────────────────────────────────────────
df = pd.read_csv(uploaded, encoding="utf-8-sig")
df.columns = df.columns.str.strip()
df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
df["従業員番号"] = df["従業員番号"].astype(str).str.strip()

df_wage = pd.read_csv(wage_uploaded, encoding="utf-8-sig")
df_wage.columns = df_wage.columns.str.strip()
df_wage["従業員番号"] = df_wage["従業員番号"].astype(str).str.strip()

wage_cols = df_wage.columns.tolist()
weekend_col = next((c for c in wage_cols if "土日祝" in c), None)

sel = ["従業員番号", "基本給"]
if weekend_col:
    sel.append(weekend_col)

rename_map = {"基本給": "基本時給"}
if weekend_col:
    rename_map[weekend_col] = "土日祝時給"

dw = df_wage[sel].copy().rename(columns=rename_map)
for c in ["基本時給", "土日祝時給"]:
    dw[c] = pd.to_numeric(dw.get(c, 0), errors="coerce").fillna(0)

df = df.merge(dw, on="従業員番号", how="left")
for c in ["基本時給", "土日祝時給"]:
    if c not in df.columns:
        df[c] = 0
    df[c] = df[c].fillna(0)

# ── 祝日計算（jpholiday 不要・内閣府ルール準拠） ──────────────────────────────
def _nth_weekday(year: int, month: int, weekday: int, n: int) -> date:
    first = date(year, month, 1)
    offset = (weekday - first.weekday()) % 7
    if n > 1:
        offset += 7 * (n - 1)
    return first.replace(day=1 + offset)

def _vernal_equinox_day(year: int) -> int:
    if year < 1980 or year > 2099:
        return 20
    return int(20.8431 + 0.242194 * (year - 1980) - (year - 1980) // 4)

def _autumnal_equinox_day(year: int) -> int:
    if year < 1980 or year > 2099:
        return 23
    return int(23.2488 + 0.242194 * (year - 1980) - (year - 1980) // 4)

@st.cache_data(show_spinner=False)
def build_jp_holiday_set(year: int) -> set:
    out: set = set()
    out.add(date(year, 1, 1))
    out.add(date(year, 2, 11))
    out.add(date(year, 2, 23))
    out.add(date(year, 4, 29))
    out.add(date(year, 5, 3))
    out.add(date(year, 5, 4))
    out.add(date(year, 5, 5))
    out.add(date(year, 11, 3))
    out.add(date(year, 11, 23))
    out.add(date(year, 8, 10 if year == 2020 else 11))
    out.add(_nth_weekday(year, 1, 0, 2))
    out.add(_nth_weekday(year, 7, 0, 3))
    out.add(_nth_weekday(year, 9, 0, 3))
    out.add(_nth_weekday(year, 10, 0, 2))
    out.add(date(year, 3, _vernal_equinox_day(year)))
    out.add(date(year, 9, _autumnal_equinox_day(year)))
    # 振替休日：日曜の祝日 → 祝日でも日曜でもない最初の日（連休で押し出しも対応）
    for d in sorted(list(out)):
        if d.weekday() == 6:  # 日曜
            candidate = d + timedelta(days=1)
            while candidate in out or candidate.weekday() == 6:
                candidate += timedelta(days=1)
            out.add(candidate)
    return out

_HOLIDAY_CACHE: dict = {}

def is_weekend_or_holiday(dt) -> bool:
    if pd.isna(dt):
        return False
    d = dt.date() if hasattr(dt, "date") else dt
    if d.weekday() >= 5:
        return True
    year = d.year
    if year not in _HOLIDAY_CACHE:
        _HOLIDAY_CACHE[year] = build_jp_holiday_set(year)
    return d in _HOLIDAY_CACHE[year]

# ── ヘルパー関数 ──────────────────────────────────────────────────────────────
def parse_hhmm(val):
    if pd.isna(val) or str(val).strip() == "":
        return 0
    parts = str(val).strip().split(":")
    if len(parts) < 2:
        return 0
    try:
        return int(parts[0]) * 60 + int(parts[1])
    except ValueError:
        return 0

def parse_col(series):
    return series.apply(parse_hhmm)

def hhmm(m):
    m = int(m)
    return f"{m // 60:02d}:{m % 60:02d}"

def yen(v):
    return f"¥{int(round(v)):,}"

def ceil_pay(rate, mult, minutes):
    return math.ceil(rate * mult * minutes / 60)

# ── 共通設定 ──────────────────────────────────────────────────────────────────
OHAKO_CATS = [
    ("法定内残業", "法定内残業時間",  1.00),
    ("時間外",    "時間外労働時間",  1.25),
    ("法定休日",  "法定休日労働時間", 1.35),
    ("深夜",     "深夜労働時間",    0.25),
]

# 対応要のハイライト色
HIGHLIGHT_COLOR = "background-color: #FFEEBA"


def render_ohako_section(df_all, dept_keyword, header_num, sub_num, store_label):
    """
    土日祝割増 & 有給休暇を表示する（おはこフーズ各店舗用）。
    Excel 出力用の数値 DataFrame を返す。
    戻り値: {"premium": DataFrame or None, "paid_leave": DataFrame or None}
    """
    st.header(f"{header_num} おはこフーズ {store_label}")
    st.caption(
        "**計算式**: 全日時間×基本時給×割増率（基本） ＋ 土日祝時間×差額時給×割増率（追加）　"
        "／ 割増率: 法定内 ×1.00 / 時間外 ×1.25 / 法定休日 ×1.35 / 深夜 ×0.25　"
        "／ 🟡 黄色ハイライト = 差額追加が発生しており手作業対応が必要な従業員"
    )

    result = {"premium": None, "paid_leave": None}

    # 土日祝時給が設定されている従業員のみ対象（月給制社員を除外）
    mask = df_all["部門"].str.contains(dept_keyword, na=False) & (df_all["土日祝時給"] > 0)
    df_s = df_all[mask].copy()
    if df_s.empty:
        st.info(f"{store_label} の対象データがありません。")
        st.divider()
        return result

    df_s["土日祝"] = df_s["日付"].apply(is_weekend_or_holiday)
    df_ot = df_s[df_s["土日祝"]].copy()

    for label, col, _ in OHAKO_CATS:
        df_s[f"_A_{label}"]  = parse_col(df_s[col])
        df_ot[f"_W_{label}"] = parse_col(df_ot[col])

    df_ot["_残業深夜_分"] = sum(df_ot[f"_W_{label}"] for label, _, _ in OHAKO_CATS)
    df_show = df_ot[df_ot["_残業深夜_分"] > 0].copy()

    # ── A. 土日祝残業・深夜 ──────────────────────────────────────────────────────
    st.subheader(f"{sub_num}-A. 土日祝の残業・深夜時間（日別明細）")

    if df_show.empty:
        st.info("土日祝の残業・深夜レコードはありません。")
    else:
        disp = pd.DataFrame()
        disp["日付"]       = df_show["日付"].dt.strftime("%Y/%m/%d")
        disp["曜日"]       = df_show["曜日"]
        disp["氏名"]       = df_show["氏名"]
        disp["基本時給"]   = df_show["基本時給"].apply(lambda v: f"¥{int(v):,}")
        disp["土日祝時給"] = df_show["土日祝時給"].apply(lambda v: f"¥{int(v):,}")
        for label, _, _ in OHAKO_CATS:
            disp[f"{label}(h)"] = df_show[f"_W_{label}"].apply(hhmm)
        st.dataframe(disp.reset_index(drop=True), use_container_width=True)

        # ── 従業員別月次集計 ──────────────────────────────────────────────────
        st.subheader("従業員別月次集計・支給額内訳")
        st.caption("🟡 差額追加（手作業対応）が発生している従業員を黄色でハイライトしています")

        a_cols = [f"_A_{label}" for label, _, _ in OHAKO_CATS]
        w_cols = [f"_W_{label}" for label, _, _ in OHAKO_CATS]

        grp_all = df_s.groupby(["従業員番号", "氏名", "基本時給", "土日祝時給"])[a_cols].sum().reset_index()
        grp_wk  = df_show.groupby(["従業員番号", "氏名", "基本時給", "土日祝時給"])[w_cols].sum().reset_index()

        grp = grp_wk.merge(grp_all, on=["従業員番号", "氏名", "基本時給", "土日祝時給"], how="left")
        for c in a_cols:
            grp[c] = grp[c].fillna(0).astype(int)
        grp = grp.sort_values("従業員番号").reset_index(drop=True)

        g = grp.copy()
        g["差額時給"] = g["土日祝時給"] - g["基本時給"]

        total_base  = pd.Series(0, index=g.index, dtype=int)
        total_extra = pd.Series(0, index=g.index, dtype=int)

        # 表示用（書式付き文字列）
        out = g[["従業員番号", "氏名", "基本時給", "土日祝時給", "差額時給"]].copy()
        out["基本時給"]   = g["基本時給"].apply(lambda v: f"¥{int(v):,}")
        out["土日祝時給"] = g["土日祝時給"].apply(lambda v: f"¥{int(v):,}")
        out["差額時給"]   = g["差額時給"].apply(lambda v: f"¥{int(v):,}")

        # Excel 用（数値）
        xl = g[["従業員番号", "氏名", "基本時給", "土日祝時給", "差額時給"]].copy()

        for label, _, mult in OHAKO_CATS:
            a_col    = f"_A_{label}"
            w_col    = f"_W_{label}"
            mult_str = f"×{mult:.2f}".rstrip("0").rstrip(".")

            base_pay  = g.apply(lambda r, ac=a_col, m=mult: ceil_pay(r["基本時給"],  m, r[ac]), axis=1)
            extra_pay = g.apply(lambda r, wc=w_col, m=mult: ceil_pay(r["差額時給"],  m, r[wc]), axis=1)
            total_base  += base_pay
            total_extra += extra_pay

            out[f"{label}({mult_str})_全日(h)"] = g[a_col].apply(hhmm)
            out[f"{label}_土日祝(h)"]           = g[w_col].apply(hhmm)
            out[f"{label}_全日基本(円)"]        = base_pay.apply(yen)
            out[f"{label}_追加(円)"]            = extra_pay.apply(yen)
            out[f"{label}_小計(円)"]            = (base_pay + extra_pay).apply(yen)

            xl[f"{label}_全日(h:mm)"]    = g[a_col].apply(hhmm)
            xl[f"{label}_土日祝(h:mm)"]  = g[w_col].apply(hhmm)
            xl[f"{label}_全日基本(円)"] = base_pay
            xl[f"{label}_追加(円)"]    = extra_pay
            xl[f"{label}_小計(円)"]    = base_pay + extra_pay

        out["基本合計(円)"] = total_base.apply(yen)
        out["差額追加(円)"] = total_extra.apply(yen)
        out["支給合計(円)"] = (total_base + total_extra).apply(yen)

        xl["基本合計(円)"] = total_base
        xl["差額追加(円)"] = total_extra
        xl["支給合計(円)"] = total_base + total_extra
        xl["対応要否"]     = total_extra.apply(lambda x: "要" if x > 0 else "")

        result["premium"] = xl

        # ── 行ハイライト（対応要 = total_extra > 0） ──────────────────────────
        needs_attention = (total_extra > 0).reset_index(drop=True)
        out_disp = out.reset_index(drop=True)

        def _highlight(row):
            bg = HIGHLIGHT_COLOR if needs_attention.iloc[row.name] else ""
            return [bg] * len(row)

        st.dataframe(out_disp.style.apply(_highlight, axis=1), use_container_width=True)

        m1, m2, m3 = st.columns(3)
        m1.metric("基本給分 合計（参照）", yen(total_base.sum()))
        m2.metric("差額追加分 合計",       yen(total_extra.sum()))
        m3.metric("支給 合計",             yen((total_base + total_extra).sum()))

    st.divider()

    # ── B. 有給休暇 ──────────────────────────────────────────────────────────────
    st.subheader(f"{sub_num}-B. 有給休暇時間（勤怠種別＝有休）")

    mask_dept = df_all["部門"].str.contains(dept_keyword, na=False)
    df_yu = df_all[mask_dept & (df_all["勤怠種別"] == "有休")].copy()
    df_yu = df_yu[df_yu["日付"].apply(is_weekend_or_holiday)].copy()

    if df_yu.empty:
        st.info("有休レコードはありません。")
    else:
        yu_cols = [c for c in ["日付", "曜日", "氏名",
                                "有休時間 - 合計", "有休時間 - 半休", "有休時間 - 時間休"]
                   if c in df_yu.columns]
        st.dataframe(df_yu[yu_cols].reset_index(drop=True), use_container_width=True)
        total_yu = parse_col(df_yu["有休時間 - 合計"]).sum()
        st.markdown(f"**有休時間 合計: `{hhmm(total_yu)}`**")

        st.subheader("従業員別 有給休暇合計")
        gy = (df_yu.groupby(["従業員番号", "氏名"])["有休時間 - 合計"]
              .apply(lambda s: parse_col(s).sum())
              .reset_index())
        gy.columns = ["従業員番号", "氏名", "有休(分)"]
        gy = gy.sort_values("従業員番号").reset_index(drop=True)
        gy["有休合計"] = gy["有休(分)"].apply(hhmm)
        st.dataframe(gy[["従業員番号", "氏名", "有休合計"]], use_container_width=True)

        result["paid_leave"] = gy[["従業員番号", "氏名", "有休合計"]].copy()

    st.divider()
    return result


# ════════════════════════════════════════════════════════════════════════════════
# ① 大雪通本店  ② 旭神店
# ════════════════════════════════════════════════════════════════════════════════
STORES = [
    ("大雪通本店", "①", "1"),
    ("旭神店",    "②", "2"),
]

all_results = {}
for store_label, hnum, snum in STORES:
    all_results[store_label] = render_ohako_section(df, store_label, hnum, snum, store_label)

# ════════════════════════════════════════════════════════════════════════════════
# Excel ダウンロード（店舗ごとにシート分け）
# ════════════════════════════════════════════════════════════════════════════════
st.header("📥 Excel ダウンロード")
st.caption("店舗ごとにシートを分けた 1 ファイルをダウンロードできます。")

buf = io.BytesIO()
has_data = False

with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    for store_label, res in all_results.items():
        premium    = res.get("premium")
        paid_leave = res.get("paid_leave")

        if premium is None and paid_leave is None:
            continue

        has_data = True
        row_cursor = 0

        if premium is not None:
            # ヘッダー行（タイトル）
            title_df = pd.DataFrame([[f"【土日祝割増計算】{store_label}"]])
            title_df.to_excel(writer, sheet_name=store_label,
                              index=False, header=False, startrow=row_cursor)
            row_cursor += 1
            premium.to_excel(writer, sheet_name=store_label,
                             index=False, startrow=row_cursor)
            row_cursor += len(premium) + 2  # データ行 + ヘッダー1 + 空白1

        if paid_leave is not None:
            title_df2 = pd.DataFrame([[f"【有給休暇合計】{store_label}"]])
            title_df2.to_excel(writer, sheet_name=store_label,
                               index=False, header=False, startrow=row_cursor)
            row_cursor += 1
            paid_leave.to_excel(writer, sheet_name=store_label,
                                index=False, startrow=row_cursor)

if has_data:
    buf.seek(0)
    st.download_button(
        label="📥 全店舗 Excel をダウンロード",
        data=buf.getvalue(),
        file_name="おはこフーズ_勤怠計算.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("ダウンロードできるデータがありません。")
