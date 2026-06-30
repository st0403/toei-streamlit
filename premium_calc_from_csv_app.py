import math
import streamlit as st
import pandas as pd
import jpholiday

st.set_page_config(page_title="勤怠データ抽出・支給額計算", layout="wide")
st.title("勤怠データ抽出・支給額計算ツール")

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
early_col   = next((c for c in wage_cols if "早朝" in c), None)
weekend_col = next((c for c in wage_cols if "土日祝" in c), None)

sel = ["従業員番号", "基本給"]
if early_col:   sel.append(early_col)
if weekend_col: sel.append(weekend_col)

rename_map = {"基本給": "基本時給"}
if early_col:   rename_map[early_col]   = "早朝時給"
if weekend_col: rename_map[weekend_col] = "土日祝時給"

dw = df_wage[sel].copy().rename(columns=rename_map)
for c in ["基本時給", "早朝時給", "土日祝時給"]:
    dw[c] = pd.to_numeric(dw.get(c, 0), errors="coerce").fillna(0)

df = df.merge(dw, on="従業員番号", how="left")
for c in ["基本時給", "早朝時給", "土日祝時給"]:
    if c not in df.columns:
        df[c] = 0
    df[c] = df[c].fillna(0)

# ── ヘルパー関数 ──────────────────────────────────────────────────────────────
def is_weekend_or_holiday(dt):
    if pd.isna(dt):
        return False
    return dt.weekday() >= 5 or jpholiday.is_holiday(dt.date())

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

def calc_early_min(row):
    ws, we = 360, 540
    s = parse_hhmm(row.get("出勤時刻", ""))
    e = parse_hhmm(row.get("退勤時刻", ""))
    if s == 0 or e == 0:
        return 0
    gross = max(0, min(e, we) - max(s, ws))
    if gross == 0:
        return 0
    for bi, bo in [("休憩1入り時刻","休憩1戻り時刻"),
                   ("休憩2入り時刻","休憩2戻り時刻"),
                   ("休憩3入り時刻","休憩3戻り時刻")]:
        bs = parse_hhmm(row.get(bi, ""))
        be = parse_hhmm(row.get(bo, ""))
        if bs and be:
            gross -= max(0, min(be, we) - max(bs, ws))
    return max(0, gross)

def calc_early_overtime_min(row):
    WINDOW_END = 540
    sched_end  = parse_hhmm(row.get("勤務予定退勤時刻", ""))
    actual_end = parse_hhmm(row.get("退勤時刻", ""))
    if sched_end <= 0 or sched_end > WINDOW_END:
        return 0
    if actual_end <= sched_end:
        return 0
    return max(0, min(actual_end, WINDOW_END) - sched_end)

# ── 支給額計算ユーティリティ ─────────────────────────────────────────────────
# 割増率
#   所定内・法定内残業: ×1.00（差額のみ、追加割増なし）
#   時間外:           ×1.25
#   法定休日:         ×1.35
#   深夜:             ×0.25（25%分の追加のみ）
#
# 計算式:
#   基本給分  = 基本時給   × 割増率 × 月合計時間  （給与ソフト既算・参照用）
#   差額追加分 = 差額時給  × 割増率 × 月合計時間  （今回追加で支給する分）
#   差額時給  = 特別時給 − 基本時給
#   合計      = 基本給分 + 差額追加分

def ceil_pay(rate, mult, minutes):
    """分数を時間に変換し、切り上げて金額を返す（1円未満切り上げ）。"""
    return math.ceil(rate * mult * minutes / 60)

def make_pay_table(grp_df, rate_col, base_col, items, base_only_items=None):
    """
    grp_df          : 従業員別月次集計済み DataFrame
    rate_col        : 特別時給カラム名 ("土日祝時給" or "早朝時給")
    base_col        : 基本時給カラム名
    items           : [(label, min_col, mult), ...]  ← 差額計算あり
    base_only_items : [(label, min_col, mult), ...]  ← 基本時給のみ（差額なし）

    計算方式（items）:
      基本  = ceil( 基本時給  × 割増率 × 分/60 )
      差額  = ceil( 差額時給  × 割増率 × 分/60 )   差額時給 = 特別時給 − 基本時給
      小計  = 基本 + 差額                           各々切り上げてから合算

    計算方式（base_only_items）:
      基本  = ceil( 基本時給  × 割増率 × 分/60 )
      差額  = 0
      小計  = 基本
    """
    g = grp_df.copy()
    g["差額時給"] = g[rate_col] - g[base_col]

    total_base  = pd.Series(0, index=g.index, dtype=int)
    total_extra = pd.Series(0, index=g.index, dtype=int)

    rows = []

    # ── 差額計算あり ──
    for label, min_col, mult in items:
        if min_col not in g.columns:
            continue
        base_pay  = g.apply(lambda r: ceil_pay(r[base_col],  mult, r[min_col]), axis=1)
        extra_pay = g.apply(lambda r: ceil_pay(r["差額時給"], mult, r[min_col]), axis=1)
        total_base  += base_pay
        total_extra += extra_pay
        rows.append((label, min_col, mult, base_pay, extra_pay, True))

    # ── 基本時給のみ（差額なし）──
    for label, min_col, mult in (base_only_items or []):
        if min_col not in g.columns:
            continue
        base_pay  = g.apply(lambda r: ceil_pay(r[base_col], mult, r[min_col]), axis=1)
        extra_pay = pd.Series(0, index=g.index, dtype=int)
        total_base += base_pay
        rows.append((label, min_col, mult, base_pay, extra_pay, False))

    id_cols = ["従業員番号"] if "従業員番号" in g.columns else []
    out = g[id_cols + ["氏名", base_col, rate_col, "差額時給"]].copy()
    out[base_col]   = g[base_col].apply(lambda v: f"¥{int(v):,}")
    out[rate_col]   = g[rate_col].apply(lambda v: f"¥{int(v):,}")
    out["差額時給"] = g["差額時給"].apply(lambda v: f"¥{int(v):,}")

    for label, min_col, mult, base_pay, extra_pay, has_extra in rows:
        mult_str = f"×{mult:.2f}".rstrip("0").rstrip(".")
        out[f"{label}({mult_str})(h)"] = g[min_col].apply(hhmm)
        out[f"{label}_基本(円)"]       = base_pay.apply(yen)
        if has_extra:
            out[f"{label}_差額(円)"]   = extra_pay.apply(yen)
        out[f"{label}_小計(円)"]       = (base_pay + extra_pay).apply(yen)

    out["基本給分_合計(円)"]  = total_base.apply(yen)
    out["差額追加分_合計(円)"] = total_extra.apply(yen)
    out["支給_合計(円)"]      = (total_base + total_extra).apply(yen)

    out["_基本計"] = total_base
    out["_差額計"] = total_extra
    out["_合計計"] = total_base + total_extra

    return out

PART_WAGE = "コメダ珈琲店　パート2025.12.1"

# ════════════════════════════════════════════════════════════════════════════════
# ① 千歳北信濃店
# ════════════════════════════════════════════════════════════════════════════════
st.header("① 千歳北信濃店　コメダ珈琲店 パート2025.12.1")
st.caption(
    "**計算式**: 全日時間×基本時給×割増率（基本） ＋ 土日祝時間×差額時給×割増率（追加）　"
    "／ 割増率: 法定内 ×1.00 / 時間外 ×1.25 / 法定休日 ×1.35 / 深夜 ×0.25"
)

CHITOSE_DEPT = "コメダ珈琲店　千歳北信濃店"
mask_c = df["部門"].str.contains(CHITOSE_DEPT, na=False) & (df["勤務・賃金"] == PART_WAGE)
df_c = df[mask_c].copy()
df_c["土日祝"] = df_c["日付"].apply(is_weekend_or_holiday)
df_ot = df_c[df_c["土日祝"]].copy()

# 対象カテゴリ（label, 元列名, 割増率）
CHITOSE_CATS = [
    ("法定内残業", "法定内残業時間",   1.00),
    ("時間外",    "時間外労働時間",   1.25),
    ("法定休日",  "法定休日労働時間",  1.35),
    ("深夜",     "深夜労働時間",     0.25),
]

# 全日・土日祝それぞれの分換算列を付与
for label, col, mult in CHITOSE_CATS:
    df_c[f"_A_{label}"] = parse_col(df_c[col])   # 全日（A = All）
    df_ot[f"_W_{label}"] = parse_col(df_ot[col]) # 土日祝（W = Weekend）

# 土日祝のうち残業・深夜のある行を日別表示対象とする
df_ot["_残業深夜_分"] = sum(df_ot[f"_W_{label}"] for label, _, _ in CHITOSE_CATS)
df_show = df_ot[df_ot["_残業深夜_分"] > 0].copy()

# ── 1-A. 土日祝の残業・深夜 ──────────────────────────────────────────────────
st.subheader("1-A. 土日祝の残業・深夜時間（日別明細）")

if df_show.empty:
    st.info("土日祝の残業・深夜レコードはありません。")
else:
    disp = pd.DataFrame()
    disp["日付"]       = df_show["日付"].dt.strftime("%Y/%m/%d")
    disp["曜日"]       = df_show["曜日"]
    disp["氏名"]       = df_show["氏名"]
    disp["基本時給"]   = df_show["基本時給"].apply(lambda v: f"¥{int(v):,}")
    disp["土日祝時給"] = df_show["土日祝時給"].apply(lambda v: f"¥{int(v):,}")
    for label, col, mult in CHITOSE_CATS:
        disp[f"{label}(h)"] = df_show[f"_W_{label}"].apply(hhmm)
    st.dataframe(disp.reset_index(drop=True), use_container_width=True)

    # ── 従業員別月次集計・支給額 ──────────────────────────────────────────────
    st.subheader("従業員別月次集計・支給額内訳")

    # 全日集計（df_c）
    a_cols = [f"_A_{label}" for label, _, _ in CHITOSE_CATS]
    grp_all = df_c.groupby(["従業員番号", "氏名", "基本時給", "土日祝時給"])[a_cols].sum().reset_index()

    # 土日祝残業ありの行のみで集計（df_show）
    w_cols = [f"_W_{label}" for label, _, _ in CHITOSE_CATS]
    grp_wk  = df_show.groupby(["従業員番号", "氏名", "基本時給", "土日祝時給"])[w_cols].sum().reset_index()

    # 土日祝残業ありの従業員のみを対象に全日時間をマージ・従業員番号順にソート
    grp = grp_wk.merge(grp_all, on=["従業員番号", "氏名", "基本時給", "土日祝時給"], how="left")
    for c in a_cols:
        grp[c] = grp[c].fillna(0).astype(int)
    grp = grp.sort_values("従業員番号").reset_index(drop=True)

    # 支給額計算
    g = grp.copy()
    g["差額時給"] = g["土日祝時給"] - g["基本時給"]

    total_base  = pd.Series(0, index=g.index, dtype=int)
    total_extra = pd.Series(0, index=g.index, dtype=int)

    out = g[["従業員番号", "氏名", "基本時給", "土日祝時給", "差額時給"]].copy()
    out["基本時給"]   = g["基本時給"].apply(lambda v: f"¥{int(v):,}")
    out["土日祝時給"] = g["土日祝時給"].apply(lambda v: f"¥{int(v):,}")
    out["差額時給"]   = g["差額時給"].apply(lambda v: f"¥{int(v):,}")

    for label, col, mult in CHITOSE_CATS:
        a_col = f"_A_{label}"
        w_col = f"_W_{label}"
        mult_str = f"×{mult:.2f}".rstrip("0").rstrip(".")

        base_pay  = g.apply(
            lambda r, ac=a_col, m=mult: ceil_pay(r["基本時給"],  m, r[ac]), axis=1)
        extra_pay = g.apply(
            lambda r, wc=w_col, m=mult: ceil_pay(r["差額時給"],  m, r[wc]), axis=1)
        total_base  += base_pay
        total_extra += extra_pay

        out[f"{label}({mult_str})_全日(h)"]  = g[a_col].apply(hhmm)
        out[f"{label}_土日祝(h)"]            = g[w_col].apply(hhmm)
        out[f"{label}_全日基本(円)"]         = base_pay.apply(yen)
        out[f"{label}_追加(円)"]             = extra_pay.apply(yen)
        out[f"{label}_小計(円)"]             = (base_pay + extra_pay).apply(yen)

    out["基本合計(円)"]  = total_base.apply(yen)
    out["差額追加(円)"]  = total_extra.apply(yen)
    out["支給合計(円)"]  = (total_base + total_extra).apply(yen)

    st.dataframe(out.reset_index(drop=True), use_container_width=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("基本給分 合計（参照）", yen(total_base.sum()))
    m2.metric("差額追加分 合計", yen(total_extra.sum()))
    m3.metric("支給 合計", yen((total_base + total_extra).sum()))

st.divider()

# ── 1-B. 有給休暇 ────────────────────────────────────────────────────────────
st.subheader("1-B. 有給休暇時間（勤怠種別＝有休）")

df_yu = df_c[df_c["勤怠種別"] == "有休"].copy()

if df_yu.empty:
    st.info("有休レコードはありません。")
else:
    yu_cols = [c for c in ["日付","曜日","氏名","有休時間 - 合計","有休時間 - 半休","有休時間 - 時間休"]
               if c in df_yu.columns]
    st.dataframe(df_yu[yu_cols].reset_index(drop=True), use_container_width=True)
    total_yu = parse_col(df_yu["有休時間 - 合計"]).sum()
    st.markdown(f"**有休時間 合計: `{hhmm(total_yu)}`**")

    st.subheader("従業員別 有給休暇合計")
    gy = (df_yu.groupby(["従業員番号", "氏名"])["有休時間 - 合計"]
          .apply(lambda s: parse_col(s).sum()).reset_index())
    gy.columns = ["従業員番号", "氏名", "有休(分)"]
    gy = gy.sort_values("従業員番号").reset_index(drop=True)
    gy["有休合計"] = gy["有休(分)"].apply(hhmm)
    st.dataframe(gy[["従業員番号", "氏名", "有休合計"]], use_container_width=True)


# ════════════════════════════════════════════════════════════════════════════════
# ② 旭川買物公園通り店
# ════════════════════════════════════════════════════════════════════════════════
st.header("② 旭川買物公園通り店　コメダ珈琲店 パート2025.12.1")
st.caption(
    "**計算式**: 差額時給（早朝時給−基本時給）× 割増率 × 月合計時間 = 差額追加分　"
    "／ 早朝実働・法定内残業ともに ×1.00"
)

ASAHIKAWA_DEPT = "コメダ珈琲店　旭川買物公園通り店"
mask_a = df["部門"].str.contains(ASAHIKAWA_DEPT, na=False) & (df["勤務・賃金"] == PART_WAGE)
df_a = df[mask_a].copy()

df_a["早朝_分"]    = df_a.apply(calc_early_min, axis=1)
df_a["早朝残業_分"] = df_a.apply(calc_early_overtime_min, axis=1)

df_early = df_a[df_a["早朝_分"] > 0].copy()
df_ot_am = df_a[df_a["早朝残業_分"] > 0].copy()

# ── 早朝実働 ─────────────────────────────────────────────────────────────────
st.subheader("早朝実働時間（6:00〜9:00）")

if df_early.empty:
    st.info("早朝（6:00〜9:00）勤務のレコードはありません。")
else:
    disp2 = pd.DataFrame()
    disp2["日付"]        = df_early["日付"].dt.strftime("%Y/%m/%d")
    disp2["曜日"]        = df_early["曜日"]
    disp2["氏名"]        = df_early["氏名"]
    disp2["出勤時刻"]    = df_early["出勤時刻"]
    disp2["退勤時刻"]    = df_early["退勤時刻"]
    disp2["早朝時間(h)"] = df_early["早朝_分"].apply(hhmm)
    st.dataframe(disp2.reset_index(drop=True), use_container_width=True)

    st.subheader("従業員別月次集計・支給額内訳（早朝実働）")
    g2 = (df_early.groupby(["従業員番号", "氏名","基本時給","早朝時給"])[["早朝_分"]]
          .sum().reset_index().sort_values("従業員番号").reset_index(drop=True))
    pt2 = make_pay_table(g2, "早朝時給", "基本時給", [("早朝実働", "早朝_分", 1.00)])
    dcols2 = [c for c in pt2.columns if not c.startswith("_")]
    st.dataframe(pt2[dcols2].reset_index(drop=True), use_container_width=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("基本給分 合計（参照）", yen(pt2["_基本計"].sum()))
    m2.metric("差額追加分 合計", yen(pt2["_差額計"].sum()))
    m3.metric("支給 合計", yen(pt2["_合計計"].sum()))

st.divider()

# ── 早朝法定内残業 ────────────────────────────────────────────────────────────
st.subheader("早朝法定内残業（予定退勤≤9:00 で実退勤がそれを超えた分）")
st.caption("早朝法定内残業 = min(実退勤, 09:00) − 勤務予定退勤時刻　×1.00（早朝時給で差額計算）")

if df_ot_am.empty:
    st.info("早朝法定内残業のレコードはありません。")
else:
    disp3 = pd.DataFrame()
    disp3["日付"]         = df_ot_am["日付"].dt.strftime("%Y/%m/%d")
    disp3["曜日"]         = df_ot_am["曜日"]
    disp3["氏名"]         = df_ot_am["氏名"]
    disp3["勤務予定退勤"] = df_ot_am["勤務予定退勤時刻"]
    disp3["退勤時刻"]     = df_ot_am["退勤時刻"]
    disp3["早朝残業(h)"]  = df_ot_am["早朝残業_分"].apply(hhmm)
    st.dataframe(disp3.reset_index(drop=True), use_container_width=True)

    st.subheader("従業員別月次集計・支給額内訳（早朝法定内残業）")
    g3 = (df_ot_am.groupby(["従業員番号", "氏名","基本時給","早朝時給"])[["早朝残業_分"]]
          .sum().reset_index().sort_values("従業員番号").reset_index(drop=True))
    pt3 = make_pay_table(g3, "早朝時給", "基本時給", [("早朝残業", "早朝残業_分", 1.00)])
    dcols3 = [c for c in pt3.columns if not c.startswith("_")]
    st.dataframe(pt3[dcols3].reset_index(drop=True), use_container_width=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("基本給分 合計（参照）", yen(pt3["_基本計"].sum()))
    m2.metric("差額追加分 合計", yen(pt3["_差額計"].sum()))
    m3.metric("支給 合計", yen(pt3["_合計計"].sum()))
