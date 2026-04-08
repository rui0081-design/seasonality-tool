import io
import re
import time
import random
import logging
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from statsmodels.tsa.seasonal import seasonal_decompose
import streamlit as st
import plotly.graph_objects as go

warnings.filterwarnings("ignore")
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 120
plt.rcParams["savefig.dpi"] = 170

try:
    import japanize_matplotlib  # noqa: F401
except Exception:
    pass

st.set_page_config(page_title="Google Trends 季節性分析ツール", page_icon="📈", layout="wide")

GEO_MAP = {
    "日本 (JP)": "JP",
    "全世界": "",
    "アメリカ (US)": "US",
    "韓国 (KR)": "KR",
    "台湾 (TW)": "TW",
    "イギリス (GB)": "GB",
}

GPROP_MAP = {
    "通常のウェブ検索": "",
    "画像検索": "images",
    "ニュース検索": "news",
    "YouTube検索": "youtube",
    "Googleショッピング": "froogle",
}

TIMEFRAME_MAP = {
    "過去12か月": "today 12-m",
    "過去3年": "today 3-y",
    "過去5年": "today 5-y",
    "2004年以降すべて": "all",
}

MONTH_LABELS = {1: "1月", 2: "2月", 3: "3月", 4: "4月", 5: "5月", 6: "6月", 7: "7月", 8: "8月", 9: "9月", 10: "10月", 11: "11月", 12: "12月"}

st.markdown(
    """
    <style>
    .block-container {max-width: 1200px; padding-top: 1.4rem; padding-bottom: 2rem;}
    .hero {background: linear-gradient(135deg, #ffffff 0%, #f8fbff 100%); border: 1px solid #e5e7eb; border-radius: 24px; padding: 28px 30px; margin-bottom: 16px;}
    .hero h1 {margin: 0 0 8px 0; color: #0f172a; font-size: 2rem;}
    .hero p {margin: 0; color: #334155; line-height: 1.8;}
    .subtle {color: #64748b; font-size: 0.92rem;}
    .softbox {background: #ffffff; border: 1px solid #e5e7eb; border-radius: 18px; padding: 16px 18px;}
    .summary-card {background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 18px; padding: 16px 18px; height: 100%;}
    .summary-card h3 {margin: 0 0 10px 0; font-size: 1rem; color: #0f172a;}
    .summary-card p {margin: 0; color: #334155; line-height: 1.8;}
    .section-label {font-size: .82rem; font-weight: 700; color: #2563eb; letter-spacing: .05em; text-transform: uppercase;}
    .stButton button, .stDownloadButton button {border-radius: 12px; font-weight: 700; height: 2.85rem;}
    div[data-testid="stMetric"] {background: #ffffff; border: 1px solid #e5e7eb; padding: 14px 16px; border-radius: 16px;}
    </style>
    """,
    unsafe_allow_html=True,
)


def clean_name(text: str) -> str:
    text = re.sub(r'[\\/:*?"<>|]+', '_', str(text))
    return text[:80] if text else "seasonality"


def metric_guide_df() -> pd.DataFrame:
    return pd.DataFrame([
        ["value", "元の検索指数。大きいほど需要が強い。"],
        ["trend", "長期トレンド。短期ノイズをならした基調。"],
        ["seasonal_index", "季節指数。1.00が平均、1.10なら平均より約10%強い。"],
        ["adjusted", "季節要因を除いた値。実力ベースの動き。"],
        ["yoy", "元データの前年同月比。"],
        ["adjusted_yoy", "季節要因を除いた前年同月比。"],
    ], columns=["metric", "meaning"])


def how_to_read_df() -> pd.DataFrame:
    return pd.DataFrame([
        ["需要が強い月", "季節指数が高い月。プロモーション強化候補。"],
        ["需要が弱い月", "季節指数が低い月。効率重視の運用や種まき検討。"],
        ["直近の伸び", "adjusted_yoyがプラスなら、季節要因を除いても伸びている。"],
        ["季節要因か実力か", "valueとadjustedの差が大きいほど季節要因の影響が大きい。"],
    ], columns=["観点", "見方"])


def friendly_error_message(err: Exception) -> str:
    msg = str(err)
    if "429" in msg or "TooManyRequests" in msg:
        return "Google Trends側が混雑しています。少し時間を置いて再実行してください。"
    if "最低24か月" in msg:
        return msg
    if "No module named" in msg:
        return f"依存ライブラリの読み込みに失敗しました。requirements.txt を確認してください。詳細: {msg}"
    return f"処理中にエラーが発生しました。詳細: {msg}"


def infer_csv_columns(df: pd.DataFrame):
    date_candidates = []
    value_candidates = []
    for c in df.columns:
        if pd.to_datetime(df[c], errors="coerce").notna().mean() >= 0.6:
            date_candidates.append(c)
        if pd.to_numeric(df[c], errors="coerce").notna().mean() >= 0.6:
            value_candidates.append(c)
    if not date_candidates:
        raise ValueError("CSVから日付列を自動判定できませんでした。")
    if not value_candidates:
        raise ValueError("CSVから数値列を自動判定できませんでした。")
    return date_candidates[0], value_candidates[0]


def normalize_monthly_index(df: pd.DataFrame, date_col: str, value_col: str) -> pd.DataFrame:
    out = df[[date_col, value_col]].copy()
    out.columns = ["date", "value"]
    out["date"] = pd.to_datetime(out["date"], errors="coerce")
    out["value"] = pd.to_numeric(out["value"], errors="coerce")
    out = out.dropna().sort_values("date")
    if out.empty:
        raise ValueError("有効な日付列と数値列が見つかりませんでした。")
    out = out.set_index("date")
    if len(out) >= 52:
        out = out.resample("W").mean().resample("MS").mean()
    else:
        out = out.resample("MS").mean()
    out["value"] = out["value"].interpolate(limit_direction="both")
    return out


@st.cache_data(show_spinner=False, ttl=1800)
def fetch_google_trends_cached(keyword: str, geo="JP", timeframe="today 5-y", gprop="", cat=0, max_retries=4) -> pd.DataFrame:
    try:
        from pytrends.request import TrendReq
    except Exception as e:
        raise RuntimeError(f"pytrendsの読み込みに失敗しました: {e}")

    last_err = None
    for attempt in range(1, max_retries + 1):
        try:
            pytrends = TrendReq(
                hl="ja-JP",
                tz=540,
                timeout=(10, 25),
                requests_args={
                    "headers": {
                        "User-Agent": (
                            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                            "AppleWebKit/537.36 (KHTML, like Gecko) "
                            "Chrome/124.0.0.0 Safari/537.36"
                        )
                    }
                },
            )
            time.sleep(random.uniform(0.8, 1.4))
            pytrends.build_payload([keyword], cat=int(cat), timeframe=timeframe, geo=geo, gprop=gprop)
            df = pytrends.interest_over_time()
            if df is None or df.empty:
                raise ValueError("Google Trendsからデータを取得できませんでした。")
            if "isPartial" in df.columns:
                df = df.drop(columns=["isPartial"])
            return df.rename(columns={keyword: "value"}).reset_index()
        except Exception as e:
            last_err = e
            if attempt < max_retries:
                wait = 6 + attempt * 3
                time.sleep(wait)
            else:
                raise last_err
    raise last_err


def build_analysis_tables(ts: pd.DataFrame, keyword: str):
    if len(ts) < 24:
        raise ValueError("季節性分析には最低24か月分の月次データが必要です。")

    result = ts.copy()
    decomp = seasonal_decompose(result["value"], model="multiplicative", period=12, extrapolate_trend="freq")
    result["trend"] = decomp.trend
    result["seasonal"] = decomp.seasonal
    result["resid"] = decomp.resid
    result["adjusted"] = result["value"] / result["seasonal"]
    result["yoy"] = result["value"].pct_change(12)
    result["adjusted_yoy"] = result["adjusted"].pct_change(12)
    result["month_num"] = result.index.month
    result["month"] = result["month_num"].map(MONTH_LABELS)

    month_tbl = (
        result.groupby(["month_num", "month"], as_index=False)["seasonal"]
        .mean()
        .sort_values("month_num")
        .rename(columns={"seasonal": "seasonal_index"})
    )
    month_tbl["平均との差"] = month_tbl["seasonal_index"] - 1
    month_tbl["平均比(%)"] = month_tbl["平均との差"] * 100

    latest = result.iloc[-1]
    latest_prev = result.iloc[-13] if len(result) >= 13 else None
    strong_month = month_tbl.sort_values("seasonal_index", ascending=False).iloc[0]
    weak_month = month_tbl.sort_values("seasonal_index", ascending=True).iloc[0]

    summary = pd.DataFrame([
        ["keyword", keyword],
        ["analysis_period", f"{result.index.min():%Y-%m} ～ {result.index.max():%Y-%m}"],
        ["latest_month", f"{result.index.max():%Y-%m}"],
        ["latest_value", round(float(latest['value']), 2)],
        ["latest_adjusted", round(float(latest['adjusted']), 2)],
        ["latest_yoy", None if pd.isna(latest['yoy']) else round(float(latest['yoy']) * 100, 2)],
        ["latest_adjusted_yoy", None if pd.isna(latest['adjusted_yoy']) else round(float(latest['adjusted_yoy']) * 100, 2)],
        ["strongest_month", strong_month["month"]],
        ["weakest_month", weak_month["month"]],
        ["max_seasonal_index", round(float(strong_month['seasonal_index']), 3)],
        ["min_seasonal_index", round(float(weak_month['seasonal_index']), 3)],
    ], columns=["item", "value"])

    return summary, month_tbl[["month_num", "month", "seasonal_index", "平均との差", "平均比(%)"]], result


def make_summary_text(month_tbl: pd.DataFrame, result_df: pd.DataFrame) -> dict:
    strong = month_tbl.sort_values("seasonal_index", ascending=False).head(3)["month"].tolist()
    weak = month_tbl.sort_values("seasonal_index", ascending=True).head(3)["month"].tolist()
    latest = result_df.iloc[-1]
    adj_yoy = latest["adjusted_yoy"]
    raw_yoy = latest["yoy"]
    seasonal_gap = latest["value"] - latest["adjusted"]

    if pd.isna(adj_yoy):
        growth_text = "直近の季節調整後前年比は算出できません。前年同月比較に必要な期間が足りません。"
    elif adj_yoy > 0.05:
        growth_text = f"直近の季節調整後前年比は {adj_yoy:.1%} で、実力ベースでも明確に伸びています。"
    elif adj_yoy >= 0:
        growth_text = f"直近の季節調整後前年比は {adj_yoy:.1%} で、緩やかに伸びています。"
    else:
        growth_text = f"直近の季節調整後前年比は {adj_yoy:.1%} で、実力ベースでは弱含みです。"

    if abs(seasonal_gap) >= 5:
        factor_text = "直近値は季節要因の影響を比較的大きく受けています。見かけの上下だけで判断しない方がいいです。"
    else:
        factor_text = "直近値と季節調整値の差は小さめで、足元は実力変動の影響が大きいです。"

    promo_text = f"需要が強い {strong[0]} 前の1〜2か月で認知施策を厚くし、{strong[0]} 当月は刈り取りを強める設計が基本です。弱い {weak[0]} は効率重視運用か次ピークへの種まきに回すのが妥当です。"

    return {
        "strong_months": " / ".join(strong),
        "weak_months": " / ".join(weak),
        "growth_text": growth_text,
        "factor_text": factor_text,
        "promo_text": promo_text,
    }


def build_plotly_trend(result_df: pd.DataFrame, keyword: str):
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=result_df.index, y=result_df["value"], mode="lines", name="実績", line=dict(width=2.6, color="#2563EB")))
    fig.add_trace(go.Scatter(x=result_df.index, y=result_df["adjusted"], mode="lines", name="季節調整後", line=dict(width=2.4, color="#F59E0B")))
    fig.update_layout(
        title=f"{keyword}｜検索需要の推移",
        height=420,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=60, b=20),
        legend=dict(orientation="h", y=1.08, x=0),
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=True, gridcolor="#e5e7eb", zeroline=False),
    )
    return fig


def build_plotly_seasonality(month_tbl: pd.DataFrame, keyword: str):
    colors = ["#7C3AED" if v >= 1 else "#CBD5E1" for v in month_tbl["seasonal_index"]]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=month_tbl["month"], y=month_tbl["seasonal_index"], marker_color=colors, text=[f"{v:.2f}" for v in month_tbl["seasonal_index"]], textposition="outside", name="季節指数"))
    fig.add_hline(y=1, line_dash="dash", line_color="#ef4444")
    fig.update_layout(
        title=f"{keyword}｜月別季節指数",
        height=420,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=60, b=20),
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=True, gridcolor="#e5e7eb", zeroline=False),
    )
    return fig


def build_matplotlib_figs(result_df: pd.DataFrame, month_tbl: pd.DataFrame, keyword: str):
    fig1, ax1 = plt.subplots(figsize=(12.5, 5.2))
    ax1.plot(result_df.index, result_df["value"], linewidth=2.6, color="#2563EB", label="実績")
    ax1.plot(result_df.index, result_df["adjusted"], linewidth=2.2, color="#F59E0B", label="季節調整後")
    ax1.set_title(f"{keyword}｜検索需要の推移", loc="left", fontsize=15, fontweight="bold")
    ax1.spines[["top", "right"]].set_visible(False)
    ax1.grid(axis="y", linestyle="--", alpha=0.25)
    ax1.legend(frameon=False, ncol=2, loc="upper left")
    ax1.xaxis.set_major_locator(mdates.YearLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
    fig1.tight_layout()

    fig2, ax2 = plt.subplots(figsize=(12.5, 5.2))
    vals = month_tbl["seasonal_index"].values
    colors = ["#7C3AED" if v >= 1 else "#CBD5E1" for v in vals]
    bars = ax2.bar(month_tbl["month"], vals, color=colors, width=0.68)
    ax2.axhline(1, color="#ef4444", linestyle="--", linewidth=1.3)
    ax2.set_title(f"{keyword}｜月別季節指数", loc="left", fontsize=15, fontweight="bold")
    ax2.spines[["top", "right"]].set_visible(False)
    ax2.grid(axis="y", linestyle="--", alpha=0.18)
    for bar, v in zip(bars, vals):
        ax2.text(bar.get_x() + bar.get_width()/2, v + 0.015, f"{v:.2f}", ha="center", va="bottom", fontsize=9)
    fig2.tight_layout()
    return fig1, fig2


def fig_to_png_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, bbox_inches="tight", format="png")
    buf.seek(0)
    return buf.getvalue()


def csv_bytes(result_df: pd.DataFrame) -> bytes:
    export_df = result_df.copy().reset_index().rename(columns={"index": "date"})
    return export_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def read_uploaded_csv(uploaded_file):
    raw = uploaded_file.getvalue()
    for enc in ["utf-8-sig", "cp932", "utf-8", "shift_jis"]:
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc)
        except Exception:
            continue
    raise ValueError("CSVを読み込めませんでした。文字コードを確認してください。")


def make_excel_bytes(summary_df: pd.DataFrame, month_tbl: pd.DataFrame, result_df: pd.DataFrame, summary_texts: dict) -> bytes:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

    overview_text = pd.DataFrame([
        ["需要が強い月", summary_texts["strong_months"]],
        ["需要が弱い月", summary_texts["weak_months"]],
        ["直近評価", summary_texts["growth_text"]],
        ["季節要因か実力か", summary_texts["factor_text"]],
        ["施策示唆", summary_texts["promo_text"]],
    ], columns=["観点", "内容"])

    result_export = result_df.copy().reset_index().rename(columns={"index": "date"})
    xbuf = io.BytesIO()
    with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
        overview_text.to_excel(writer, sheet_name="summary", index=False)
        month_tbl.to_excel(writer, sheet_name="seasonality_by_month", index=False)
        result_export.to_excel(writer, sheet_name="result_detail", index=False)
        metric_guide_df().to_excel(writer, sheet_name="metric_guide", index=False)
        how_to_read_df().to_excel(writer, sheet_name="how_to_read", index=False)

    xbuf.seek(0)
    wb = load_workbook(xbuf)
    header_fill = PatternFill(fill_type="solid", fgColor="DCEAFE")
    border = Border(
        left=Side(style="thin", color="D1D5DB"), right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"), bottom=Side(style="thin", color="D1D5DB")
    )
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name="Yu Gothic UI", size=10)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = border
        for cell in ws[1]:
            cell.font = Font(name="Yu Gothic UI", size=10, bold=True)
            cell.fill = header_fill
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 44)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


st.markdown(
    """
    <div class="hero">
      <div class="section-label">Google Trends Seasonality Tool</div>
      <h1>Keywordを入れるだけで、需要の波と直近の実力がわかる</h1>
      <p>検索需要のトレンド・季節性・直近の実力を、社内共有しやすい形で可視化します。<br>Keyword / 地域 / 期間だけで始められます。</p>
    </div>
    """,
    unsafe_allow_html=True,
)

left, right = st.columns([1.2, 0.8], gap="large")

with left:
    st.markdown('<div class="softbox">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">Input</div>', unsafe_allow_html=True)
    keyword = st.text_input("Keyword", placeholder="例：自動車保険", label_visibility="visible")
    c1, c2, c3 = st.columns(3)
    with c1:
        geo_label = st.selectbox("地域", list(GEO_MAP.keys()), index=0)
    with c2:
        timeframe_label = st.selectbox("期間", list(TIMEFRAME_MAP.keys()), index=2)
    with c3:
        input_method = st.selectbox("データ取得", ["Google Trends", "CSVアップロード"], index=0)

    gprop_label = "通常のウェブ検索"
    cat = 0
    retry_count = 4
    uploaded_file = None

    with st.expander("詳細設定"):
        gprop_label = st.selectbox("検索種別", list(GPROP_MAP.keys()), index=0)
        cat = st.number_input("カテゴリID", min_value=0, step=1, value=0)
        retry_count = st.slider("再試行回数", min_value=2, max_value=6, value=4)
        if input_method == "CSVアップロード":
            uploaded_file = st.file_uploader("CSVファイル", type=["csv"])
        st.caption("通常はデフォルトのままで大丈夫です。")

    run = st.button("分析開始", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with right:
    st.markdown('<div class="summary-card"><h3>画面で見られるもの</h3><p>トレンドグラフ、月別季節指数、月別テーブル、直近結果テーブル、サマリーを表示します。</p></div>', unsafe_allow_html=True)
    st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)
    st.markdown('<div class="summary-card"><h3>ダウンロード</h3><p>Excel / CSV / PNG でそのまま共有できます。Excelは列幅・固定・見出し色まで整えています。</p></div>', unsafe_allow_html=True)
    st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)
    with st.expander("指標の見方"):
        st.dataframe(metric_guide_df(), use_container_width=True, hide_index=True)


if run:
    try:
        keyword_clean = keyword.strip()
        if not keyword_clean and input_method == "Google Trends":
            st.warning("Keywordを入力してください。")
            st.stop()

        with st.spinner("分析しています…"):
            if input_method == "Google Trends":
                raw_df = fetch_google_trends_cached(
                    keyword=keyword_clean,
                    geo=GEO_MAP[geo_label],
                    timeframe=TIMEFRAME_MAP[timeframe_label],
                    gprop=GPROP_MAP[gprop_label],
                    cat=cat,
                    max_retries=retry_count,
                )
                ts = normalize_monthly_index(raw_df, raw_df.columns[0], "value")
            else:
                if uploaded_file is None:
                    st.warning("CSVファイルをアップロードしてください。")
                    st.stop()
                raw_csv = read_uploaded_csv(uploaded_file)
                date_col, value_col = infer_csv_columns(raw_csv)
                ts = normalize_monthly_index(raw_csv, date_col, value_col)
                if not keyword_clean:
                    keyword_clean = Path(uploaded_file.name).stem

            summary_df, month_tbl, result_df = build_analysis_tables(ts, keyword_clean or "Keyword")
            summary_texts = make_summary_text(month_tbl, result_df)
            plotly_trend = build_plotly_trend(result_df, keyword_clean)
            plotly_seasonality = build_plotly_seasonality(month_tbl, keyword_clean)
            fig1, fig2 = build_matplotlib_figs(result_df, month_tbl, keyword_clean)

        c1, c2, c3, c4 = st.columns(4)
        latest_adj = result_df.iloc[-1]["adjusted"]
        latest_adj_yoy = result_df.iloc[-1]["adjusted_yoy"]
        with c1:
            st.metric("需要が強い月", summary_texts["strong_months"].split(" / ")[0])
        with c2:
            st.metric("需要が弱い月", summary_texts["weak_months"].split(" / ")[0])
        with c3:
            st.metric("直近の季節調整値", f"{latest_adj:.1f}")
        with c4:
            st.metric("季節調整後前年比", "-" if pd.isna(latest_adj_yoy) else f"{latest_adj_yoy:.1%}")

        s1, s2 = st.columns(2)
        with s1:
            st.markdown('<div class="summary-card"><h3>サマリー</h3><p>' + summary_texts["growth_text"] + '<br><br>' + summary_texts["factor_text"] + '</p></div>', unsafe_allow_html=True)
        with s2:
            st.markdown('<div class="summary-card"><h3>施策示唆</h3><p>強い月: ' + summary_texts["strong_months"] + '<br>弱い月: ' + summary_texts["weak_months"] + '<br><br>' + summary_texts["promo_text"] + '</p></div>', unsafe_allow_html=True)

        g1, g2 = st.columns(2)
        with g1:
            st.plotly_chart(plotly_trend, use_container_width=True)
        with g2:
            st.plotly_chart(plotly_seasonality, use_container_width=True)

        t1, t2 = st.columns(2)
        with t1:
            st.subheader("月別季節指数")
            st.dataframe(month_tbl, use_container_width=True, hide_index=True)
        with t2:
            st.subheader("直近結果")
            recent_df = result_df.copy().reset_index().rename(columns={"index": "date"}).tail(12)
            st.dataframe(recent_df[["date", "value", "adjusted", "yoy", "adjusted_yoy"]], use_container_width=True, hide_index=True)

        excel_data = make_excel_bytes(summary_df, month_tbl, result_df, summary_texts)
        csv_data = csv_bytes(result_df)
        png1_data = fig_to_png_bytes(fig1)
        png2_data = fig_to_png_bytes(fig2)
        safe = clean_name(keyword_clean or "keyword")

        st.subheader("ダウンロード")
        d1, d2, d3, d4 = st.columns(4)
        with d1:
            st.download_button("Excel", excel_data, file_name=f"{safe}_seasonality.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with d2:
            st.download_button("CSV", csv_data, file_name=f"{safe}_seasonality_result.csv", mime="text/csv", use_container_width=True)
        with d3:
            st.download_button("トレンドPNG", png1_data, file_name=f"{safe}_trend.png", mime="image/png", use_container_width=True)
        with d4:
            st.download_button("季節性PNG", png2_data, file_name=f"{safe}_seasonality.png", mime="image/png", use_container_width=True)

    except Exception as e:
        st.error(friendly_error_message(e))
        with st.expander("見直しポイント"):
            st.markdown(
                "- GitHub上の requirements.txt が最新か確認する\n"
                "- Streamlit Cloudで Clear cache を実行する\n"
                "- 同じ条件で短時間に連続実行しすぎない\n"
                "- 24か月未満のデータでは季節性分析できない\n"
                "- CSVは日付列と数値列が必要"
            )
