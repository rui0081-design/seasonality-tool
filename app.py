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
from pytrends.request import TrendReq
from pytrends import exceptions as pytrends_exceptions
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

warnings.filterwarnings("ignore")
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 120
plt.rcParams["savefig.dpi"] = 170

# Optional Japanese font support
try:
    import japanize_matplotlib  # noqa: F401
except Exception:
    for _font in ["IPAexGothic", "Noto Sans CJK JP", "Yu Gothic", "Meiryo", "DejaVu Sans"]:
        try:
            plt.rcParams["font.family"] = _font
            break
        except Exception:
            pass

st.set_page_config(
    page_title="Google Trends 季節性分解ツール",
    page_icon="📈",
    layout="wide",
)


def clean_name(text: str) -> str:
    text = re.sub(r'[\\/:*?"<>|]+', '_', str(text))
    return text[:80] if text else "seasonality"


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


st.markdown(
    """
    <style>
      .block-container {padding-top: 1.6rem; padding-bottom: 2.2rem; max-width: 1180px;}
      .hero {
        background: linear-gradient(135deg, #0B1F4D 0%, #102A6B 58%, #153A91 100%);
        padding: 28px 30px;
        border-radius: 22px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 12px 30px rgba(15, 23, 42, .18);
      }
      .hero h1 {font-size: 2.1rem; margin: 0 0 8px 0;}
      .hero p {margin: 0; font-size: 1rem; line-height: 1.75; opacity: .98;}
      .pillrow {display:flex; gap:10px; flex-wrap:wrap; margin-top:12px;}
      .pill {background: rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.18); border-radius:999px; padding:6px 12px; font-size:.9rem;}
      .smallcard {
        background:#F8FAFC; border:1px solid #E5E7EB; border-radius:16px; padding:14px 16px; height:100%;
      }
      .smallcard h3 {font-size: 1rem; margin:0 0 8px 0;}
      .smallcard p {font-size: .95rem; margin:0; line-height:1.7; color:#334155;}
      .keywordbox {
        background: white;
        border: 2px solid #C7D2FE;
        border-radius: 20px;
        padding: 18px 18px 12px 18px;
        margin-bottom: 14px;
        box-shadow: 0 10px 28px rgba(37,99,235,.08);
      }
      .resultbox {
        background:#F8FAFC; border:1px solid #E5E7EB; border-radius:16px; padding:14px 16px; margin-bottom: 12px;
      }
      .stDownloadButton button, .stButton button {
        border-radius: 12px; height: 2.8rem; font-weight: 700;
      }
      .kicker {font-size:.82rem; color:#475569; font-weight:700; letter-spacing:.04em; text-transform:uppercase; margin-bottom:6px;}
    </style>
    """,
    unsafe_allow_html=True,
)


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


def friendly_error_message(err: Exception) -> str:
    msg = str(err)
    if "429" in msg or "TooManyRequests" in msg:
        return "Google Trends側が一時的に混み合っています。少し待ってから再実行するか、期間を短くして試してください。"
    return "処理中にエラーが発生しました。入力条件かデータ形式を見直して、もう一度試してください。"


@st.cache_data(show_spinner=False, ttl=1800)
def fetch_google_trends_cached(keyword: str, geo="JP", timeframe="today 5-y", gprop="", cat=0, max_retries=5) -> pd.DataFrame:
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
            time.sleep(random.uniform(0.8, 1.6))
            pytrends.build_payload([keyword], cat=int(cat), timeframe=timeframe, geo=geo, gprop=gprop)
            df = pytrends.interest_over_time()
            if df is None or df.empty:
                raise ValueError("Google Trendsからデータを取得できませんでした。")
            if "isPartial" in df.columns:
                df = df.drop(columns=["isPartial"])
            return df.rename(columns={keyword: "value"}).reset_index()
        except Exception as e:
            last_err = e
            msg = str(e)
            is_429 = (
                ("429" in msg)
                or ("TooManyRequests" in msg)
                or isinstance(e, getattr(pytrends_exceptions, "TooManyRequestsError", tuple()))
            )
            if attempt < max_retries:
                wait = 8 + (attempt - 1) * 10 if is_429 else 4 + attempt * 2
                time.sleep(wait)
            else:
                raise last_err
    raise last_err


def build_analysis_tables(ts: pd.DataFrame, keyword: str):
    if len(ts) < 24:
        raise ValueError("季節性分解には最低24か月分のデータを推奨します。")

    result = ts.copy()
    decomp = seasonal_decompose(result["value"], model="multiplicative", period=12, extrapolate_trend="freq")
    result["trend"] = decomp.trend
    result["seasonal"] = decomp.seasonal
    result["resid"] = decomp.resid
    result["adjusted"] = result["value"] / result["seasonal"]
    result["yoy"] = result["value"].pct_change(12)
    result["adjusted_yoy"] = result["adjusted"].pct_change(12)
    result["month_num"] = result.index.month
    month_map = {1:"1月",2:"2月",3:"3月",4:"4月",5:"5月",6:"6月",7:"7月",8:"8月",9:"9月",10:"10月",11:"11月",12:"12月"}
    result["month"] = result["month_num"].map(month_map)

    month_tbl = (
        result.groupby(["month_num", "month"], as_index=False)["seasonal"]
        .mean()
        .sort_values("month_num")
        .rename(columns={"seasonal": "seasonal_index"})
    )
    month_tbl["平均との差"] = month_tbl["seasonal_index"] - 1
    month_tbl["平均比(%)"] = (month_tbl["seasonal_index"] - 1) * 100

    top_month = month_tbl.sort_values("seasonal_index", ascending=False).iloc[0]
    low_month = month_tbl.sort_values("seasonal_index", ascending=True).iloc[0]
    latest = result.iloc[-1]

    overview = pd.DataFrame({
        "項目": ["対象キーワード", "分析期間", "最新月", "最新値", "最も強い月", "最も弱い月", "最大季節指数", "最小季節指数"],
        "内容": [
            keyword,
            f"{result.index.min().strftime('%Y-%m')} ～ {result.index.max().strftime('%Y-%m')}",
            result.index.max().strftime('%Y-%m'),
            float(latest["value"]),
            str(top_month["month"]),
            str(low_month["month"]),
            round(float(top_month["seasonal_index"]), 3),
            round(float(low_month["seasonal_index"]), 3),
        ]
    })
    return overview, month_tbl[["month_num", "month", "seasonal_index", "平均との差", "平均比(%)"]], result


def plot_results(result_df: pd.DataFrame, month_tbl: pd.DataFrame, keyword: str):
    fig1, ax = plt.subplots(figsize=(12.5, 5.6))
    ax.plot(result_df.index, result_df["value"], linewidth=2.8, color="#2563EB", label="実績")
    ax.plot(result_df.index, result_df["adjusted"], linewidth=2.4, color="#F59E0B", label="季節調整後")
    ax.fill_between(result_df.index, result_df["value"], result_df["adjusted"], color="#DBEAFE", alpha=0.45)
    ax.set_title(f"{keyword}｜推移と季節調整後", loc="left", fontsize=16, fontweight="bold")
    ax.set_ylabel("検索指数")
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", linestyle="--", alpha=0.22)
    ax.legend(frameon=False, ncol=2, loc="upper left")
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
    fig1.tight_layout()

    fig2, ax = plt.subplots(figsize=(12.5, 5.2))
    vals = month_tbl["seasonal_index"].values
    colors = ["#7C3AED" if v >= 1 else "#CBD5E1" for v in vals]
    bars = ax.bar(month_tbl["month"], vals, color=colors, width=0.68)
    ax.axhline(1, color="#EF4444", linestyle="--", linewidth=1.4, alpha=0.8)
    ax.set_title(f"{keyword}｜月別の強弱", loc="left", fontsize=16, fontweight="bold")
    ax.set_ylabel("季節指数")
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", linestyle="--", alpha=0.18)
    ax.set_ylim(max(0, vals.min() - 0.12), vals.max() + 0.12)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_x() + bar.get_width()/2, v + 0.015, f"{v:.2f}", ha="center", va="bottom", fontsize=9)
    fig2.tight_layout()
    return fig1, fig2


def metric_description_sheet() -> pd.DataFrame:
    return pd.DataFrame([
        ["value（検索指数）", "Google TrendsやCSVの元データ。大きいほど検索需要が強いことを示します。"],
        ["trend", "長期的な流れ。短期のブレをならして見た基調です。"],
        ["seasonal / seasonal_index（季節指数）", "1.00が平均。1.10なら平均より約10%強く、0.90なら約10%弱い状態です。"],
        ["resid", "トレンドと季節性で説明しきれない一時的なブレです。"],
        ["adjusted（季節調整後）", "季節要因を取り除いた値。実力ベースの動きを見たいときに使います。"],
        ["yoy（前年比）", "元データの前年同月比。直近の伸び縮みを確認できます。"],
        ["adjusted_yoy（季節調整後前年比）", "季節要因を除いたうえでの前年同月比。実質的な伸びを見やすい指標です。"],
        ["month_num / month", "月番号と月名です。月別の強弱を見るために使います。"],
    ], columns=["指標", "意味"])


def make_excel_bytes(overview, month_tbl, result_df) -> bytes:
    excel_buffer = io.BytesIO()
    result_export = result_df.copy().reset_index().rename(columns={"index": "date"})
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        overview.to_excel(writer, sheet_name="overview", index=False)
        month_tbl.to_excel(writer, sheet_name="month", index=False)
        result_export.to_excel(writer, sheet_name="result", index=False)
        metric_description_sheet().to_excel(writer, sheet_name="指標の説明", index=False)

    excel_buffer.seek(0)
    wb = load_workbook(excel_buffer)
    head_fill = PatternFill(fill_type="solid", fgColor="DCEAFE")
    border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name="Yu Gothic UI", size=10)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = border
        for cell in ws[1]:
            cell.font = Font(name="Yu Gothic UI", size=10, bold=True)
            cell.fill = head_fill
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 42)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def fig_to_png_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, bbox_inches="tight", format="png")
    buf.seek(0)
    return buf.getvalue()


def csv_bytes(result_df: pd.DataFrame) -> bytes:
    result_export = result_df.copy().reset_index().rename(columns={"index": "date"})
    return result_export.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def read_uploaded_csv(uploaded_file):
    raw = uploaded_file.getvalue()
    for enc in ["utf-8-sig", "cp932", "utf-8", "shift_jis"]:
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc)
        except Exception:
            continue
    raise ValueError("CSVを読み込めませんでした。文字コードを見直してください。")


st.markdown(
    """
    <div class="hero">
      <h1>Google Trends 季節性分解ツール</h1>
      <p>
        キーワードの検索需要を、<b>長期トレンド</b> と <b>月ごとの季節性</b> に分けて見られるツールです。<br>
        需要のピーク月、弱い月、直近の実質的な伸びを、Google Trends または CSV から確認できます。
      </p>
      <div class="pillrow">
        <div class="pill">Keyword中心のUI</div>
        <div class="pill">Google Trends / CSV対応</div>
        <div class="pill">Excel / CSV / PNG出力</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown('<div class="smallcard"><h3>何ができる？</h3><p>検索需要の波を分解して、いつ強いか・弱いか・最近の伸びが本物かを見られます。</p></div>', unsafe_allow_html=True)
with c2:
    st.markdown('<div class="smallcard"><h3>分析ロジック</h3><p>検索指数を月次化し、12か月周期の季節性分解でトレンドと季節性に切り分けています。</p></div>', unsafe_allow_html=True)
with c3:
    st.markdown('<div class="smallcard"><h3>使い方</h3><p>Keywordを入れて分析開始するだけです。まずは一般的な単語1つから試すのがおすすめです。</p></div>', unsafe_allow_html=True)

st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

left, right = st.columns([1.18, 0.82], gap="large")

with left:
    st.markdown('<div class="keywordbox">', unsafe_allow_html=True)
    st.markdown('<div class="kicker">STEP 1</div>', unsafe_allow_html=True)
    st.subheader("Keywordを入力")
    keyword = st.text_input("分析したいキーワード", placeholder="例：転職、花粉、NISA", label_visibility="collapsed")
    st.caption("ここにKeywordを入れて実行します。まずは単語1つがおすすめです。")
    st.markdown('</div>', unsafe_allow_html=True)

    input_method = st.radio("入力方法", ["Google Trendsで取得", "CSVをアップロード"], horizontal=True)

    if input_method == "Google Trendsで取得":
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            geo_label = st.selectbox("地域", list(GEO_MAP.keys()))
        with col_b:
            timeframe_label = st.selectbox("期間", list(TIMEFRAME_MAP.keys()), index=2)
        with col_c:
            gprop_label = st.selectbox("検索種別", list(GPROP_MAP.keys()))
        with st.expander("上級設定"):
            cat = st.number_input("カテゴリID", min_value=0, step=1, value=0)
            retry_count = st.slider("再試行回数", min_value=2, max_value=6, value=5)
            st.caption("通常はそのままで大丈夫です。")
        uploaded_file = None
    else:
        uploaded_file = st.file_uploader("CSVをアップロード", type=["csv"])
        st.caption("日付列と数値列が入ったCSVを読み込みます。週次データも月次に自動変換します。")
        geo_label = "日本 (JP)"
        timeframe_label = "過去5年"
        gprop_label = "通常のウェブ検索"
        cat = 0
        retry_count = 5

    run = st.button("分析開始", type="primary", use_container_width=True)

with right:
    st.markdown('<div class="resultbox"><b>出力内容</b><br>① 需要推移と季節調整後の比較<br>② 月別の強弱グラフ<br>③ 分析結果テーブル<br>④ Excel / CSV / PNG ダウンロード</div>', unsafe_allow_html=True)
    st.markdown('<div class="resultbox"><b>見るポイント</b><br>・最も需要が強い月はいつか<br>・平常時の実力は伸びているか<br>・一時的な季節要因に引っ張られていないか</div>', unsafe_allow_html=True)
    with st.expander("指標の見方"):
        st.dataframe(metric_description_sheet(), use_container_width=True, hide_index=True)

if run:
    try:
        keyword_clean = keyword.strip()
        if input_method == "Google Trendsで取得":
            if not keyword_clean:
                st.warning("Keywordを入力してください。")
                st.stop()
            with st.spinner("Google Trendsから取得して分析しています…"):
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
            with st.spinner("CSVを読み込んで分析しています…"):
                raw_csv = read_uploaded_csv(uploaded_file)
                date_col, value_col = infer_csv_columns(raw_csv)
                ts = normalize_monthly_index(raw_csv, date_col, value_col)
                if not keyword_clean:
                    keyword_clean = Path(uploaded_file.name).stem

        overview, month_tbl, result_df = build_analysis_tables(ts, keyword_clean or "Keyword")
        fig1, fig2 = plot_results(result_df, month_tbl, keyword_clean or "Keyword")

        top_month = month_tbl.sort_values("seasonal_index", ascending=False).iloc[0]["month"]
        low_month = month_tbl.sort_values("seasonal_index", ascending=True).iloc[0]["month"]
        latest_adj_yoy = result_df["adjusted_yoy"].dropna().iloc[-1] if result_df["adjusted_yoy"].dropna().size else np.nan
        if pd.isna(latest_adj_yoy):
            summary = f"最も需要が強い月は {top_month}、弱い月は {low_month} です。前年比はデータ不足のため未算出です。"
        else:
            sign = "伸びています" if latest_adj_yoy >= 0 else "弱含みです"
            summary = f"最も需要が強い月は {top_month}、弱い月は {low_month}。季節要因を除いた直近の前年比は {latest_adj_yoy:.1%} で、足元は {sign}。"

        st.success(summary)

        g1, g2 = st.columns(2)
        with g1:
            st.pyplot(fig1, use_container_width=True)
        with g2:
            st.pyplot(fig2, use_container_width=True)

        st.subheader("指標の意味")
        st.dataframe(metric_description_sheet(), use_container_width=True, hide_index=True)

        st.subheader("分析結果テーブル")
        show_df = result_df.copy().reset_index().rename(columns={"index": "date"})
        st.dataframe(show_df.tail(24), use_container_width=True, hide_index=True)

        excel_data = make_excel_bytes(overview, month_tbl, result_df)
        csv_data = csv_bytes(result_df)
        png1_data = fig_to_png_bytes(fig1)
        png2_data = fig_to_png_bytes(fig2)
        safe = clean_name(keyword_clean or "keyword")

        st.subheader("ダウンロード")
        d1, d2, d3, d4 = st.columns(4)
        with d1:
            st.download_button("Excelを保存", excel_data, file_name=f"{safe}_seasonality.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="excel_dl")
        with d2:
            st.download_button("CSVを保存", csv_data, file_name=f"{safe}_seasonality_result.csv", mime="text/csv", use_container_width=True, key="csv_dl")
        with d3:
            st.download_button("推移グラフPNG", png1_data, file_name=f"{safe}_trend.png", mime="image/png", use_container_width=True, key="png1_dl")
        with d4:
            st.download_button("季節性グラフPNG", png2_data, file_name=f"{safe}_seasonality.png", mime="image/png", use_container_width=True, key="png2_dl")

    except Exception as e:
        st.error(friendly_error_message(e))
        with st.expander("見直しポイント"):
            st.markdown(
                "- Google Trendsなら期間を短くする\n"
                "- 検索語を少し具体化する\n"
                "- 同じ条件で連続実行しすぎない\n"
                "- CSVなら日付列と数値列が入っているか確認する"
            )
