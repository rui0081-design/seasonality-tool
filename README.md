# Google Trends 季節性分析ツール

## 起動方法
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 概要
- Google Trendsからキーワードの検索指数を取得
- 月次に整形して季節性分解
- 季節指数、季節調整値、前年比、季節調整後前年比を算出
- Excel / CSV / PNGをダウンロード可能

## Streamlit Cloud
- リポジトリ直下に `app.py` と `requirements.txt` を置く
- Pythonバージョン固定のため `runtime.txt` を同梱
- 依存ライブラリの相性崩れ対策としてバージョン固定
