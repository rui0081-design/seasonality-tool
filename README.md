# Google Trends 季節性分析ツール

Google Trendsの検索指数を使って、キーワードの
- 長期トレンド
- 月別の季節性
- 直近の実力（季節調整後）
- 前年比 / 季節調整後前年比

を一画面で確認できるStreamlitアプリです。

## 主な機能
- Google Trendsからデータ取得（pytrends）
- 季節性分解（trend / seasonal / residual）
- 月別季節指数の可視化
- 季節調整値の算出
- 前年比 / 季節調整後前年比
- サマリー自動生成
- Excel / CSV / PNG ダウンロード

## 画面で見られるもの
- 検索需要の推移（実績 / 季節調整後 / 長期トレンド）
- 月別季節指数
- 月別季節指数テーブル
- 直近結果テーブル
- 強い月 / 弱い月 / 実力評価 / 施策示唆のサマリー

## Excel出力の内容
- `summary`
- `seasonality_by_month`
- `result_detail`
- `metric_guide`
- `how_to_read`

列幅、ヘッダー色、固定表示、フィルタなどを整えています。

## ファイル構成
- `app.py` : アプリ本体
- `requirements.txt` : 必要ライブラリ
- `.streamlit/config.toml` : Streamlitテーマ設定
- `README.md` : 説明
- `DEPLOY_GUIDE_JP.md` : GitHub / Streamlit Cloud公開手順
- `.gitignore` : Git管理用

## ローカル起動
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 使い方
1. Keyword を入力
2. 地域 / 期間を選択
3. 必要なら詳細設定を開く
4. 分析開始
5. グラフ確認後、Excel / CSV / PNG をダウンロード

## 補足
- Google Trendsの仕様上、取得が一時的に失敗することがあります
- その場合は少し待つか、期間やKeywordを見直してください
- 季節性分析には24か月以上の月次データが必要です
