# Streamlit Cloud 公開手順

## 1. GitHubへアップロード
このフォルダの中身をそのままGitHubリポジトリ直下に置いてください。

必要な状態
- `app.py` がルートにある
- `requirements.txt` がルートにある
- `.streamlit/config.toml` がある

## 2. Streamlit Cloudでデプロイ
1. Streamlit Cloud にログイン
2. GitHub リポジトリを選択
3. Main file path に `app.py` を指定
4. Deploy を実行

## 3. 公開前の確認
- Keyword入力が最初にわかりやすいか
- Google Trends取得が通るか
- グラフが同サイズで表示されるか
- 日本語が文字化けしないか
- Excel / CSV / PNG が正常にダウンロードできるか

## 4. よくある注意
- Google Trendsの429制限が出る場合は時間を空けて再実行
- 12か月では前年比が十分に見えない場合があるため、基本は3年または5年推奨
