# 再デプロイ手順

1. GitHubのリポジトリ直下にこのzipの中身をそのまま置く
2. Streamlit Cloudで対象アプリを開く
3. Settings から Main file path が `app.py` になっているか確認
4. `Clear cache` を実行
5. 再デプロイする

## 重要
- 以前のフォルダ構成が残っていると Main file path がズレることがあります
- blank画面になる時は `app.py` がルートにないか、依存関係のインストール失敗のことが多いです
- この版は `pytrends` の読み込みを遅延させているので、少なくとも画面自体は出る構成です
