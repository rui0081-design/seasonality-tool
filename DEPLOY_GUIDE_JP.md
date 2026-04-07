# 1分デプロイガイド

## 最短ルート
1. GitHubで新規リポジトリ作成
2. このフォルダ内のファイルを全部アップロード
3. Streamlitのデプロイ画面でそのリポジトリを選択
4. `app.py` を指定して公開
5. 発行されたURLを配布

## GitHubに置くもの
- app.py
- requirements.txt
- README.md
- .streamlit/config.toml
- .gitignore

## ハマりやすいポイント
- `app.py` がフォルダの中に入りすぎている
- requirements.txt をアップしていない
- 公開後すぐは初回起動に少し時間がかかる

## 迷ったら
- Main file path は `app.py`
- Python versionは通常そのままでOK
- URLが出たらそれが配布用リンク
