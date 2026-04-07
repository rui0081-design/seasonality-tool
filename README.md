# Google Trends 季節性分解ツール

キーワードの検索需要を、長期トレンドと月ごとの季節性に分けて見られるStreamlitアプリです。

## できること
- Google Trendsから検索指数を取得
- CSVをアップロードして同じ分析を実行
- 長期トレンドと季節調整後の比較
- 月別の強い月 / 弱い月の可視化
- Excel / CSV / PNG のダウンロード
- Excelに「指標の説明」シートを自動追加

## ファイル構成
- `app.py` : アプリ本体
- `requirements.txt` : 必要ライブラリ
- `.streamlit/config.toml` : テーマ設定
- `.gitignore` : GitHub用の除外設定

## ローカルで起動する
```bash
pip install -r requirements.txt
streamlit run app.py
```

## GitHubにアップする手順
1. GitHubで新しいリポジトリを作る
2. このフォルダの中身をそのままアップする
3. `app.py` がリポジトリ直下にある状態にする

## URL公開する手順（Streamlit）
1. Streamlitのデプロイ画面を開く
2. GitHub連携を行う
3. 対象リポジトリを選ぶ
4. Main file path に `app.py` を指定する
5. Deploy を押す

## 公開前の最終確認
- Keywordを入れて実行できるか
- Excel / CSV / PNG が保存できるか
- CSVアップロードでも動くか
- 日本語グラフが崩れていないか

## 注意
- Google Trends側が混み合うと、一時的に取得失敗することがあります
- その場合は少し待って再実行するか、期間を短くしてください
