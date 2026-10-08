# 📁 ディレクトリ構成

```text
.
├── main.py            # 実行ファイル
├── config.py          # 環境変数と定数管理
├── collection_service.py # 集計処理ロジック
├── parser.py          # テキスト解析・正規化ロジック
├── daily_sheet_writer.py # 日別シート更新ロジック
├── sheet_handler.py   # スプレッドシート操作
├── .env               # 設定ファイル（Gitには含めない）
├── credentials.json   # GCPサービスアカウントキー（Gitには含めない）
└── README.md          # 説明書

```



# 🚀 初回セットアップの流れ
1. Gitをインストール
2. Pythonをインストール
3. VS Codeをインストール
4. リポジトリをclone
5. 仮想環境を作る
6. ライブラリをインストール
7. 環境変数を設定
8. COMPLETION_MESSAGEの変更
9. スプレッドシートを設定
10. Botを起動

### 1.Gitのインストール
以下の記事に従えばできると思います。
https://qiita.com/takeru-hirai/items/4fbe6593d42f9a844b1c

### 2.Pythonのインストール
以下の記事に従えばできると思います。
https://qiita.com/Obataskill/items/4cfe0d4d8cec8a140b8e

### 3.VSCodeのインストール
以下の記事に従えばできると思います。
https://qiita.com/yuri_777/items/b2d0472ba68ebe431859

また、Pythonの拡張機能をインストールしてください。

### 4.リポジトリをクローン
ターミナルで以下を実行すればできます。
`git clone https://github.com/uonoeuouo/mahjong-individual-score.git`


### 5.仮想環境を作る
(Macの場合)
```
# プロジェクトフォルダへ移動(パスは例です)
cd /User/mahjong-indivisual-score

# 仮想環境の作成と有効化
python -m venv venv
source venv/bin/activate

# (先頭に (venv) と表示されればOKです)
```

(Windowsの場合)
```cmd
# プロジェクトフォルダへ移動 (例: Desktop\discord-bot)
cd Desktop\mahjong-indivisual-score

# 仮想環境の作成
python -m venv venv

# 仮想環境の有効化 (コマンドプロンプトの場合)
venv\Scripts\activate.bat

# ※PowerShellの場合はこちら
:: venv\Scripts\Activate.ps1

# (先頭に (venv) と表示されればOKです)
# 仮想環境の有効化でエラーが出る場合があります。エラー文をLLMにぶち込んで解決策を教えてもらいましょう。
```


### 6.ライブラリをインストール
仮想環境下(先頭にvenvとついてる状態)で以下を実行します。
```
pip install discord.py gspread google-auth python-dotenv jaconv
```

### 7.環境関数の設定
`.env`と`credentials.json`を作ります。
ファイルの位置はこのREADMEの一番上にある構成の通りです。
中身は管理者に教えてもらってください。

### 8.COMPLETION_MESSAGEの変更
`config.py`の`COMPLETION_MESSAGE`のスプレッドシートのリンクを、新しいスプレッドシートのリンクに置き換えてください。


### 9. Googleスプレッドシートのセットアップ
個人戦を始める前に、新しいスプレッドシートを用意してください。

1. サークルのGoogleアカウントにあるテンプレートスプレッドシートを開きます。
2. そのテンプレートをコピーし、ファイル名を「2026後期個人戦」という形式で作成します。
3. コピーしたスプレッドシートを `mahjong-score@...` に編集者として共有します。
4. その後、`RawData` と `Stats` のシートについて、シートの保護設定を変更します。
5. 各シート名を右クリックし、`シートを保護` を開いて `キャンセル` を押したあと、`権限を変更` を選択します。
6. `mahjong-score@...` にチェックを入れて、編集権限を付与します。

サービスアカウント(mahjong-score@...)の正しいメールアドレスは管理者に聞いてください。