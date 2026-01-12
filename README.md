# Discord 麻雀スコア自動集計Bot

Discordのチャットに投稿された麻雀のスコアを読み取り、Googleスプレッドシートに自動で集計・保存するBotです。
増分更新（前回の続きから読み込み）、名前の表記揺れ吸収、スタッツの自動計算に対応しています。





## 📁 ディレクトリ構成

```text
.
├── main.py            # 実行ファイル
├── parser.py          # テキスト解析・正規化ロジック
├── sheet_handler.py   # スプレッドシート操作
├── .env               # 設定ファイル（Gitには含めない）
├── credentials.json   # GCPサービスアカウントキー（Gitには含めない）
└── README.md          # 説明書

```





## セットアップ手順（Mac）
### 1.環境構築
Python3が必要です。仮想環境を推奨します。
```
# プロジェクトフォルダへ移動
cd /path/to/project

# 仮想環境の作成と有効化
python -m venv venv
source venv/bin/activate

# 必要なライブラリのインストール
pip install discord.py gspread google-auth python-dotenv jaconv
```

### 2.Google Cloud設定
1. GCPコンソールでプロジェクトを作成し、「Google Sheets API」と「Google Drive API」を有効化する。

2. サービスアカウントを作成し、キー（JSON形式）をダウンロードして credentials.json にリネームしてフォルダに配置する。

3. 重要: スプレッドシートの「共有」ボタンから、サービスアカウントのメールアドレス（xxx@xxx.iam.gserviceaccount.com）を「編集者」として招待する。

### 3.設定ファイルの作成
`.env`ファイルを作成し、以下の内容を記述してください。
```
DISCORD_TOKEN=あなたのBotトークン
CHANNEL_ID=読み込むチャンネルのID(数字)
SPREADSHEET_KEY=スプレッドシートのID(URLの/d/xxx/editのxxx部分)
JSON_KEYFILE=credentials.json
TARGET_SHEET_NAME=RawData
```

## セットアップ手順（Windows）
## 🚀 セットアップ手順 (Windows)

### 1. Pythonのインストール
公式サイトからPythonインストーラーをダウンロードして実行します。
**重要:** インストール画面の下部にある **"Add Python to PATH"** に必ずチェックを入れてください。

### 2. 環境構築
コマンドプロンプト（またはPowerShell）を開き、以下のコマンドを実行します。

```cmd
# プロジェクトフォルダへ移動 (例: Desktop\discord-bot)
cd Desktop\discord-bot

# 仮想環境の作成
python -m venv venv

# 仮想環境の有効化 (コマンドプロンプトの場合)
venv\Scripts\activate.bat

# ※PowerShellの場合はこちら
:: venv\Scripts\Activate.ps1

# (先頭に (venv) と表示されればOKです)

# 必要なライブラリのインストール
pip install discord.py gspread google-auth python-dotenv jaconv
```



## 使い方
### スコアの投稿フォーマット
Discordの指定チャンネルで以下のように投稿します。

名前と点数の間はスペース（全角・半角問わず）。

4人の合計が0である必要があります。
```
Aさん 54.3
Bさん 12.1
Cさん -15.5
Dさん -50.9
```

### 集計の実行
Botが起動しているタイミングで、以下のコマンドを入力します。
```!collect```

Botの挙動:

前回の「✅ 集計完了」メッセージを探します。

それ以降の投稿を読み込みます。

名前を正規化（カタカナ変換＋名寄せ）します。

合計点チェックを行い、スプレッドシートに追記します。

完了メッセージを投稿します。




## ⚠️ トラブルシューティング
Error: <Response [404]>: スプレッドシートのIDが間違っているか、サービスアカウントが招待されていません。

合計エラー: 点数の合計が0になっていません。入力ミスを確認してください。