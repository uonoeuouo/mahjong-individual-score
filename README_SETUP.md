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





# 🚀 Python Setup (Mac)
## 1.環境構築
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


# 🚀 Python Setup (Windows)
## 1. Pythonのインストール
公式サイトからPythonインストーラーをダウンロードして実行します。
**重要:** インストール画面の下部にある **"Add Python to PATH"** に必ずチェックを入れてください。

## 2. 環境構築
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



# 🚀 共通設定
## 1. Googleスプレッドシートのセットアップ
個人戦を始める前に、新しいスプレッドシートを用意してください。

1. サークルのGoogleアカウントにあるテンプレートスプレッドシートを開きます。
2. そのテンプレートをコピーし、ファイル名を「2026後期個人戦」という形式で作成します。
3. コピーしたスプレッドシートを `mahjong-score@...` に編集者として共有します。
4. その後、`RawData` と `Stats` のシートについて、シートの保護設定を変更します。
5. 各シート名を右クリックし、`シートを保護` を開いて `キャンセル` を押したあと、`権限を変更` を選択します。
6. `mahjong-score@...` にチェックを入れて、編集権限を付与します。

サービスアカウント(mahjong-score@...)の正しいメールアドレスは管理者に聞いてください。

## 2.git clone
`git clone https://...`でこのリポジトリをクローンしてください。

## 3.設定ファイルの作成
プロジェクトディレクトリに`.env`ファイルを作成してください。

また、`credentials.json`ファイルを作成してください。

各ファイルの内容は管理者に聞いてください。

## 4.COMPLETION_MESSAGEの変更
`config.py`の`COMPLETION_MESSAGE`のスプレッドシートのリンクを、新しいスプレッドシートのリンクに置き換えてください。