import os

from dotenv import load_dotenv


# .envファイルを読み込む
load_dotenv()

# 設定値の取得
TOKEN = os.getenv('DISCORD_TOKEN')
CHANNEL_ID = int(os.getenv('CHANNEL_ID'))
SPREADSHEET_KEY = os.getenv('SPREADSHEET_KEY')
JSON_KEYFILE = os.getenv('JSON_KEYFILE')
SHEET_NAME = os.getenv('TARGET_SHEET_NAME')

COMPLETION_MESSAGE = "✅ 集計完了しました。スプレッドシートを更新しました。https://docs.google.com/spreadsheets/d/1ObtQLSf9g3F-94Uzti7-5RLvCPEZEvMFklVeJh7fC1c/edit?gid=1098435240#gid=1098435240"
