import os
import discord
from dotenv import load_dotenv
from parser import parse_and_validate_message
from sheet_handler import SheetHandler

# .envファイルを読み込む
load_dotenv()

# 設定値の取得
TOKEN = os.getenv('DISCORD_TOKEN')
CHANNEL_ID = int(os.getenv('CHANNEL_ID'))
SPREADSHEET_KEY = os.getenv('SPREADSHEET_KEY')
JSON_KEYFILE = os.getenv('JSON_KEYFILE')
SHEET_NAME = os.getenv('TARGET_SHEET_NAME')

COMPLETION_MESSAGE = "✅ 集計完了しました。スプレッドシートを更新しました。"

# Discord設定
intents = discord.Intents.default()
intents.message_content = True
client = discord.Client(intents=intents)

# SheetHandlerの初期化
sheet_handler = SheetHandler(JSON_KEYFILE, SPREADSHEET_KEY, SHEET_NAME)

@client.event
async def on_ready():
    print(f'Logged in as {client.user}')

@client.event
async def on_message(message):
    if message.author.bot or message.channel.id != CHANNEL_ID:
        return

    if message.content.startswith('!collect'):
        await message.channel.send("🔄 ログの読み込みと集計を開始します...")
        await run_collection_process(message.channel)

async def run_collection_process(channel):
    try:
        # 1. 前回の完了地点を探す
        last_checkpoint = None
        async for msg in channel.history(limit=100):
            if msg.author == client.user and msg.content.startswith("✅ 集計完了"):
                last_checkpoint = msg
                break
        
        # 履歴取得のイテレータ作成
        if last_checkpoint:
            print(f"前回の完了地点: {last_checkpoint.created_at}")
            history = channel.history(after=last_checkpoint, limit=None, oldest_first=True)
        else:
            print("全件読み込みモード")
            history = channel.history(limit=200, oldest_first=True)

        # 2. メッセージ処理ループ
        all_rows_to_add = []
        error_logs = []

        async for msg in history:
            if msg.author.bot or msg.content.startswith('!'):
                continue

            # parserモジュールを使って解析
            timestamp = msg.created_at.strftime('%Y/%m/%d %H:%M')
            rows, error = parse_and_validate_message(msg.content, timestamp)

            if error:
                error_logs.append(f"⚠️ {timestamp} の投稿: {error}")
            elif rows:
                all_rows_to_add.extend(rows)

        # 3. 書き込みと結果報告
        if all_rows_to_add:
            sheet_handler.append_game_data(all_rows_to_add)
            result_msg = f"{COMPLETION_MESSAGE}\n追加件数: {len(all_rows_to_add)//4} 試合"
        else:
            result_msg = "✅ 新しいスコア投稿はありませんでした。"

        if error_logs:
            result_msg += "\n\n【エラー報告】\n" + "\n".join(error_logs)

        await channel.send(result_msg)

    except Exception as e:
        await channel.send(f"❌ エラーが発生しました: {e}")
        print(f"Error: {e}")

# Bot起動
if __name__ == '__main__':
    client.run(TOKEN)