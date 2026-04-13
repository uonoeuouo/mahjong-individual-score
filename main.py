import discord
from collection_service import run_collection_process
from config import (
    CHANNEL_ID,
    COMPLETION_MESSAGE,
    JSON_KEYFILE,
    SHEET_NAME,
    SPREADSHEET_KEY,
    TOKEN,
)
from sheet_handler import SheetHandler

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
        await run_collection_process(message.channel, client, sheet_handler, COMPLETION_MESSAGE)

# Bot起動
if __name__ == '__main__':
    client.run(TOKEN)