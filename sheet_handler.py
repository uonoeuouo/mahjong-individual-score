import gspread
from google.oauth2.service_account import Credentials

class SheetHandler:
    def __init__(self, keyfile, spreadsheet_key, sheet_name):
        self.spreadsheet_key = spreadsheet_key
        self.sheet_name = sheet_name
        
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_file(keyfile, scopes=scopes)
        self.client = gspread.authorize(creds)

    def append_game_data(self, rows):
        """データを追記する"""
        if not rows:
            return
        
        sh = self.client.open_by_key(self.spreadsheet_key)
        worksheet = sh.worksheet(self.sheet_name)
        worksheet.append_rows(rows)

    def get_id_mapping(self):
        """
        Membersシートを読み込み、{DiscordID(str): 選手名(str)} の辞書を返す
        """
        try:
            sh = self.client.open_by_key(self.spreadsheet_key)
            worksheet = sh.worksheet('ID_Mapping')
            rows = worksheet.get_all_values()
            
            mapping = {}
            # 1行目はヘッダーとしてスキップ
            for row in rows[1:]:
                # A列(名前)とB列(ID)が両方ある場合のみ
                if len(row) >= 2 and row[0] and row[1]:
                    name = row[0].strip()
                    discord_id = row[1].strip()
                    mapping[discord_id] = name
            
            return mapping
        except Exception as e:
            print(f"ID辞書読み込みエラー: {e}")
            return {}