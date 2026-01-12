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

    def get_name_mapping(self):
        """
        Membersシートを読み込み、表記揺れ変換用の辞書を作成する
        戻り値: {'aさん': 'Aさん', 'Ａさん': 'Aさん', ...}
        """
        try:
            sh = self.client.open_by_key(self.spreadsheet_key)
            worksheet = sh.worksheet('Members')
            rows = worksheet.get_all_values()
            
            mapping = {}
            # 1行目はヘッダーとしてスキップし、2行目から処理
            for row in rows[1:]:
                # 空行対策
                if not row or not row[0]:
                    continue
                
                official_name = row[0].strip()
                
                # B列以降（表記揺れ）をすべてループ
                for alias in row[1:]:
                    if alias.strip():
                        mapping[alias.strip()] = official_name
            
            return mapping
        except Exception as e:
            print(f"辞書読み込みエラー: {e}")
            return {} # エラー時は空の辞書を返す（変換なしで動作させるため）

    def append_game_data(self, rows):
        """データを追記する"""
        if not rows:
            return
        
        sh = self.client.open_by_key(self.spreadsheet_key)
        worksheet = sh.worksheet(self.sheet_name)
        worksheet.append_rows(rows)