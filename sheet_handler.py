import gspread
from google.oauth2.service_account import Credentials
import jaconv

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
            for row in rows[1:]:
                # 空行対策
                if not row or not row[0]:
                    continue
                
                official_name = row[0].strip()
                
                # B列以降（表記揺れ）をすべてループ
                for alias in row[1:]:
                    if alias.strip():
                        normalized_alias = jaconv.hira2kata(alias.strip())
                        mapping[normalized_alias] = official_name
            
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

    def record_daily_activities_batch(self, daily_data_map):
        """
        日別シートにデータをまとめて一括で書き込む
        daily_data_map: { 'YYYYMMDD': [ [(name, score), ...], ... ], ... }
        """
        sh = self.client.open_by_key(self.spreadsheet_key)
        
        for date_str, games_list in daily_data_map.items():
            # 1. シートの取得または作成
            try:
                worksheet = sh.worksheet(date_str)
                all_values = worksheet.get_all_values()
                if not all_values:
                    header = ["No."]
                    current_match_no = 0
                else:
                    header = all_values[0]
                    current_match_no = len(all_values) - 1 # ヘッダー分を引く
            except gspread.exceptions.WorksheetNotFound:
                # 存在しない場合は新規作成
                worksheet = sh.add_worksheet(title=date_str, rows="100", cols="20")
                header = ["No."]
                worksheet.update_cell(1, 1, "No.")
                current_match_no = 0
            
            player_to_col = {name: i + 1 for i, name in enumerate(header)}
            
            # 2. バッチ内の全プレイヤーを確認し、必要ならヘッダーを一度だけ拡張
            new_players_in_batch = []
            seen_in_batch = set()
            for game in games_list:
                for name, _ in game:
                    if name not in player_to_col and name not in seen_in_batch:
                        new_players_in_batch.append(name)
                        seen_in_batch.add(name)
            
            if new_players_in_batch:
                start_col = len(header) + 1
                cells = []
                for i, name in enumerate(new_players_in_batch):
                    col_idx = start_col + i
                    header.append(name)
                    player_to_col[name] = col_idx
                    cells.append(gspread.cell.Cell(row=1, col=col_idx, value=name))
                # ヘッダーの不足分を一括更新
                worksheet.update_cells(cells)
            
            # 3. 全試合分の行データを作成
            rows_to_append = []
            for game in games_list:
                current_match_no += 1
                row = [""] * len(header)
                row[0] = current_match_no
                for name, score in game:
                    col_idx = player_to_col[name]
                    row[col_idx - 1] = score
                rows_to_append.append(row)
            
            # 4. まとめて書き込み
            if rows_to_append:
                worksheet.append_rows(rows_to_append)