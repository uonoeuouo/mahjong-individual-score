import gspread
from google.oauth2.service_account import Credentials
import jaconv
from daily_sheet_writer import DailySheetWriter

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
                        normalized_alias = jaconv.kata2hira(alias.strip())
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
        日別シートに一括書き込みを行う。
        集計はスプレッドシートの配列数式（BYCOL）に任せる。
        """
        sh = self.client.open_by_key(self.spreadsheet_key)
        writer = DailySheetWriter(sh)
        writer.write_batch(daily_data_map)

    def record_stats_chombo_counts(self, _chombo_counts, stats_sheet_name='Stats'):
        """Statsシートの合計/チョンボ行を名前ベースの配列数式で維持する"""
        sh = self.client.open_by_key(self.spreadsheet_key)
        worksheet = sh.worksheet(stats_sheet_name)

        # 旧実装の固定値が残っていると配列数式がスピルできず #REF になるため先にクリアする
        worksheet.batch_clear(['B2:ZZ2', 'B12:ZZ12'])

        worksheet.update('A2', [['合計スコア']])
        worksheet.update(
            'B2',
            [[
                '=BYCOL(B1:1, LAMBDA(name, IF(name="", "", '
                'SUMIF(RawData!$B:$B, name, RawData!$C:$C) + '
                'COUNTIFS(RawData!$B:$B, name, RawData!$E:$E, "チョンボ") * -20)))'
            ]],
            value_input_option='USER_ENTERED'
        )

        worksheet.update('A12', [['チョンボ数']])
        worksheet.update(
            'B12',
            [[
                '=BYCOL(B1:1, LAMBDA(name, IF(name="", "", '
                'COUNTIFS(RawData!$B:$B, name, RawData!$E:$E, "チョンボ"))))'
            ]],
            value_input_option='USER_ENTERED'
        )