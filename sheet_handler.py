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
        self.service_account_email = getattr(creds, 'service_account_email', None)

    def _get_spreadsheet_owner_emails(self, spreadsheet):
        try:
            permissions = spreadsheet.list_permissions()
        except Exception as e:
            print(f"オーナー情報の取得に失敗しました: {e}")
            return []

        owner_emails = {
            permission.get('emailAddress', '').strip()
            for permission in permissions
            if permission.get('role') == 'owner' and permission.get('emailAddress')
        }
        return sorted(owner_emails)

    def _build_daily_sheet_protected_editors(self, spreadsheet):
        allowed_editors = set(self._get_spreadsheet_owner_emails(spreadsheet))
        if self.service_account_email:
            allowed_editors.add(self.service_account_email)
        return sorted(allowed_editors)

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
        protected_editors = self._build_daily_sheet_protected_editors(sh)
        writer = DailySheetWriter(sh, protected_editor_emails=protected_editors)
        writer.write_batch(daily_data_map)

    def record_stats_chombo_counts(self, stats_sheet_name='Stats'):
        """Statsシートの合計/チョンボ行を名前ベースの配列数式で維持する"""
        sh = self.client.open_by_key(self.spreadsheet_key)
        worksheet = sh.worksheet(stats_sheet_name)

        worksheet.update('A2', [['合計スコア']])
        worksheet.update(
            'B2',
            [[
                '=BYCOL(B1:1, LAMBDA(name, IF(name="", "", '
                'SUMIF(RawData!$B:$B, name, RawData!$C:$C) + '
                'SUMPRODUCT((RawData!$B:$B=name) * REGEXMATCH(RawData!$E:$E, "^チョンボ([0-9０-９]+)?$") * '
                'IFERROR(VALUE(REGEXEXTRACT(RawData!$E:$E, "([0-9０-９]+)$")), 1)) * -20)))'
            ]],
            value_input_option='USER_ENTERED'
        )

        worksheet.update('A12', [['チョンボ数']])
        worksheet.update(
            'B12',
            [[
                '=BYCOL(B1:1, LAMBDA(name, IF(name="", "", '
                'SUMPRODUCT((RawData!$B:$B=name) * REGEXMATCH(RawData!$E:$E, "^チョンボ([0-9０-９]+)?$") * '
                'IFERROR(VALUE(REGEXEXTRACT(RawData!$E:$E, "([0-9０-９]+)$")), 1)))))'
            ]],
            value_input_option='USER_ENTERED'
        )