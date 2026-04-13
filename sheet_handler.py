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
        日別シートに一括書き込みを行う。
        集計はスプレッドシートの配列数式（BYCOL）に任せる。
        """
        sh = self.client.open_by_key(self.spreadsheet_key)
        
        for date_str, games_list in daily_data_map.items():
            try:
                worksheet = sh.worksheet(date_str)
                all_values = worksheet.get_all_values()
            except gspread.exceptions.WorksheetNotFound:
                # 1. 新規作成: シートを作り、A列のラベルとB列に配列数式をセットする
                # 列は少し余裕を持たせてZ列(26列)まで確保
                worksheet = sh.add_worksheet(title=date_str, rows="200", cols="26")
                
                # B列に配置するだけで、1行目に名前がある全ての列を自動計算する魔法の数式群
                formulas = [
                    ["No. / 選手名", ""],
                    ["合計スコア", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(c))))'],
                    ["1着数", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(RANK(INDEX(c, r), INDEX($B$9:$Z, r))=1, 1, 0))))))))'],
                    ["2着数", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(RANK(INDEX(c, r), INDEX($B$9:$Z, r))=2, 1, 0))))))))'],
                    ["3着数", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(RANK(INDEX(c, r), INDEX($B$9:$Z, r))=3, 1, 0))))))))'],
                    ["4着数", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(RANK(INDEX(c, r), INDEX($B$9:$Z, r))=4, 1, 0))))))))'],
                    ["平均順位", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", IFERROR(SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, RANK(INDEX(c, r), INDEX($B$9:$Z, r)))))) / COUNT(c), ""))))'],
                    ["▼ 試合記録", ""]
                ]
                
                # 数式として認識させるため value_input_option='USER_ENTERED' を必須とする
                worksheet.update("A1:B8", formulas, value_input_option='USER_ENTERED')
                all_values = [["No. / 選手名"]]

            # 2. ヘッダー（1行目）の確認と拡張
            header = all_values[0] if all_values and len(all_values) > 0 else ["No. / 選手名"]
            player_to_col = {name: i + 1 for i, name in enumerate(header) if name}
            
            new_players = []
            for game in games_list:
                for name, _, _ in game:
                    if name not in player_to_col:
                        new_players.append(name)
                        player_to_col[name] = len(header) + len(new_players)
            
            if new_players:
                # 1行目に新しい選手名を追加するだけで、B列の配列数式が自動で計算範囲を広げてくれる
                start_col = len(header) + 1
                header.extend(new_players)
                cells = [gspread.cell.Cell(row=1, col=start_col + i, value=name) for i, name in enumerate(new_players)]
                worksheet.update_cells(cells)

            # 3. 試合スコアの追記（9行目以降）
            rows_to_append = []
            start_no = len(all_values) - 8 if len(all_values) > 8 else 0
            
            for game in games_list:
                start_no += 1
                row = [""] * len(header)
                row[0] = start_no
                for name, score, _ in game:
                    col_idx = player_to_col[name]
                    row[col_idx - 1] = score
                rows_to_append.append(row)
            
            if rows_to_append:
                worksheet.append_rows(rows_to_append)