import gspread


DAILY_SHEET_FORMULAS = [
    ["No. / 選手名", ""],
    ["合計スコア", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(c))))'],
    ["1着数", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(RANK(INDEX(c, r), INDEX($B$9:$Z, r))=1, 1, 0))))))))'],
    ["2着数", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(RANK(INDEX(c, r), INDEX($B$9:$Z, r))=2, 1, 0))))))))'],
    ["3着数", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(RANK(INDEX(c, r), INDEX($B$9:$Z, r))=3, 1, 0))))))))'],
    ["4着数", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(RANK(INDEX(c, r), INDEX($B$9:$Z, r))=4, 1, 0))))))))'],
    ["平均順位", '=BYCOL($B$9:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", IFERROR(SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, RANK(INDEX(c, r), INDEX($B$9:$Z, r)))))) / COUNT(c), ""))))'],
    ["▼ 試合記録", ""],
]


class DailySheetWriter:
    def __init__(self, spreadsheet):
        self.spreadsheet = spreadsheet

    def write_batch(self, daily_data_map):
        for date_str, games_list in daily_data_map.items():
            worksheet, all_values = self._get_or_create_daily_worksheet(date_str)
            header, player_to_col = self._ensure_header_players(worksheet, all_values, games_list)
            self._append_game_rows(worksheet, all_values, header, player_to_col, games_list)

    def _get_or_create_daily_worksheet(self, date_str):
        try:
            worksheet = self.spreadsheet.worksheet(date_str)
            all_values = worksheet.get_all_values()
            return worksheet, all_values
        except gspread.exceptions.WorksheetNotFound:
            worksheet = self.spreadsheet.add_worksheet(title=date_str, rows="200", cols="26")
            worksheet.update("A1:B8", DAILY_SHEET_FORMULAS, value_input_option='USER_ENTERED')
            return worksheet, [["No. / 選手名"]]

    def _ensure_header_players(self, worksheet, all_values, games_list):
        header = all_values[0] if all_values else ["No. / 選手名"]
        player_to_col = {name: i + 1 for i, name in enumerate(header) if name}

        new_players = []
        for game in games_list:
            for name, _, _ in game:
                if name not in player_to_col:
                    new_players.append(name)
                    player_to_col[name] = len(header) + len(new_players)

        if new_players:
            start_col = len(header) + 1
            header.extend(new_players)
            cells = [gspread.cell.Cell(row=1, col=start_col + i, value=name) for i, name in enumerate(new_players)]
            worksheet.update_cells(cells)

        return header, player_to_col

    def _append_game_rows(self, worksheet, all_values, header, player_to_col, games_list):
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
