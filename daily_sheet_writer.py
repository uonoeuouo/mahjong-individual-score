import gspread


MANAGED_DAILY_SHEET_PROTECTION_DESCRIPTION = "managed-by-bot-daily-sheet-lock"


DAILY_SHEET_FORMULAS = [
    ["No. / 選手名", ""],
    ["合計スコア", '=BYCOL($B$10:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(c) + IFERROR(INDEX($8:$8, 1, COLUMN(c)), 0) * -20)))'],
    ["1着数", '=BYCOL($B$10:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(AND(COUNTIF(INDEX($B$10:$Z, r), ">"&INDEX(c, r)) + 1 <= 1, 1 <= COUNTIF(INDEX($B$10:$Z, r), ">"&INDEX(c, r)) + COUNTIF(INDEX($B$10:$Z, r), INDEX(c, r))), 1 / COUNTIF(INDEX($B$10:$Z, r), INDEX(c, r)), 0))))))))'],
    ["2着数", '=BYCOL($B$10:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(AND(COUNTIF(INDEX($B$10:$Z, r), ">"&INDEX(c, r)) + 1 <= 2, 2 <= COUNTIF(INDEX($B$10:$Z, r), ">"&INDEX(c, r)) + COUNTIF(INDEX($B$10:$Z, r), INDEX(c, r))), 1 / COUNTIF(INDEX($B$10:$Z, r), INDEX(c, r)), 0))))))))'],
    ["3着数", '=BYCOL($B$10:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(AND(COUNTIF(INDEX($B$10:$Z, r), ">"&INDEX(c, r)) + 1 <= 3, 3 <= COUNTIF(INDEX($B$10:$Z, r), ">"&INDEX(c, r)) + COUNTIF(INDEX($B$10:$Z, r), INDEX(c, r))), 1 / COUNTIF(INDEX($B$10:$Z, r), INDEX(c, r)), 0))))))))'],
    ["4着数", '=BYCOL($B$10:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, IF(AND(COUNTIF(INDEX($B$10:$Z, r), ">"&INDEX(c, r)) + 1 <= 4, 4 <= COUNTIF(INDEX($B$10:$Z, r), ">"&INDEX(c, r)) + COUNTIF(INDEX($B$10:$Z, r), INDEX(c, r))), 1 / COUNTIF(INDEX($B$10:$Z, r), INDEX(c, r)), 0))))))))'],
    ["平均順位", '=BYCOL($B$10:$Z, LAMBDA(c, IF(INDEX($1:$1, 1, COLUMN(c))="", "", IFERROR(SUM(MAP(SEQUENCE(ROWS(c)), LAMBDA(r, IF(INDEX(c, r)="", 0, (COUNTIF(INDEX($B$10:$Z, r), ">"&INDEX(c, r)) + 1 + COUNTIF(INDEX($B$10:$Z, r), ">"&INDEX(c, r)) + COUNTIF(INDEX($B$10:$Z, r), INDEX(c, r))) / 2)))) / COUNT(c), ""))))'],
    ["チョンボ数", ""],
    ["▼ 試合記録", ""],
]


class DailySheetWriter:
    def __init__(self, spreadsheet, protected_editor_emails=None):
        self.spreadsheet = spreadsheet
        self.protected_editor_emails = sorted({
            email.strip() for email in (protected_editor_emails or []) if email and email.strip()
        })

    def write_batch(self, daily_data_map):
        for date_str, batch_data in daily_data_map.items():
            games_list = batch_data.get('games', [])
            chombo_counts = batch_data.get('chombo_counts', {})
            worksheet, all_values = self._get_or_create_daily_worksheet(date_str)
            self._ensure_summary_formulas(worksheet)
            header, player_to_col = self._ensure_header_players(worksheet, all_values, games_list)
            self._append_game_rows(worksheet, all_values, header, player_to_col, games_list)
            self._update_chombo_row(worksheet, player_to_col, chombo_counts)
            self._ensure_sheet_protection(worksheet)

    def _ensure_sheet_protection(self, worksheet):
        if not self.protected_editor_emails:
            return

        metadata = self.spreadsheet.fetch_sheet_metadata(params={
            "fields": "sheets(properties(sheetId),protectedRanges(protectedRangeId,description,range))"
        })
        target_sheet = next(
            (s for s in metadata.get("sheets", []) if s.get("properties", {}).get("sheetId") == worksheet.id),
            None,
        )
        if not target_sheet:
            return

        protected_ranges = target_sheet.get("protectedRanges", [])
        managed_ranges = [
            r for r in protected_ranges
            if r.get("description") == MANAGED_DAILY_SHEET_PROTECTION_DESCRIPTION
        ]

        requests = []
        if managed_ranges:
            primary = managed_ranges[0]
            requests.append({
                "updateProtectedRange": {
                    "protectedRange": {
                        "protectedRangeId": primary["protectedRangeId"],
                        "range": {"sheetId": worksheet.id},
                        "description": MANAGED_DAILY_SHEET_PROTECTION_DESCRIPTION,
                        "warningOnly": False,
                        "editors": {"users": self.protected_editor_emails},
                    },
                    "fields": "range,description,warningOnly,editors",
                }
            })

            for duplicate in managed_ranges[1:]:
                requests.append({
                    "deleteProtectedRange": {
                        "protectedRangeId": duplicate["protectedRangeId"],
                    }
                })
        else:
            requests.append({
                "addProtectedRange": {
                    "protectedRange": {
                        "range": {"sheetId": worksheet.id},
                        "description": MANAGED_DAILY_SHEET_PROTECTION_DESCRIPTION,
                        "warningOnly": False,
                        "editors": {"users": self.protected_editor_emails},
                    }
                }
            })

        if requests:
            self.spreadsheet.batch_update({"requests": requests})

    def _ensure_summary_formulas(self, worksheet):
        labels = [[row[0]] for row in DAILY_SHEET_FORMULAS]
        worksheet.update("A1:A9", labels, value_input_option='USER_ENTERED')

        summary_formulas = [[row[1]] for row in DAILY_SHEET_FORMULAS[1:7]]
        worksheet.update("B2:B7", summary_formulas, value_input_option='USER_ENTERED')

    def _get_or_create_daily_worksheet(self, date_str):
        try:
            worksheet = self.spreadsheet.worksheet(date_str)
            all_values = worksheet.get_all_values()

            # 旧レイアウト(8行目が試合記録開始)は1行追加して移行する
            if len(all_values) >= 8 and all_values[7] and all_values[7][0] == '▼ 試合記録':
                worksheet.insert_row([""], index=8)
                self._ensure_summary_formulas(worksheet)
                all_values = worksheet.get_all_values()

            return worksheet, all_values
        except gspread.exceptions.WorksheetNotFound:
            worksheet = self.spreadsheet.add_worksheet(title=date_str, rows="200", cols="26")
            self._ensure_summary_formulas(worksheet)
            return worksheet, [["No. / 選手名"]]

    def _ensure_header_players(self, worksheet, all_values, games_list):
        header = all_values[0] if all_values else ["No. / 選手名"]
        player_to_col = {name: i + 1 for i, name in enumerate(header) if name}

        new_players = []
        for game in games_list:
            for name, _, _, _ in game:
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
        start_no = len(all_values) - 9 if len(all_values) > 9 else 0

        for game in games_list:
            start_no += 1
            row = [""] * len(header)
            row[0] = start_no
            for name, score, _, _ in game:
                col_idx = player_to_col[name]
                row[col_idx - 1] = score
            rows_to_append.append(row)

        if rows_to_append:
            worksheet.append_rows(rows_to_append)

    def _update_chombo_row(self, worksheet, player_to_col, chombo_counts):
        if not chombo_counts:
            return

        row8 = worksheet.row_values(8)
        cells = []
        for name, add_count in chombo_counts.items():
            col_idx = player_to_col.get(name)
            if not col_idx:
                continue

            current_val = 0.0
            if len(row8) >= col_idx and row8[col_idx - 1] != '':
                try:
                    current_val = float(row8[col_idx - 1])
                except ValueError:
                    current_val = 0.0

            cells.append(gspread.cell.Cell(row=8, col=col_idx, value=current_val + add_count))

        if cells:
            worksheet.update_cells(cells, value_input_option='USER_ENTERED')
