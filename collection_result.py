from dataclasses import dataclass, field
from datetime import timedelta, timezone


JST = timezone(timedelta(hours=9))

RAW_ROW_NAME_INDEX = 1
RAW_ROW_SCORE_INDEX = 2
RAW_ROW_RANK_INDEX = 3
RAW_ROW_CHOMBO_INDEX = 4


@dataclass
class DailySheetBatch:
    games: list[list[tuple[str, float, float, str]]] = field(default_factory=list)
    chombo_counts: dict[str, int] = field(default_factory=dict)

    def add_game(self, rows, chombo_names):
        self.games.append(_to_daily_game(rows))

        for name in chombo_names:
            self.chombo_counts[name] = self.chombo_counts.get(name, 0) + 1

    def to_sheet_payload(self):
        return {
            'games': self.games,
            'chombo_counts': self.chombo_counts,
        }


@dataclass
class CollectionResult:
    raw_rows: list[list] = field(default_factory=list)
    error_logs: list[str] = field(default_factory=list)
    daily_batches: dict[str, DailySheetBatch] = field(default_factory=dict)

    @property
    def game_count(self):
        return len(self.raw_rows) // 4

    def add_error(self, timestamp, error):
        self.error_logs.append(f"⚠️ {timestamp} の投稿: {error}")

    def add_game(self, rows, created_at, chombo_names):
        self.raw_rows.extend(rows)
        sheet_date = created_at.strftime('%Y%m%d')
        batch = self.daily_batches.setdefault(sheet_date, DailySheetBatch())
        batch.add_game(rows, chombo_names)

    def has_rows(self):
        return bool(self.raw_rows)

    def daily_sheet_payload(self):
        return {
            date_str: batch.to_sheet_payload()
            for date_str, batch in self.daily_batches.items()
        }


def _to_daily_game(rows):
    return [
        (
            row[RAW_ROW_NAME_INDEX],
            row[RAW_ROW_SCORE_INDEX],
            row[RAW_ROW_RANK_INDEX],
            row[RAW_ROW_CHOMBO_INDEX],
        )
        for row in rows
    ]
