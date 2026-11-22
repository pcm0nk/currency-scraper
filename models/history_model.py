from datetime import date, datetime, timedelta
from pathlib import Path
import re
from typing import Dict, List, Optional, Tuple
import openpyxl


DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")


class HistoryModel:
    def __init__(self, excel_path: Path, sheet_name: str, targets: List[str]):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.targets = targets
        self.headers = ["Date"] + targets

    def load_all_records(self) -> Dict[str, Dict[str, float]]:
        """
        Returns historical records keyed by ISO date string ('YYYY-MM-DD').
        """
        if not self.excel_path.exists():
            return {}

        wb = openpyxl.load_workbook(self.excel_path)
        if self.sheet_name not in wb.sheetnames:
            return {}

        ws = wb[self.sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            return {}

        header_row = [str(h).strip() if h is not None else "" for h in rows[0]]
        date_col_idx = header_row.index("Date") if "Date" in header_row else 0

        records: Dict[str, Dict[str, float]] = {}

        for row in rows[1:]:
            if not row or all(cell is None for cell in row):
                continue

            raw_date = row[date_col_idx] if date_col_idx < len(row) else None
            if raw_date is None:
                continue

            row_date = self._normalize_date_str(raw_date)
            if not DATE_PATTERN.match(row_date):
                continue

            entry: Dict[str, float] = {}
            for idx, col_name in enumerate(header_row):
                if idx == date_col_idx or not col_name:
                    continue
                if idx < len(row) and row[idx] is not None:
                    val = row[idx]
                    if isinstance(val, str) and DATE_PATTERN.match(val.strip()):
                        continue
                    try:
                        entry[col_name] = float(val)
                    except (ValueError, TypeError):
                        continue

            records[row_date] = entry

        return records

    def upsert_day_record(
        self, target_date: date, prices: Dict[str, float]
    ) -> Dict[str, Dict[str, float]]:
        """
        Inserts or updates `target_date` in the in-memory history dict and returns
        the chronologically sorted records.
        """
        records = self.load_all_records()
        date_key = target_date.isoformat()
        existing = records.get(date_key, {})
        merged = {**existing, **prices}
        records[date_key] = merged
        return dict(sorted(records.items(), key=lambda x: x[0]))

    def compute_day_comparison(
        self,
        records: Dict[str, Dict[str, float]],
        target_date: date,
        report_targets: Optional[List[str]] = None,
    ) -> List[Dict]:
        """
        Computes yesterday/previous-day comparison for `report_targets`.
        """
        date_key = target_date.isoformat()
        if date_key not in records:
            raise KeyError(f"No data found for {date_key}")

        items_to_compare = report_targets if report_targets is not None else self.targets
        today_data = records[date_key]
        prev_key, prev_data = self._find_previous_day(records, target_date)

        comparison = []
        for item in items_to_compare:
            today_val = today_data.get(item)
            yest_val = prev_data.get(item) if prev_data else None

            if today_val is not None and yest_val is not None:
                diff = today_val - yest_val
                if diff > 0:
                    direction = "▲"
                    status = "up"
                elif diff < 0:
                    direction = "▼"
                    status = "down"
                else:
                    direction = "—"
                    status = "unchanged"

                pct = (diff / yest_val) * 100 if yest_val != 0 else None
            else:
                diff = None
                pct = None
                direction = "—"
                status = "unchanged"

            comparison.append(
                {
                    "item": item,
                    "date": date_key,
                    "previous_date": prev_key,
                    "today": today_val,
                    "yesterday": yest_val,
                    "diff": diff,
                    "pct": pct,
                    "direction": direction,
                    "status": status,
                }
            )
        return comparison

    def _find_previous_day(
        self, records: Dict[str, Dict[str, float]], target_date: date
    ) -> Tuple[Optional[str], Optional[Dict[str, float]]]:
        yesterday_key = (target_date - timedelta(days=1)).isoformat()
        if yesterday_key in records:
            return yesterday_key, records[yesterday_key]

        prior_dates = [d for d in sorted(records.keys()) if d < target_date.isoformat()]
        if prior_dates:
            latest_prior = prior_dates[-1]
            return latest_prior, records[latest_prior]

        return None, None

    @staticmethod
    def _normalize_date_str(cell_val) -> str:
        if isinstance(cell_val, datetime):
            return cell_val.date().isoformat()
        if isinstance(cell_val, date):
            return cell_val.isoformat()
        return str(cell_val).strip()