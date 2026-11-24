from datetime import date
from typing import List, Optional
from models.history_model import HistoryModel
from models.scraper_model import ServatmandiScraperModel
from views.excel_view import ExcelView


class TrackerController:
    def __init__(
        self,
        scraper: ServatmandiScraperModel,
        history: HistoryModel,
        view: ExcelView,
        report_targets: Optional[List[str]] = None,
    ):
        self.scraper = scraper
        self.history = history
        self.view = view
        self.report_targets = report_targets

    def run_daily_capture(self, target_date: Optional[date] = None) -> None:
        run_date = target_date or date.today()
        date_str = run_date.isoformat()

        print(f"Scraping live prices for {date_str}...")
        prices = self.scraper.scrape_prices()

        for item, val in prices.items():
            print(f"  {item}: {val:,.2f}")

        records = self.history.upsert_day_record(run_date, prices)
        comparison_rows = self.history.compute_day_comparison(
            records=records,
            target_date=run_date,
            report_targets=self.report_targets,
        )

        prev_date_str = comparison_rows[0]["previous_date"] if comparison_rows else None

        self.view.save_daily_history(records)
        saved_path = self.view.save_daily_report(
            report_date_str=date_str,
            comparison_rows=comparison_rows,
            previous_date_str=prev_date_str,
        )
        print(f"Saved updated workbook to: {saved_path}")