from datetime import date
from typing import List, Optional
from models.scraper_model import ServatmandiScraperModel
from models.history_model import HistoryModel
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
        print(f"Scraping live prices for {run_date.isoformat()}...")

        prices = self.scraper.fetch_prices()
        for name, val in prices.items():
            print(f"  {name}: {val:,.2f}")

        records = self.history.upsert_day_record(run_date, prices)
        self.view.save_daily_history(records)

        comparison = self.history.compute_day_comparison(
            records=records,
            target_date=run_date,
            report_targets=self.report_targets,
        )
        prev_date = comparison[0]["previous_date"] if comparison else None
        saved_path = self.view.save_daily_report(
            report_date_str=run_date.isoformat(),
            comparison_rows=comparison,
            previous_date_str=prev_date,
        )
        print(f"Updated history and generated '{self.view.report_sheet_name}' in: {saved_path}")