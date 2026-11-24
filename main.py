import config
from controllers.tracker_controller import TrackerController
from models.history_model import HistoryModel
from models.scraper_model import ServatmandiScraperModel
from models.text_maker_model import TextMakerModel
from views.excel_view import ExcelView


def main() -> None:
    scraper = ServatmandiScraperModel(
        url=config.URL,
        targets=config.SCRAPE_TARGETS,
    )
    history = HistoryModel(
        excel_path=config.EXCEL_FILE,
        sheet_name=config.HISTORY_SHEET_NAME,
        targets=config.SCRAPE_TARGETS,
    )
    text_maker = TextMakerModel(
        item_affixes=config.ITEM_TEXT_AFFICES,
        report_title=config.REPORT_TITLE,
    )
    view = ExcelView(
        excel_path=config.EXCEL_FILE,
        history_sheet_name=config.HISTORY_SHEET_NAME,
        report_sheet_name=config.REPORT_SHEET_NAME,
        history_targets=config.SCRAPE_TARGETS,
        colors=config.THEME_COLORS,
        report_headers=config.REPORT_HEADERS,
        report_columns=config.REPORT_COLUMNS,
        column_widths=config.REPORT_COLUMN_WIDTHS,
        text_maker=text_maker,
        show_number_change=config.SHOW_NUMBER_CHANGE,
        show_pct_change=config.SHOW_PCT_CHANGE,
        color_change_with_direction=config.COLOR_CHANGE_WITH_DIRECTION,
        direction=config.TABLE_DIRECTION,
    )

    controller = TrackerController(
        scraper=scraper,
        history=history,
        view=view,
        report_targets=config.REPORT_TARGETS,
    )
    controller.run_daily_capture()


if __name__ == "__main__":
    main()