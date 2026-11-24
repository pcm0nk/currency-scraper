from pathlib import Path

URL = "https://servatmandi.com/"

# Table reading direction on Daily_Report: "ltr" (left-to-right) or "rtl" (right-to-left)
TABLE_DIRECTION = "rtl"  # or "ltr"
REPORT_TITLE = "نرخ های مهم"
# Exact column order on the report sheet (follows this order as written)
REPORT_COLUMNS = [
    "item",
    "current",
    "previous",
    "trend",
    "change",
    "change_pct",
]
# Optional fixed column widths for Daily_Report (use None for auto-fit)
REPORT_COLUMN_WIDTHS = {
    "item": 18,
    "current": 18,
    "previous": 18,
    "trend": 12,
    "change": 18,
    "change_pct": 15,
}
# Configurable headers for the Daily_Report sheet
REPORT_HEADERS = {
    "item": "نماد",
    "current": "قیمت فعلی",
    "previous": "قیمت دیروز",
    "trend": "روند",
    "change": "تغییر",
    "change_pct": "تغییر %",
}

# Base item keys used for scraping, history tracking, and calculations
SCRAPE_TARGETS = [
    "تتر",
    "دلار آمریکا",
    "شاخص کل بورس",
    "سکه امامی",
    "نقره 999",
    "طلا 18 عیار",
    "انس طلا",
    "بیت‌کوین",
    "اتریوم",
]

# Items included in report calculations and display
REPORT_TARGETS = [
    "تتر",
    "دلار آمریکا",
    "شاخص کل بورس",
    "سکه امامی",
    "نقره 999",
    "طلا 18 عیار",
    "انس طلا",
    "بیت‌کوین",
    "اتریوم",
]

# Optional pretext (prefix) and post-text (suffix) per item for display
ITEM_TEXT_AFFICES = {
    "تتر": {"prefix": "دلار آمریکا ", "suffix": "(تتر-ریال)"},
    "دلار آمریکا": {"prefix": "", "suffix": " (آزاد-ریال)"},
    "شاخص کل بورس": {"prefix": "", "suffix": ""},
    "سکه امامی": {"prefix": "", "suffix": " (ریال)"},
    "نقره 999": {"prefix": "", "suffix": "(ریال)"},
    "طلا 18 عیار": {"prefix": "", "suffix": " هر گرم (ریال)"},
    "انس طلا": {"prefix": "($) ", "suffix": "انس جهانی"},
    "بیت‌کوین": {"prefix": "($) ", "suffix": "بیت کوین"},
    "اتریوم": {"prefix": "($) ", "suffix": "اتریوم"},
}

OUTPUT_DIR = Path("output")
EXCEL_FILE = OUTPUT_DIR / "market_tracker.xlsx"
HISTORY_SHEET_NAME = "Daily_Data"
REPORT_SHEET_NAME = "Daily_Report"

# Report column toggles
SHOW_NUMBER_CHANGE = True
SHOW_PCT_CHANGE = True
COLOR_CHANGE_WITH_DIRECTION = True  # Default: match up/down/unchanged colors

# Theme colors in hexadecimal (editable)
THEME_COLORS = {
    "header_bg": "#0F2942",
    "header_fg": "#FFFFFF",
    "row_even": "#FFFFFF",
    "row_odd": "#1DAFFF",
    "row_text_even": "#111827",
    "row_text_odd": "#071E36",
    "up": "#1E7E34",         # A bit dark green
    "down": "#B71C1C",       # A bit dark red
    "unchanged": "#5C6672",  # No change
    "border": "#CBD5E1",
}