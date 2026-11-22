from pathlib import Path

URL = "https://servatmandi.com/"

# Table reading direction on Daily_Report: "ltr" (left-to-right) or "rtl" (right-to-left)
TABLE_DIRECTION = "rtl"  # or "ltr"
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

# Items scraped from the homepage
SCRAPE_TARGETS = [
    "طلا 18 عیار",
    "انس طلا",
    "تتر",
    "بیت‌کوین",
    "اتریوم",
    "دلار آمریکا",
]

# Items displayed on the themed report sheet (add or deduct freely)
REPORT_TARGETS = [
    "طلا 18 عیار",
    "انس طلا",
    "تتر",
    "بیت‌کوین",
    "اتریوم",
    "دلار آمریکا",
]

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