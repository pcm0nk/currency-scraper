from pathlib import Path

URL = "https://servatmandi.com/"

# Base item keys used for scraping, history storage, and delta calculations
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

# Optional fallback search labels when a target appears under a slightly different name
SCRAPE_ALIASES = {
    "شاخص کل بورس": ["شاخص کل بورس", "شاخص کل"],
    "سکه امامی": ["سکه امامی", "سکه طرح جدید"],
    "نقره 999": ["نقره 999", "نقره ۹۹۹"],
    "طلا 18 عیار": ["طلا 18 عیار", "طلا ۱۸ عیار"],
}

# Optional numeric multipliers applied to scraped values (e.g., multiplying نقره by 1,000)
ITEM_MULTIPLIERS = {
    "نقره 999": 1000,
}

# Optional pretext (prefix) and post-text (suffix) applied only to report item labels

ITEM_TEXT_AFFICES = {
    "تتر": {
        "label": "دلار آمریکا",
        "prefix": "",
        "suffix": " (تتر-ریال)",
    },
    "دلار آمریکا": {
        "label": "دلار آمریکا",
        "prefix": "",
        "suffix": " (آزاد-ریال)",
    },
    "شاخص کل بورس": {"prefix": "", "suffix": ""},
    "سکه امامی": {"prefix": "", "suffix": " (ریال)"},
    "نقره 999": {"prefix": "", "suffix": "(ریال)"},
    "طلا 18 عیار": {"prefix": "", "suffix": " هر گرم (ریال)"},
    "انس طلا": {
        "label": "انس جهانی",
        "prefix": "",
        "suffix": "",
    },
    "بیت‌کوین": {"prefix": "($) ", "suffix": "بیت کوین"},
    "اتریوم": {"prefix": "($) ", "suffix": "اتریوم"},
}

# Items included in report calculations and display (keep aligned with SCRAPE_TARGETS)
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


OUTPUT_DIR = Path("output")
EXCEL_FILE = OUTPUT_DIR / "market_tracker.xlsx"
HISTORY_SHEET_NAME = "Daily_Data"
REPORT_SHEET_NAME = "Daily_Report"

# Banner title prefix; formatted with Jalali date (<REPORT_TITLE> <day> <Persian month> <year>)
REPORT_TITLE = "نرخ های مهم"

# Comparison & change column options
SHOW_NUMBER_CHANGE = True
SHOW_PCT_CHANGE = True

# False keeps numeric change & % text neutral while coloring only the trend arrow
COLOR_CHANGE_WITH_DIRECTION = False

# Table reading direction: "ltr" or "rtl"
TABLE_DIRECTION = "rtl"

# Theme colors (hex strings): alternating white and blue rows with white text
THEME_COLORS = {
    "header_bg": "#4472C4",
    "header_fg": "#FFFFFF",
    "row_even": "#FFFFFF",
    "row_odd": "#1DAFFF",
    "row_text_even": "#000000",
    "row_text_odd": "#000000",
    "neutral_change": "#000000",
    "up": "#1E7E34",
    "down": "#B71C1C",
    "unchanged": "#000000",
    "border": "#B4C6E7",
}

# Column headers on Daily_Report
REPORT_HEADERS = {
    "item": "شرح نرخ",
    "current": "نرخ فعلی",
   # "previous": "نرخ قبلی",
    "trend": "روند",
    "change": "تغییر",
    "change_pct": "درصد تغییر",
}

# Exact column order on Daily_Report (follows this order as written)
REPORT_COLUMNS = [
    "item",
    "current",
   # "previous",
    "trend",
    "change",
    "change_pct",
]

# Specific width per column key (set a float/int for fixed width, or None for auto-fit)
REPORT_COLUMN_WIDTHS = {
    "item": 30,
    "current": 22,
    "previous": 18,
    "trend": 12,
    "change": 22,
    "change_pct": 16,
}