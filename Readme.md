# Currency & Market Scraper (MVC)

A modular Python application built with an **MVC (Model–View–Controller)** architecture that scrapes live market rates from `servatmandi.com`, maintains a historical daily log in Excel, compares values against yesterday (or the most recent prior recorded day), and generates a customizable, themed daily report sheet.

---

## Table of Contents

- [How the Project Works](#how-the-project-works)
- [Project Structure](#project-structure)
- [Installation & Setup](#installation--setup)
- [Where & How to Configure Everything (`config.py`)](#where--how-to-configure-everything-configpy)
  - [Scraping & Report Targets](#scraping--report-targets)
  - [Workbook & Sheet Names](#workbook--sheet-names)
  - [Comparison & Column Toggles](#comparison--column-toggles)
  - [Hexadecimal Theme Colors](#hexadecimal-theme-colors)
  - [Customizable Column Headers](#customizable-column-headers)
  - [Exact Column Order & Per-Column Widths](#exact-column-order--per-column-widths)
  - [Table Direction (LTR / RTL)](#table-direction-ltr--rtl)
- [Running the App & Excel Lock Handling](#running-the-app--excel-lock-handling)
- [Example `config.py` Reference](#example-configpy-reference)
- [Troubleshooting Tips](#troubleshooting-tips)

---

## How the Project Works

The application follows a **Model–View–Controller** workflow so scraping, historical math, and spreadsheet styling stay cleanly separated:

1. **Controller (`controllers/tracker_controller.py`)**
   - Triggered from `main.py` for a target date (defaulting to today's date).
   - Asks the **Scraper Model** to fetch live prices from the homepage.
   - Passes those prices to the **History Model** to insert or update today's entry in the chronological dataset.
   - Computes day-over-day comparisons (current value, previous value, trend arrow `▲` / `▼` / `—`, raw numeric difference, and percentage change) for the items listed in `REPORT_TARGETS`.
   - Passes the updated history and comparison rows to the **Excel View** to write both sheets.

2. **Models (`models/`)**
   - **`scraper_model.py` (`ServatmandiScraperModel`)**: Uses Playwright (headless Chromium) to load `https://servatmandi.com/`, locate each target label (accounting for spacing/zero-width joiners and dynamic DOM updates), extract the latest value cell, and normalize Persian (`۰-۹`), Arabic, and English (`0-9`) digits/commas into clean Python `float` values.
   - **`history_model.py` (`HistoryModel`)**: Loads existing records from the history sheet (`Daily_Data`), merges/updates the current day's values without losing earlier days, finds yesterday (or falls back to the latest prior recorded date if a day was skipped), and calculates differences and percentage changes.

3. **View (`views/excel_view.py` (`ExcelView`)**
   - Manages the Excel workbook (`.xlsx`) via `openpyxl`.
   - **History Sheet (`Daily_Data`)**: Stores one chronological row per date across tracked items in standard `["Date", ...items]` order so historical data can always be parsed reliably.
   - **Report Sheet (`Daily_Report`)**: Builds a formatted daily comparison table using your configured headers, column order, column widths, LTR/RTL direction, alternating row colors, and up/down/unchanged styling.
   - **File-Lock Guard**: Checks whether the target Excel file is currently open/locked before saving; if it is, it prompts you to close Excel and press a key before writing.

---

## Project Structure

```text
currency-scraper/
├── main.py                        # Entry point; wires config, models, view, and controller
├── config.py                      # Central configuration file for targets, theme, layout & toggles
├── README.md                      # Project documentation
├── .gitignore                     # Git ignore rules
├── models/
│   ├── __init__.py
│   ├── scraper_model.py           # Playwright browser scraper & number parser
│   └── history_model.py           # Daily history storage & yesterday-over-yesterday math
├── views/
│   ├── __init__.py
│   └── excel_view.py              # Excel formatting, RTL/LTR layout, widths & lock check
├── controllers/
│   ├── __init__.py
│   └── tracker_controller.py      # Orchestrates scraping, comparison, and Excel export
└── output/
    └── market_tracker.xlsx        # Generated workbook containing history & report sheets
```

---

## Installation & Setup

1. **Install Python dependencies** (`playwright` and `openpyxl`):
   ```bash
   pip install playwright openpyxl
   ```

2. **Install the Chromium browser binary for Playwright**:
   ```bash
   playwright install chromium
   ```

3. **Customize `config.py`** to match your preferred items, colors, headers, column order, column widths, and table direction.

---

## Where & How to Configure Everything (`config.py`)

All user-adjustable settings live in **`config.py`** in the root project folder so you don't need to edit core code for routine customization.

### Scraping & Report Targets
Keep scraping separate from what appears on your daily report:
- **`SCRAPE_TARGETS`**: List of exact item labels to scrape from `servatmandi.com` (for example: `"طلا 18 عیار"`, `"انس طلا"`, `"تتر"`, `"بیت‌کوین"`, `"اتریوم"`, `"دلار آمریکا"`).
- **`REPORT_TARGETS`**: List of items to display on the report sheet. You can **add or deduct** items here freely—if an item exists in `Daily_Data`, its comparison will be shown; if it's brand-new and has no prior record yet, missing comparison fields display `"N/A"`.

### Workbook & Sheet Names
- **`OUTPUT_DIR`** and **`EXCEL_FILE`**: Where the `.xlsx` file is saved (default: `output/market_tracker.xlsx`).
- **`HISTORY_SHEET_NAME`**: Name of the raw/chronological daily log tab (default: `"Daily_Data"`).
- **`REPORT_SHEET_NAME`**: Name of the styled daily comparison tab (default: `"Daily_Report"`).

### Comparison & Column Toggles
In `config.py`, control whether change columns appear and how they are colored:
- **`SHOW_NUMBER_CHANGE`** (`True` / `False`): Shows or hides the numeric difference column (`current - previous`).
- **`SHOW_PCT_CHANGE`** (`True` / `False`): Shows or hides the percentage change column.
- **`COLOR_CHANGE_WITH_DIRECTION`** (`True` / `False`): When `True` (default), numeric and `%` change values use the same color as the up/down/unchanged trend (`up`, `down`, `unchanged`). When `False`, they use a neutral text color (`neutral_change` or fallback row text color).

### Hexadecimal Theme Colors
Customize colors in **`THEME_COLORS`** using hexadecimal strings (such as `"#FFFFFF"` or `"#1DAFFF"`):
- **`header_bg`** & **`header_fg`**: Title banner and header row background and text colors.
- **`row_even`** & **`row_odd`**: Alternating data row fill colors (defaulting to white `#FFFFFF` and blue `#1DAFFF`).
- **`row_text_even`** & **`row_text_odd`**: Base text colors for even and odd rows so text stays readable against each background.
- **`up`**: Color for positive movement (default dark green, e.g., `#1E7E34`).
- **`down`**: Color for negative movement (default dark red, e.g., `#B71C1C`).
- **`unchanged`**: Color when there is no change (`0`) or no prior value to compare against.
- **`border`**: Grid border color around cells.

### Customizable Column Headers
Use **`REPORT_HEADERS`** in `config.py` to rename any report column label (in English, Persian, or any custom label) without changing internal logic:
- `"item"`
- `"current"`
- `"previous"`
- `"trend"`
- `"change"`
- `"change_pct"`

### Exact Column Order & Per-Column Widths
- **`REPORT_COLUMNS`**: Controls the **exact left-to-right column order** in the generated report table based on how you list the keys in `config.py` (e.g., `["item", "current", "previous", "trend", "change", "change_pct"]`). Reordering the list reorders the table columns in that exact sequence (excluding `"change"` or `"change_pct"` if their toggle is `False`).
- **`REPORT_COLUMN_WIDTHS`**: Maps each column key (`"item"`, `"current"`, `"previous"`, `"trend"`, `"change"`, `"change_pct"`) to a specific numeric width (e.g., `24`, `18`, `12`), or `None` if you want that column auto-sized based on cell content.

### Table Direction (LTR / RTL)
Set **`TABLE_DIRECTION`** in `config.py` to:
- `"rtl"` for Right-to-Left worksheet view (`ws.sheet_view.rightToLeft = True`) and RTL-friendly text alignment.
- `"ltr"` for Left-to-Right worksheet view.

---

## Running the App & Excel Lock Handling

Run the script from the project root folder:

```bash
python main.py
```
*(On Windows, you can also run `py main.py`.)*

- **If `market_tracker.xlsx` is open in Excel when saving:** `ExcelView.ensure_file_unlocked()` detects the file lock and prints a prompt asking you to close the Excel file and press any key to continue. Once you close Excel and press a key, it retries and saves the updated workbook.

---

## Example `config.py` Reference

```python
from pathlib import Path

URL = "[https://servatmandi.com/](https://servatmandi.com/)"

SCRAPE_TARGETS = [
    "طلا 18 عیار",
    "انس طلا",
    "تتر",
    "بیت‌کوین",
    "اتریوم",
    "دلار آمریکا",
]

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

SHOW_NUMBER_CHANGE = True
SHOW_PCT_CHANGE = True
COLOR_CHANGE_WITH_DIRECTION = True

TABLE_DIRECTION = "rtl"  # "ltr" or "rtl"

THEME_COLORS = {
    "header_bg": "#0F2942",
    "header_fg": "#FFFFFF",
    "row_even": "#FFFFFF",
    "row_odd": "#1DAFFF",
    "row_text_even": "#111827",
    "row_text_odd": "#071E36",
    "up": "#1E7E34",
    "down": "#B71C1C",
    "unchanged": "#5C6672",
    "border": "#CBD5E1",
}

REPORT_HEADERS = {
    "item": "آیتم",
    "current": "فعلی",
    "previous": "قبلی",
    "trend": "روند",
    "change": "تغییر",
    "change_pct": "درصد تغییر",
}

REPORT_COLUMNS = [
    "item",
    "current",
    "previous",
    "trend",
    "change",
    "change_pct",
]

REPORT_COLUMN_WIDTHS = {
    "item": 24,
    "current": 18,
    "previous": 18,
    "trend": 12,
    "change": 18,
    "change_pct": 15,
}
```

---

## Troubleshooting Tips

- **Label mismatch or timeout while scraping:** Verify that each string in `SCRAPE_TARGETS` matches the label on `servatmandi.com` (including spaces or Persian characters). If a page takes longer or keeps background connections open, try `wait_until="domcontentloaded"` or temporarily set `headless=False` in `models/scraper_model.py` to inspect how the page renders.
- **First-time comparison shows `"—"` or `"N/A"`:** On the very first day a target item is recorded, there is no previous day in `Daily_Data` to compare against yet. Once a second day is recorded, day-over-day arrows, numeric differences, and `%` changes will populate automatically.
- **Corrupted or misordered history sheet after manual/early edits:** If `Daily_Data` has mismatched columns from an older run, remove the old `output/market_tracker.xlsx` file (after backing up any data you want to keep) and run `main.py` to generate a clean workbook.
```