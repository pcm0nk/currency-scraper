import os
from pathlib import Path
from typing import Dict, List, Optional
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from models.text_maker_model import TextMakerModel


DEFAULT_REPORT_HEADERS: Dict[str, str] = {
    "item": "شرح نرخ",
    "current": "نرخ فعلی",
    "previous": "نرخ قبلی",
    "trend": "روند",
    "change": "تغییر",
    "change_pct": "درصد تغییر",
}

DEFAULT_REPORT_COLUMNS: List[str] = [
    "item",
    "current",
    "previous",
    "trend",
    "change",
    "change_pct",
]

FONT_FAMILY = "B Titr"


def _to_int(val):
    if isinstance(val, (int, float)):
        return int(round(val))
    return val


def _to_one_decimal(val):
    if isinstance(val, (int, float)):
        return round(float(val), 1)
    return val


def _to_argb(hex_color: str) -> str:
    """Converts '#RRGGBB' or 'RRGGBB' into an 8-char ARGB hex string for openpyxl."""
    cleaned = hex_color.strip().lstrip("#")
    if len(cleaned) == 6:
        return f"FF{cleaned.upper()}"
    if len(cleaned) == 8:
        return cleaned.upper()
    raise ValueError(f"Invalid hex color: '{hex_color}'")


def _is_file_locked(file_path: Path) -> bool:
    """Returns True if `file_path` exists and cannot be opened for read/write."""
    if not file_path.exists():
        return False
    try:
        with open(file_path, "r+b"):
            return False
    except (PermissionError, OSError):
        return True


def _wait_for_any_key() -> None:
    """Waits for a single key press on Windows, falling back to Enter on other OSes."""
    if os.name == "nt":
        import msvcrt

        msvcrt.getch()
    else:
        try:
            input()
        except EOFError:
            pass


class ExcelView:
    def __init__(
        self,
        excel_path: Path,
        history_sheet_name: str,
        report_sheet_name: str,
        history_targets: List[str],
        colors: Dict[str, str],
        report_headers: Optional[Dict[str, str]] = None,
        report_columns: Optional[List[str]] = None,
        column_widths: Optional[Dict[str, Optional[float]]] = None,
        item_affixes: Optional[Dict[str, Dict[str, str]]] = None,
        report_title: Optional[str] = None,
        show_number_change: bool = True,
        show_pct_change: bool = True,
        color_change_with_direction: bool = True,
        direction: str = "ltr",
        text_maker: Optional[TextMakerModel] = None,
    ):
        self.excel_path = excel_path
        self.history_sheet_name = history_sheet_name
        self.report_sheet_name = report_sheet_name
        self.history_targets = history_targets
        self.colors = colors
        self.headers_map = {**DEFAULT_REPORT_HEADERS, **(report_headers or {})}
        self.report_columns = (
            list(report_columns)
            if report_columns is not None
            else list(DEFAULT_REPORT_COLUMNS)
        )
        self.text_maker = text_maker or TextMakerModel(
            item_affixes=item_affixes,
            report_title=report_title,
        )
        self.column_widths = column_widths or {}
        self.item_affixes = item_affixes or {}
        self.report_title = report_title
        self.show_number_change = show_number_change
        self.show_pct_change = show_pct_change
        self.color_change_with_direction = color_change_with_direction

        normalized_dir = direction.strip().lower()
        if normalized_dir not in ("ltr", "rtl"):
            raise ValueError(
                f"Invalid table direction '{direction}'. Use 'ltr' or 'rtl'."
            )
        self.direction = normalized_dir

    @property
    def is_rtl(self) -> bool:
        return self.direction == "rtl"

    def ensure_file_unlocked(self) -> None:
        while _is_file_locked(self.excel_path):
            print(
                f"'{self.excel_path}' is currently open. "
                "Please close the Excel file first and press any key to continue..."
            )
            _wait_for_any_key()

    def _load_or_create_workbook(self) -> openpyxl.Workbook:
        self.excel_path.parent.mkdir(parents=True, exist_ok=True)
        if self.excel_path.exists():
            return openpyxl.load_workbook(self.excel_path)
        wb = openpyxl.Workbook()
        wb.active.title = self.history_sheet_name
        return wb

    def save_daily_history(self, records: Dict[str, Dict[str, float]]) -> Path:
        wb = self._load_or_create_workbook()

        if self.history_sheet_name in wb.sheetnames:
            idx = wb.sheetnames.index(self.history_sheet_name)
            wb.remove(wb[self.history_sheet_name])
            ws = wb.create_sheet(title=self.history_sheet_name, index=idx)
        else:
            ws = wb.create_sheet(title=self.history_sheet_name)

        ws.sheet_view.rightToLeft = self.is_rtl

        dynamic_items: List[str] = list(self.history_targets)
        for day_vals in records.values():
            for key in day_vals.keys():
                if key not in dynamic_items:
                    dynamic_items.append(key)

        headers = ["Date"] + dynamic_items
        ws.append(headers)

        for date_key, values in records.items():
            row = [date_key] + [_to_int(values.get(item)) for item in dynamic_items]
            ws.append(row)

        for row_idx in range(1, ws.max_row + 1):
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = Font(
                    name=FONT_FAMILY,
                    size=11,
                    bold=(row_idx == 1),
                    color="000000",
                )
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if row_idx >= 2 and col_idx >= 2 and isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0"

        for col_idx, header in enumerate(headers, start=1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = max(len(str(header)) + 6, 16)

        self.ensure_file_unlocked()
        wb.save(self.excel_path)
        return self.excel_path

    def save_daily_report(
        self,
        report_date_str: str,
        comparison_rows: List[Dict],
        previous_date_str: Optional[str] = None,
    ) -> Path:
        wb = self._load_or_create_workbook()

        if self.report_sheet_name in wb.sheetnames:
            idx = wb.sheetnames.index(self.report_sheet_name)
            wb.remove(wb[self.report_sheet_name])
            ws = wb.create_sheet(title=self.report_sheet_name, index=idx)
        else:
            ws = wb.create_sheet(title=self.report_sheet_name)

        ws.sheet_view.rightToLeft = self.is_rtl

        valid_keys = set(DEFAULT_REPORT_COLUMNS)
        header_specs = []
        for col_key in self.report_columns:
            if col_key not in valid_keys:
                continue
            if col_key == "change" and not self.show_number_change:
                continue
            if col_key == "change_pct" and not self.show_pct_change:
                continue
            header_specs.append((col_key, self.headers_map[col_key]))

        if not header_specs:
            header_specs = [("item", self.headers_map["item"])]

        headers = [label for _, label in header_specs]
        num_cols = len(headers)

        # Row 1: Title banner & date — B Titr, size 20, centered
        title_text = self.text_maker.format_report_title(
            report_date_str=report_date_str,
            previous_date_str=previous_date_str,
        )

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        title_cell = ws.cell(row=1, column=1, value=title_text)
        title_cell.font = Font(
            name=FONT_FAMILY,
            size=20,
            bold=True,
            color=_to_argb(self.colors["header_fg"]),
        )
        title_cell.fill = PatternFill(
            start_color=_to_argb(self.colors["header_bg"]),
            end_color=_to_argb(self.colors["header_bg"]),
            fill_type="solid",
        )
        title_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )
        ws.row_dimensions[1].height = 34

        # Row 2: Column headers (including شرح نرخ) — B Titr, size 18, centered
        header_fill = PatternFill(
            start_color=_to_argb(self.colors["header_bg"]),
            end_color=_to_argb(self.colors["header_bg"]),
            fill_type="solid",
        )
        header_font = Font(
            name=FONT_FAMILY,
            size=18,
            bold=True,
            color=_to_argb(self.colors["header_fg"]),
        )
        thin_side = Side(style="thin", color=_to_argb(self.colors["border"]))
        cell_border = Border(
            left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
        )

        ws.row_dimensions[2].height = 30
        for col_idx, (_, col_name) in enumerate(header_specs, start=1):
            h_cell = ws.cell(row=2, column=col_idx, value=col_name)
            h_cell.fill = header_fill
            h_cell.font = header_font
            h_cell.border = cell_border
            h_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        # Alternating row fills & status colors
        row_even_fill = PatternFill(
            start_color=_to_argb(self.colors["row_even"]),
            end_color=_to_argb(self.colors["row_even"]),
            fill_type="solid",
        )
        row_odd_fill = PatternFill(
            start_color=_to_argb(self.colors["row_odd"]),
            end_color=_to_argb(self.colors["row_odd"]),
            fill_type="solid",
        )

        up_hex = _to_argb(self.colors["up"])
        down_hex = _to_argb(self.colors["down"])
        unchanged_hex = _to_argb(self.colors["unchanged"])
        neutral_change_hex = _to_argb(
            self.colors.get("neutral_change", self.colors["row_text_even"])
        )

        for idx, entry in enumerate(comparison_rows):
            excel_row = idx + 3
            ws.row_dimensions[excel_row].height = 24

            is_even = idx % 2 == 0
            row_fill = row_even_fill if is_even else row_odd_fill
            base_text_hex = _to_argb(
                self.colors["row_text_even"] if is_even else self.colors["row_text_odd"]
            )

            status = entry.get("status", "unchanged")
            if status == "up":
                dir_hex = up_hex
            elif status == "down":
                dir_hex = down_hex
            else:
                dir_hex = unchanged_hex

            delta_hex = dir_hex if self.color_change_with_direction else neutral_change_hex

            today_val = _to_int(entry.get("today"))
            yest_val = _to_int(entry.get("yesterday"))
            diff_val = _to_int(entry.get("diff"))
            pct_val = _to_one_decimal(entry.get("pct"))
            direction_symbol = entry.get("direction", "—")

            values_by_key: Dict[str, object] = {
                "item": self.text_maker.format_item_label(entry["item"]),
                "current": today_val if today_val is not None else "N/A",
                "previous": yest_val if yest_val is not None else "N/A",
                "trend": direction_symbol,
                "change": diff_val if diff_val is not None else "N/A",
                "change_pct": (pct_val / 100.0) if pct_val is not None else "N/A",
            }

            for col_idx, (col_key, _) in enumerate(header_specs, start=1):
                val = values_by_key[col_key]
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.fill = row_fill
                cell.border = cell_border
                cell.alignment = Alignment(horizontal="center", vertical="center")

                if col_key == "item":
                    cell.font = Font(
                        name=FONT_FAMILY, size=18, bold=True, color=base_text_hex
                    )
                elif col_key in ("current", "previous"):
                    cell.font = Font(
                        name=FONT_FAMILY, size=18, bold=True, color=base_text_hex
                    )
                    if isinstance(val, (int, float)):
                        cell.number_format = "#,##0"
                elif col_key == "trend":
                    cell.font = Font(
                        name=FONT_FAMILY, size=18, bold=True, color=dir_hex
                    )
                elif col_key == "change":
                    cell.font = Font(
                        name=FONT_FAMILY, size=18, bold=True, color=delta_hex
                    )
                    if isinstance(val, (int, float)):
                        cell.number_format = "+#,##0;-#,##0;0"
                elif col_key == "change_pct":
                    cell.font = Font(
                        name=FONT_FAMILY, size=18, bold=True, color=delta_hex
                    )
                    if isinstance(val, (int, float)):
                        cell.number_format = "+0.0%;-0.0%;0.0%"

        for col_idx, (col_key, col_name) in enumerate(header_specs, start=1):
            col_letter = get_column_letter(col_idx)
            custom_width = self.column_widths.get(col_key)
            if custom_width is not None:
                ws.column_dimensions[col_letter].width = float(custom_width)
            else:
                max_len = len(str(col_name))
                for r in range(1, ws.max_row + 1):
                    v = ws.cell(row=r, column=col_idx).value
                    if v is not None and not isinstance(v, openpyxl.cell.cell.MergedCell):
                        max_len = max(max_len, len(str(v)))
                ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

        self.ensure_file_unlocked()
        wb.save(self.excel_path)
        return self.excel_path