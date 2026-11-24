from datetime import date
from typing import Dict, Optional


PERSIAN_MONTHS = [
    "فروردین",
    "اردیبهشت",
    "خرداد",
    "تیر",
    "مرداد",
    "شهریور",
    "مهر",
    "آبان",
    "آذر",
    "دی",
    "بهمن",
    "اسفند",
]


class TextMakerModel:
    def __init__(
        self,
        item_affixes: Optional[Dict[str, Dict[str, str]]] = None,
        report_title: Optional[str] = None,
    ):
        self.item_affixes = item_affixes or {}
        self.report_title = report_title

    def format_item_label(self, raw_item: str) -> str:
        """
        Formats an item's report label using an optional replacement label (`label` or `replace`)
        plus configured `prefix` and `suffix` without altering `raw_item` for history/math.
        """
        affix = self.item_affixes.get(raw_item, {})
        base_label = affix.get("label")
        if base_label is None:
            base_label = affix.get("replace")
        if base_label is None:
            base_label = raw_item

        prefix = affix.get("prefix", "")
        suffix = affix.get("suffix", "")
        return f"{prefix}{base_label}{suffix}"

    @staticmethod
    def gregorian_to_jalali(g_year: int, g_month: int, g_day: int) -> tuple[int, int, int]:
        """Converts a Gregorian date (Y, M, D) to Solar Hijri / Jalali (JY, JM, JD)."""
        g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
        gy2 = g_year + 1 if g_month > 2 else g_year
        days = (
            355666
            + (365 * g_year)
            + ((gy2 + 3) // 4)
            - ((gy2 + 99) // 100)
            + ((gy2 + 399) // 400)
            + g_day
            + g_d_m[g_month - 1]
        )
        jy = -1595 + (33 * (days // 12053))
        days %= 12053
        jy += 4 * (days // 1461)
        days %= 1461
        if days > 365:
            jy += (days - 1) // 365
            days = (days - 1) % 365
        if days < 186:
            jm = 1 + (days // 31)
            jd = 1 + (days % 31)
        else:
            jm = 7 + ((days - 186) // 30)
            jd = 1 + ((days - 186) % 30)
        return jy, jm, jd

    def format_report_title(
        self,
        report_date_str: str,
        previous_date_str: Optional[str] = None,
    ) -> str:
        """
        Formats the banner title as: <REPORT_TITLE> <day> <Persian month> <year>
        (e.g., 'نرخ های مهم 2 آذر 1405').
        """
        parsed = date.fromisoformat(report_date_str)
        jy, jm, jd = self.gregorian_to_jalali(parsed.year, parsed.month, parsed.day)
        month_name = PERSIAN_MONTHS[jm - 1]
        jalali_date = f"{jd} {month_name} {jy}"

        if self.report_title:
            return f"{self.report_title.strip()} {jalali_date}".strip()

        prev_label = previous_date_str if previous_date_str else "N/A"
        return f"Daily Market Report — {jalali_date} (vs. {prev_label})"