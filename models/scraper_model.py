import re
from typing import Dict, List, Optional
from playwright.sync_api import sync_playwright


PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


class ServatmandiScraperModel:
    def __init__(self, url: str, targets: List[str], headless: bool = True, timeout_ms: int = 30000):
        self.url = url
        self.targets = targets
        self.headless = headless
        self.timeout_ms = timeout_ms

    @staticmethod
    def _normalize_text(text: str) -> str:
        cleaned = text.replace("\u200c", " ").replace("\u200f", " ")
        return re.sub(r"\s+", " ", cleaned).strip()

    @classmethod
    def _parse_number(cls, raw_text: str) -> Optional[float]:
        normalized = cls._normalize_text(raw_text).translate(PERSIAN_DIGITS)
        match = re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", normalized)
        if not match:
            return None
        num_str = match.group(0).replace(",", "")
        try:
            return float(num_str)
        except ValueError:
            return None

    def scrape_prices(self, targets: Optional[List[str]] = None) -> Dict[str, float]:
        items_to_scrape = targets if targets is not None else self.targets
        results: Dict[str, float] = {}

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=self.headless)
            page = browser.new_page()
            page.set_default_timeout(self.timeout_ms)
            page.goto(self.url, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)

            extracted = page.evaluate(
                """() => {
                    const rows = Array.from(document.querySelectorAll('tr, .price-row, .row, li, div'));
                    const pairs = [];
                    for (const el of rows) {
                        if (el.children.length >= 2 && el.children.length <= 8) {
                            const texts = Array.from(el.children)
                                .map(c => (c.innerText || '').trim())
                                .filter(Boolean);
                            if (texts.length >= 2) {
                                pairs.push(texts);
                            }
                        }
                    }
                    return pairs;
                }"""
            )

            browser.close()

        normalized_pairs = [
            [self._normalize_text(t) for t in row if self._normalize_text(t)]
            for row in extracted
        ]

        for target in items_to_scrape:
            norm_target = self._normalize_text(target)
            parsed_val: Optional[float] = None

            for row_texts in normalized_pairs:
                if any(norm_target in cell or cell in norm_target for cell in row_texts):
                    for cell in reversed(row_texts):
                        if cell == norm_target:
                            continue
                        val = self._parse_number(cell)
                        if val is not None:
                            parsed_val = val
                            break
                if parsed_val is not None:
                    break

            if parsed_val is not None:
                results[target] = parsed_val

        return results