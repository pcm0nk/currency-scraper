import re
from typing import Dict, List, Optional
from playwright.sync_api import sync_playwright


PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
TIME_PATTERN = re.compile(r"^\d{1,2}:\d{2}(?::\d{2})?$")
PCT_ONLY_PATTERN = re.compile(r"^[+-]?\d+(?:\.\d+)?\s*%$")


class ServatmandiScraperModel:
    def __init__(
        self,
        url: str,
        targets: List[str],
        aliases: Optional[Dict[str, List[str]]] = None,
        multipliers: Optional[Dict[str, float]] = None,
        headless: bool = True,
        timeout_ms: int = 30000,
    ):
        self.url = url
        self.targets = targets
        self.aliases = aliases or {}
        self.multipliers = multipliers or {}
        self.headless = headless
        self.timeout_ms = timeout_ms

    @staticmethod
    def _normalize_text(text: str) -> str:
        cleaned = text.replace("\u200c", " ").replace("\u200f", " ")
        cleaned = re.sub(r"(?<!\d)\.(?!\d)", " ", cleaned)
        cleaned = re.sub(r"\.+(?=\s|$)", "", cleaned)
        return re.sub(r"\s+", " ", cleaned).strip()

    @classmethod
    def _normalize_for_match(cls, text: str) -> str:
        return cls._normalize_text(text).translate(PERSIAN_DIGITS).lower()

    @classmethod
    def _parse_number(cls, raw_text: str) -> Optional[float]:
        normalized = cls._normalize_text(raw_text).translate(PERSIAN_DIGITS)
        if not normalized or TIME_PATTERN.match(normalized) or PCT_ONLY_PATTERN.match(normalized):
            return None
        match = re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", normalized)
        if not match:
            return None
        num_str = match.group(0).replace(",", "").rstrip(".")
        if not num_str or num_str in ("+", "-"):
            return None
        try:
            return float(num_str)
        except ValueError:
            return None

    def scrape_prices(self, targets: Optional[List[str]] = None) -> Dict[str, int]:
        items_to_scrape = targets if targets is not None else self.targets
        results: Dict[str, int] = {}

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=self.headless)
            page = browser.new_page()
            page.set_default_timeout(self.timeout_ms)
            page.goto(self.url, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)

            extracted = page.evaluate(
                r"""() => {
                    const clean = (s) => (s || '').replace(/\s+/g, ' ').trim();

                    const rowsOut = [];
                    const allRows = Array.from(
                        document.querySelectorAll('tr, [role="row"], li, .price-row, .row')
                    );

                    for (const row of allRows) {
                        const fullText = clean(row.innerText);
                        if (!fullText || fullText.length > 350) continue;

                        let cellEls = Array.from(row.querySelectorAll('td, th, [role="cell"]'));
                        if (cellEls.length === 0) {
                            cellEls = Array.from(row.children).filter(el => clean(el.innerText).length > 0);
                        }
                        if (cellEls.length < 2) continue;

                        let lastValColIdx = -1;
                        const parentContainer = row.closest('table, [role="table"], [role="grid"], section, .table-wrapper') || document;
                        const headers = Array.from(
                            parentContainer.querySelectorAll('thead th, th, [role="columnheader"]')
                        );
                        headers.forEach((h, idx) => {
                            if (clean(h.innerText).includes('آخرین مقدار')) {
                                lastValColIdx = idx;
                            }
                        });

                        const cells = cellEls.map(c => clean(c.innerText));
                        rowsOut.push({
                            cells: cells,
                            fullText: fullText,
                            lastValColIdx: lastValColIdx
                        });
                    }

                    return rowsOut;
                }"""
            )

            browser.close()

        for target in items_to_scrape:
            search_terms = self.aliases.get(target) or [target]
            if target not in search_terms:
                search_terms = [target] + search_terms

            parsed_val: Optional[float] = None

            for term in search_terms:
                norm_term = self._normalize_for_match(term)

                for row in extracted:
                    norm_full = self._normalize_for_match(row["fullText"])
                    if norm_term not in norm_full:
                        continue

                    raw_cells = [self._normalize_text(c) for c in row["cells"]]
                    col_idx = row.get("lastValColIdx", -1)

                    if 0 <= col_idx < len(raw_cells):
                        val = self._parse_number(raw_cells[col_idx])
                        if val is not None:
                            parsed_val = val
                            break

                    label_indices = {
                        i
                        for i, c in enumerate(raw_cells)
                        if norm_term in self._normalize_for_match(c)
                    }

                    for i, c in enumerate(raw_cells):
                        if i in label_indices:
                            continue
                        val = self._parse_number(c)
                        if val is not None:
                            parsed_val = val
                            break

                    if parsed_val is not None:
                        break

                if parsed_val is not None:
                    break

            if parsed_val is not None:
                factor = self.multipliers.get(target, 1.0)
                results[target] = int(round(parsed_val * factor))
            else:
                print(f"  [Warning] Could not find 'آخرین مقدار' for: {target}")

        return results