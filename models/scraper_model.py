import re
from typing import Dict, List, Optional
from playwright.sync_api import sync_playwright


PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
TIME_PATTERN = re.compile(r"^\d{1,2}:\d{2}(?::\d{2})?$")
PCT_ONLY_PATTERN = re.compile(r"^[+-]?\d+(?:\.\d+)?\s*%$")

EXCLUDED_TERMS_BY_TARGET: Dict[str, List[str]] = {
    "دلار آمریکا": ["حواله دولتی"],
}


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
        if (
            not normalized
            or TIME_PATTERN.match(normalized)
            or PCT_ONLY_PATTERN.match(normalized)
        ):
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

    def _is_valid_label_text(
        self, norm_text: str, norm_term: str, excluded_terms: List[str]
    ) -> bool:
        if not norm_text or norm_term not in norm_text:
            return False
        if any(ex and ex in norm_text for ex in excluded_terms):
            return False
        return True

    def scrape_prices(self, targets: Optional[List[str]] = None) -> Dict[str, float]:
        items_to_scrape = targets if targets is not None else self.targets
        results: Dict[str, float] = {}

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=self.headless)
            page = browser.new_page()
            page.set_default_timeout(self.timeout_ms)
            page.goto(self.url, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)

            # Scroll through the page so lower sections/watchlists load and populate
            page.evaluate(
                r"""async () => {
                    const step = Math.max(400, Math.floor(window.innerHeight * 0.8));
                    let y = 0;
                    const maxScroll = Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);
                    while (y < maxScroll) {
                        y += step;
                        window.scrollTo(0, y);
                        await new Promise(r => setTimeout(r, 250));
                    }
                    window.scrollTo(0, 0);
                }"""
            )
            page.wait_for_timeout(1500)

            extracted = page.evaluate(
                r"""() => {
                    const clean = (s) => (s || '').replace(/\s+/g, ' ').trim();
                    const isVisible = (el) => {
                        const style = window.getComputedStyle(el);
                        if (!style || style.display === 'none' || style.visibility === 'hidden' || parseFloat(style.opacity) === 0) {
                            return false;
                        }
                        const rect = el.getBoundingClientRect();
                        return rect.width > 0 && rect.height > 0;
                    };

                    const rowsOut = [];
                    const seenKeys = new Set();
                    const tables = Array.from(document.querySelectorAll('table'));

                    for (const tbl of tables) {
                        if (!isVisible(tbl)) continue;
                        const headers = Array.from(tbl.querySelectorAll('thead th, th')).filter(isVisible);
                        let lastValColIdx = -1;
                        headers.forEach((h, idx) => {
                            const hText = clean(h.innerText);
                            if (hText.includes('آخرین مقدار') || hText.includes('آخرین قیمت')) {
                                lastValColIdx = idx;
                            }
                        });

                        const trs = Array.from(tbl.querySelectorAll('tbody tr, tr'));
                        for (const tr of trs) {
                            if (!isVisible(tr)) continue;
                            const fullText = clean(tr.innerText);
                            if (!fullText) continue;

                            const nameEls = Array.from(tr.querySelectorAll('.name a, .name, a'));
                            const nameTexts = nameEls
                                .map(el => clean(el.innerText))
                                .filter(Boolean);

                            let cellEls = Array.from(tr.querySelectorAll('th, td')).filter(isVisible);
                            if (cellEls.length === 0) {
                                cellEls = Array.from(tr.querySelectorAll('th, td'));
                            }
                            const cells = cellEls.map(c => clean(c.innerText));

                            const key = `${fullText}||${cells.join('|')}`;
                            if (seenKeys.has(key)) continue;
                            seenKeys.add(key);

                            rowsOut.push({
                                nameTexts: nameTexts,
                                cells: cells,
                                fullText: fullText,
                                lastValColIdx: lastValColIdx
                            });
                        }
                    }

                    // Also capture visible non-table row/card containers
                    const cardCandidates = Array.from(
                        document.querySelectorAll('[role="row"], li, .price-row, .row, div')
                    );
                    for (const el of cardCandidates) {
                        if (!isVisible(el)) continue;
                        const fullText = clean(el.innerText);
                        if (!fullText || fullText.length > 350) continue;

                        const childCells = Array.from(el.children).filter(
                            c => isVisible(c) && clean(c.innerText).length > 0
                        );
                        if (childCells.length < 2 || childCells.length > 10) continue;

                        const cells = childCells.map(c => clean(c.innerText));
                        const key = `${fullText}||${cells.join('|')}`;
                        if (seenKeys.has(key)) continue;
                        seenKeys.add(key);

                        rowsOut.push({
                            nameTexts: cells,
                            cells: cells,
                            fullText: fullText,
                            lastValColIdx: -1
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

            excluded_terms = [
                self._normalize_for_match(ex)
                for ex in EXCLUDED_TERMS_BY_TARGET.get(target, [])
            ]

            parsed_val: Optional[float] = None

            for term in search_terms:
                norm_term = self._normalize_for_match(term)

                for row in extracted:
                    norm_full = self._normalize_for_match(row["fullText"])
                    if norm_term not in norm_full:
                        continue
                    if any(ex and ex in norm_full for ex in excluded_terms):
                        continue

                    raw_cells = [self._normalize_text(c) for c in row["cells"]]
                    norm_cells = [self._normalize_for_match(c) for c in raw_cells]

                    has_label = any(
                        self._is_valid_label_text(nc, norm_term, excluded_terms)
                        for nc in norm_cells
                    ) or any(
                        self._is_valid_label_text(self._normalize_for_match(n), norm_term, excluded_terms)
                        for n in row.get("nameTexts", [])
                    )
                    if not has_label:
                        continue

                    label_cell_indices = {
                        i
                        for i, nc in enumerate(norm_cells)
                        if norm_term in nc
                    }

                    col_idx = row.get("lastValColIdx", -1)

                    # 1. Use 'آخرین مقدار' / 'آخرین قیمت' column index if valid and not a label cell
                    if 0 <= col_idx < len(raw_cells) and col_idx not in label_cell_indices:
                        val = self._parse_number(raw_cells[col_idx])
                        if val is not None:
                            parsed_val = val
                            break

                    # 2. First valid price cell after the label cell
                    max_label_idx = max(label_cell_indices) if label_cell_indices else -1
                    for i in range(max_label_idx + 1, len(raw_cells)):
                        if i in label_cell_indices:
                            continue
                        val = self._parse_number(raw_cells[i])
                        if val is not None:
                            parsed_val = val
                            break

                    # 3. Fallback across any other non-label cell in the row
                    if parsed_val is None:
                        for i, c in enumerate(raw_cells):
                            if i in label_cell_indices:
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
                results[target] = parsed_val * factor
            else:
                print(f"  [Warning] Could not find 'آخرین مقدار' for: {target}")

        return results