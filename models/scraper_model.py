import re
from typing import Dict, List
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


PERSIAN_TO_ENGLISH_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def parse_market_number(raw_text: str) -> float:
    """
    Converts Persian/Arabic digits and formatted strings (e.g., '۴,۵۲۰,۰۰۰')
    into a standard float.
    """
    if not raw_text:
        raise ValueError("Empty value cannot be parsed as a number.")

    normalized = raw_text.translate(PERSIAN_TO_ENGLISH_DIGITS)
    # Remove thousands commas, spaces, and currency/symbol noise
    normalized = normalized.replace(",", "").replace("٬", "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", normalized)
    if not match:
        raise ValueError(f"Could not extract numeric value from: '{raw_text}'")
    return float(match.group(0))


class ServatmandiScraperModel:
    def __init__(self, url: str, targets: List[str], timeout_ms: int = 30000):
        self.url = url
        self.targets = targets
        self.timeout_ms = timeout_ms

    def fetch_prices(self) -> Dict[str, float]:
        """
        Loads the homepage and extracts the 'آخرین مقدار' (latest value)
        for each target item.
        """
        results: Dict[str, float] = {}

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                locale="fa-IR",
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/131.0.0.0 Safari/537.36"
                ),
            )
            page = context.new_page()

            try:
                page.goto(self.url, wait_until="networkidle", timeout=self.timeout_ms)
                # Wait until at least the first target label is rendered in DOM
                page.get_by_text(self.targets[0], exact=False).first.wait_for(
                    timeout=self.timeout_ms
                )
                # Small buffer for live websocket/XHR price ticks to populate
                page.wait_for_timeout(1500)

                for item_name in self.targets:
                    raw_price = self._extract_item_latest_value(page, item_name)
                    results[item_name] = parse_market_number(raw_price)

            except PlaywrightTimeoutError as exc:
                raise RuntimeError(f"Timed out loading {self.url}: {exc}") from exc
            finally:
                browser.close()

        return results

    def _extract_item_latest_value(self, page, item_name: str) -> str:
        """
        Finds the smallest ancestor container around `item_name` that includes
        numeric price text, then returns the first numeric value after the label.
        """
        label_locator = page.get_by_text(item_name, exact=True).first
        label_locator.wait_for(timeout=self.timeout_ms)

        # Check ancestors from closest parent upward to find the row/card holding the price
        ancestors = label_locator.locator("xpath=ancestor::*")
        count = ancestors.count()

        for i in range(count - 1, -1, -1):
            candidate = ancestors.nth(i)
            try:
                text = candidate.inner_text().strip()
            except Exception:
                continue

            if not text or len(text) > 250:
                continue

            if item_name in text and re.search(r"[0-9۰-۹]", text):
                after_label = text.split(item_name, 1)[-1]
                tokens = re.findall(r"[0-9۰-۹][0-9۰-۹,٬.]*", after_label)
                if tokens:
                    return tokens[0]

        raise LookupError(f"Could not find latest price cell for '{item_name}'.")