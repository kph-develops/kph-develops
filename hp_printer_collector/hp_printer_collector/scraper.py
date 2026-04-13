"""
HP Printer web-interface scraper.

Fetches the Usage Page and Supplies Status page from an HP printer's
embedded web server, then parses the relevant values from the HTML.

HP printers use non-standard element IDs that contain dots and dashes
(e.g. "UsagePage.EquivalentImpressionsTable.Total.Total").  BeautifulSoup
handles these correctly when you pass the id value as a dict attribute
rather than as a keyword argument to .find(), which avoids any CSS-class
ambiguity with dot-separated strings.
"""

import logging
import re
from typing import Optional
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

# HP printer endpoints
USAGE_ENDPOINT = "/hp/device/InternalPages/Index?id=UsagePage"
SUPPLIES_ENDPOINT = "/hp/device/InternalPages/Index?id=SuppliesStatus"

# Element IDs for toner levels on the Supplies Status page
TONER_IDS = {
    "black": "BlackCartridge1-Header_Level",
    "cyan": "CyanCartridge1-Header_Level",
    "yellow": "YellowCartridge1-Header_Level",
    # Magenta uses a section ID; percentage is extracted via regex fallback
    "magenta_section": "MagentaCartridge1-Header",
    "magenta_level": "MagentaCartridge1-Header_Level",
}

# Regex that matches percentage values, including the "<10%" edge case
PERCENT_RE = re.compile(r"<?\d+\s*%")


# ---------------------------------------------------------------------------
# Low-level HTTP helpers
# ---------------------------------------------------------------------------


def _build_url(ip: str, endpoint: str) -> str:
    """Construct the full URL for a printer endpoint."""
    base = f"http://{ip}"
    return urljoin(base, endpoint)


def fetch_page(ip: str, endpoint: str, timeout: int = 15) -> Optional[str]:
    """
    Perform an HTTP GET against the printer and return the response body.

    Disables SSL verification because most HP embedded web servers use
    self-signed certificates.  The verify=False warning is suppressed via
    urllib3 so it does not flood the log.

    Returns None on any network or HTTP error, after logging the details.
    """
    import urllib3

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    url = _build_url(ip, endpoint)
    logger.debug("Fetching %s", url)

    try:
        response = requests.get(
            url,
            timeout=timeout,
            verify=False,
            headers={"User-Agent": "HPPrinterCollector/1.0"},
        )
        response.raise_for_status()
        logger.debug("Received %d bytes from %s", len(response.content), url)
        return response.text

    except requests.exceptions.ConnectionError as exc:
        logger.error("Cannot connect to printer at %s: %s", ip, exc)
    except requests.exceptions.Timeout:
        logger.error("Timed out connecting to printer at %s (timeout=%ds)", ip, timeout)
    except requests.exceptions.HTTPError as exc:
        logger.error("HTTP error from %s: %s", url, exc)
    except requests.exceptions.RequestException as exc:
        logger.error("Unexpected request error for %s: %s", url, exc)

    return None


# ---------------------------------------------------------------------------
# HTML parsing helpers
# ---------------------------------------------------------------------------


def _find_by_id(soup: BeautifulSoup, element_id: str):
    """
    Locate an element by its exact HTML id attribute.

    Using attrs={'id': ...} avoids BeautifulSoup treating dots inside the
    id string as CSS class separators.
    """
    return soup.find(attrs={"id": element_id})


def _clean_percent(raw: str) -> Optional[str]:
    """
    Normalise a toner percentage string.

    - Strips asterisks used by HP as footnote markers.
    - Preserves the '<' prefix for sub-threshold values (e.g. '<10%').
    - Returns None if no percentage pattern is found.
    """
    cleaned = raw.replace("*", "").strip()

    # If the whole string looks like a percentage, return it directly
    if PERCENT_RE.fullmatch(cleaned):
        return cleaned

    # Otherwise try to extract a percentage substring
    match = PERCENT_RE.search(cleaned)
    if match:
        return match.group(0).strip()

    logger.warning("Could not extract percentage from: %r", raw)
    return None


# ---------------------------------------------------------------------------
# Page-count parser (Usage Page)
# ---------------------------------------------------------------------------


def parse_page_count(html: str) -> Optional[int]:
    """
    Extract the total page count from the Usage Page HTML.

    Target element:
        <td id="UsagePage.EquivalentImpressionsTable.Total.Total">12,345</td>

    Commas in the number are removed before conversion to int.
    """
    soup = BeautifulSoup(html, "html.parser")
    element = _find_by_id(soup, "UsagePage.EquivalentImpressionsTable.Total.Total")

    if element is None:
        logger.warning("Page count element not found in Usage Page HTML")
        return None

    raw = element.get_text(strip=True)
    logger.debug("Raw page count text: %r", raw)

    numeric_str = raw.replace(",", "").replace(".", "").strip()
    try:
        return int(numeric_str)
    except ValueError:
        logger.warning("Could not convert page count %r to integer", raw)
        return None


# ---------------------------------------------------------------------------
# Toner-level parsers (Supplies Status page)
# ---------------------------------------------------------------------------


def _parse_simple_toner(soup: BeautifulSoup, element_id: str) -> Optional[str]:
    """Parse a toner level from a direct _Level element."""
    element = _find_by_id(soup, element_id)
    if element is None:
        logger.warning("Toner element %r not found", element_id)
        return None

    raw = element.get_text(strip=True)
    logger.debug("Raw toner text for %s: %r", element_id, raw)
    return _clean_percent(raw)


def _parse_magenta(soup: BeautifulSoup) -> Optional[str]:
    """
    Extract the magenta toner percentage.

    Tries the direct _Level element first; if absent, falls back to
    searching inside the parent section identified by MagentaCartridge1-Header.
    This handles HP firmware variants where the level is embedded as text
    within the header section rather than in a dedicated child element.
    """
    # Preferred: dedicated level element (same pattern as other colours)
    level = _parse_simple_toner(soup, TONER_IDS["magenta_level"])
    if level is not None:
        return level

    # Fallback: scan the section element for any percentage string
    section = _find_by_id(soup, TONER_IDS["magenta_section"])
    if section is None:
        logger.warning("Magenta cartridge section %r not found", TONER_IDS["magenta_section"])
        return None

    text = section.get_text(separator=" ", strip=True)
    logger.debug("Magenta section text: %r", text)

    # Remove asterisks before searching
    cleaned = text.replace("*", "")
    match = PERCENT_RE.search(cleaned)
    if match:
        return match.group(0).strip()

    logger.warning("No percentage found in magenta section: %r", text)
    return None


def parse_toner_levels(html: str) -> dict:
    """
    Parse all four toner percentages from the Supplies Status HTML.

    Returns a dict with keys: black, cyan, yellow, magenta.
    Each value is a string like '75%', '<10%', or None if parsing failed.
    """
    soup = BeautifulSoup(html, "html.parser")

    return {
        "black": _parse_simple_toner(soup, TONER_IDS["black"]),
        "cyan": _parse_simple_toner(soup, TONER_IDS["cyan"]),
        "yellow": _parse_simple_toner(soup, TONER_IDS["yellow"]),
        "magenta": _parse_magenta(soup),
    }


# ---------------------------------------------------------------------------
# Top-level collector
# ---------------------------------------------------------------------------


def collect_printer_data(printer: dict, timeout: int = 15) -> dict:
    """
    Collect all usage data for a single printer configuration entry.

    Args:
        printer: Dict with at minimum 'ip', optionally 'name'.
        timeout: Per-request HTTP timeout in seconds.

    Returns:
        Dict with keys:
            name, ip, page_count,
            toner_black, toner_cyan, toner_yellow, toner_magenta,
            error (None on success, error message string on failure).
    """
    ip = printer["ip"]
    name = printer.get("name", ip)

    result = {
        "name": name,
        "ip": ip,
        "page_count": None,
        "toner_black": None,
        "toner_cyan": None,
        "toner_yellow": None,
        "toner_magenta": None,
        "error": None,
    }

    logger.info("Collecting data from printer '%s' (%s)", name, ip)

    # --- Usage page (page count) ---
    usage_html = fetch_page(ip, USAGE_ENDPOINT, timeout=timeout)
    if usage_html is None:
        result["error"] = f"Failed to fetch usage page from {ip}"
        logger.error(result["error"])
        return result

    result["page_count"] = parse_page_count(usage_html)
    if result["page_count"] is None:
        logger.warning("Page count unavailable for printer '%s'", name)

    # --- Supplies page (toner levels) ---
    supplies_html = fetch_page(ip, SUPPLIES_ENDPOINT, timeout=timeout)
    if supplies_html is None:
        result["error"] = f"Failed to fetch supplies page from {ip}"
        logger.error(result["error"])
        return result

    toner = parse_toner_levels(supplies_html)
    result["toner_black"] = toner["black"]
    result["toner_cyan"] = toner["cyan"]
    result["toner_yellow"] = toner["yellow"]
    result["toner_magenta"] = toner["magenta"]

    logger.info(
        "Collected data for '%s': pages=%s  B=%s  C=%s  Y=%s  M=%s",
        name,
        result["page_count"],
        result["toner_black"],
        result["toner_cyan"],
        result["toner_yellow"],
        result["toner_magenta"],
    )

    return result
