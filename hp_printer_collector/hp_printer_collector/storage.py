"""
Historical data storage for HP Printer Collector.

Appends one row per printer per run to a CSV file so you can track
page-count trends and toner consumption over time.  The file is created
with a header row on first write; subsequent runs append without
re-writing the header.
"""

import csv
import logging
import os
from datetime import datetime
from typing import List

logger = logging.getLogger(__name__)

# Ordered column names for the CSV file
CSV_COLUMNS = [
    "timestamp",
    "printer_name",
    "printer_ip",
    "page_count",
    "toner_black",
    "toner_cyan",
    "toner_yellow",
    "toner_magenta",
    "error",
]


def _file_has_header(path: str) -> bool:
    """Return True if the CSV file already contains a header row."""
    if not os.path.exists(path):
        return False
    try:
        with open(path, newline="", encoding="utf-8") as fh:
            first_line = fh.readline().strip()
            return first_line.startswith("timestamp")
    except OSError:
        return False


def save_to_csv(printer_results: List[dict], csv_path: str) -> None:
    """
    Append printer data rows to the CSV history file.

    Args:
        printer_results: List of result dicts as returned by
                         scraper.collect_printer_data().
        csv_path:        Path to the CSV file (created if absent).
    """
    if not printer_results:
        logger.debug("No printer results to save; skipping CSV write")
        return

    csv_dir = os.path.dirname(csv_path)
    if csv_dir and not os.path.exists(csv_dir):
        os.makedirs(csv_dir, exist_ok=True)

    write_header = not _file_has_header(csv_path)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        with open(csv_path, "a", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=CSV_COLUMNS, extrasaction="ignore")

            if write_header:
                writer.writeheader()
                logger.debug("Created new CSV file with header: %s", csv_path)

            for result in printer_results:
                row = {
                    "timestamp": timestamp,
                    "printer_name": result.get("name", ""),
                    "printer_ip": result.get("ip", ""),
                    "page_count": result.get("page_count", ""),
                    "toner_black": result.get("toner_black", ""),
                    "toner_cyan": result.get("toner_cyan", ""),
                    "toner_yellow": result.get("toner_yellow", ""),
                    "toner_magenta": result.get("toner_magenta", ""),
                    "error": result.get("error", ""),
                }
                writer.writerow(row)
                logger.debug("Saved row for printer '%s'", result.get("name", result.get("ip")))

        logger.info("Saved %d record(s) to %s", len(printer_results), csv_path)

    except OSError as exc:
        logger.error("Failed to write CSV file %s: %s", csv_path, exc)
