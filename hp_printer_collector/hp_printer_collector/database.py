"""
SQLite snapshot store for HP Printer Collector.

Saves one row per printer per successful run so the monthly page-count
delta can be calculated as:

    monthly_pages = current_total - previous_run_total

The database is created automatically on first use.  No third-party
packages are required — sqlite3 is part of the Python standard library.

Workflow (enforced by main.py):
    1. init_db()                        -- create table if absent
    2. get_previous_snapshot(ip)        -- read BEFORE saving current run
    3. ... calculate deltas ...
    4. save_snapshots(results)          -- write current run
    5. ... build and send report ...
"""

import logging
import os
import sqlite3
from datetime import datetime
from typing import List, Optional

logger = logging.getLogger(__name__)

_DDL = """
CREATE TABLE IF NOT EXISTS snapshots (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    run_date          TEXT    NOT NULL,
    printer_ip        TEXT    NOT NULL,
    printer_name      TEXT,
    page_count        INTEGER,
    page_count_color  INTEGER,
    page_count_mono   INTEGER
);
CREATE INDEX IF NOT EXISTS idx_snapshots_ip_date
    ON snapshots (printer_ip, run_date DESC);
"""


def _connect(db_path: str) -> sqlite3.Connection:
    db_dir = os.path.dirname(os.path.abspath(db_path))
    os.makedirs(db_dir, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path: str) -> None:
    """Create the snapshots table and index if they do not already exist."""
    with _connect(db_path) as conn:
        conn.executescript(_DDL)
    logger.debug("Database ready: %s", db_path)


def get_previous_snapshot(printer_ip: str, db_path: str) -> Optional[dict]:
    """
    Return the most recent snapshot row for this printer IP, or None when
    no prior run has been recorded (i.e. this is the first ever run).

    Must be called BEFORE save_snapshots() so the current run is not yet
    in the database and cannot be mistaken for the previous one.
    """
    with _connect(db_path) as conn:
        row = conn.execute(
            """
            SELECT * FROM snapshots
            WHERE  printer_ip = ?
            ORDER  BY run_date DESC
            LIMIT  1
            """,
            (printer_ip,),
        ).fetchone()
    if row:
        logger.debug("Previous snapshot for %s: date=%s total=%s", printer_ip, row["run_date"], row["page_count"])
    else:
        logger.debug("No previous snapshot found for %s (first run)", printer_ip)
    return dict(row) if row else None


def save_snapshots(results: List[dict], db_path: str) -> None:
    """
    Insert one row per successful printer result into the snapshots table.
    Results where result['error'] is set are skipped — a failed collection
    should not overwrite the last good baseline.
    """
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rows = [
        (
            run_date,
            r["ip"],
            r.get("name", ""),
            r.get("page_count"),
            r.get("page_count_color"),
            r.get("page_count_mono"),
        )
        for r in results
        if not r.get("error") and r.get("page_count") is not None
    ]

    if not rows:
        logger.warning("No successful page counts to snapshot; skipping DB write")
        return

    with _connect(db_path) as conn:
        conn.executemany(
            """
            INSERT INTO snapshots
                (run_date, printer_ip, printer_name,
                 page_count, page_count_color, page_count_mono)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            rows,
        )
    logger.info("Saved %d snapshot(s) to %s", len(rows), db_path)
