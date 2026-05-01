#!/usr/bin/env python3
"""
HP Printer Usage Collector – main entry point.

Usage
-----
Run manually (collect now and email):
    python main.py

Specify an alternate config file:
    python main.py --config /path/to/config.yaml

Collect data only (no email):
    python main.py --no-email

Schedule via cron (Linux/macOS/WSL) – 1st of every month at 07:00:
    0 7 1 * * /path/to/venv/bin/python /path/to/hp_printer_collector/main.py

Schedule via Windows Task Scheduler – see README.md for GUI instructions.

Exit codes
----------
0  – success (all printers collected, email sent)
1  – partial failure (some printers failed or email not sent)
2  – configuration error (bad file, missing required keys)
"""

import argparse
import logging
import os
import sys
from typing import Optional

import yaml

from hp_printer_collector.logger_setup import setup_logger
from hp_printer_collector.scraper import collect_printer_data
from hp_printer_collector.email_reporter import send_report, test_smtp_connection
from hp_printer_collector.storage import save_to_csv

# Module-level logger; configured properly after config is loaded.
logger = logging.getLogger("hp_printer_collector")

DEFAULT_CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.yaml")


# ---------------------------------------------------------------------------
# Config loader
# ---------------------------------------------------------------------------


def load_config(path: str) -> dict:
    """
    Load and validate the YAML configuration file.

    Raises SystemExit(2) if the file is missing or lacks required keys.
    """
    if not os.path.exists(path):
        print(f"ERROR: Config file not found: {path}", file=sys.stderr)
        print("Copy config.example.yaml to config.yaml and fill in your settings.", file=sys.stderr)
        sys.exit(2)

    with open(path, encoding="utf-8") as fh:
        try:
            cfg = yaml.safe_load(fh)
        except yaml.YAMLError as exc:
            print(f"ERROR: Could not parse config file: {exc}", file=sys.stderr)
            sys.exit(2)

    if not cfg:
        print("ERROR: Config file is empty.", file=sys.stderr)
        sys.exit(2)

    # Required: at least one printer
    printers = cfg.get("printers")
    if not printers or not isinstance(printers, list):
        print("ERROR: 'printers' list is missing or empty in config.", file=sys.stderr)
        sys.exit(2)

    for idx, printer in enumerate(printers):
        if "ip" not in printer:
            print(f"ERROR: Printer #{idx + 1} in config is missing the 'ip' field.", file=sys.stderr)
            sys.exit(2)

    return cfg


# ---------------------------------------------------------------------------
# CLI argument parsing
# ---------------------------------------------------------------------------


def parse_args(argv=None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="hp_printer_collector",
        description="Collect HP printer usage data and send a monthly email report.",
    )
    parser.add_argument(
        "--config",
        default=DEFAULT_CONFIG_PATH,
        metavar="PATH",
        help=f"Path to the YAML config file (default: {DEFAULT_CONFIG_PATH})",
    )
    parser.add_argument(
        "--no-email",
        action="store_true",
        help="Collect data and save to CSV but do not send an email.",
    )
    parser.add_argument(
        "--no-csv",
        action="store_true",
        help="Do not append results to the CSV history file.",
    )
    parser.add_argument(
        "--test-smtp",
        action="store_true",
        help="Test SMTP credentials and exit without collecting printer data.",
    )
    return parser.parse_args(argv)


# ---------------------------------------------------------------------------
# Main logic
# ---------------------------------------------------------------------------


def run(config_path: str, send_email: bool = True, write_csv: bool = True) -> int:
    """
    Core execution function.

    Returns an exit code: 0 = full success, 1 = partial failure.
    """
    cfg = load_config(config_path)

    # --- Logging setup ---
    log_cfg = cfg.get("logging", {})
    log_file = log_cfg.get("file", "printer_collector.log")
    log_level = log_cfg.get("level", "INFO")
    setup_logger(log_file=log_file, level=log_level)

    logger.info("HP Printer Collector starting")
    logger.info("Config: %s", config_path)

    printers = cfg["printers"]
    timeout = cfg.get("timeout", 15)
    threshold = cfg.get("alerts", {}).get("toner_low_threshold", 20)

    # --- Collect data from all printers ---
    results = []
    all_success = True

    for printer in printers:
        result = collect_printer_data(printer, timeout=timeout)
        results.append(result)
        if result.get("error"):
            all_success = False

    # --- Save to CSV ---
    if write_csv:
        csv_path = cfg.get("storage", {}).get("csv_file", "printer_history.csv")
        save_to_csv(results, csv_path)

    # --- Send email ---
    email_sent = True
    if send_email:
        smtp_cfg = cfg.get("smtp", {})
        email_cfg = cfg.get("email", {})
        recipients = email_cfg.get("recipients", [])
        subject = email_cfg.get("subject")

        if not smtp_cfg:
            logger.warning("No SMTP configuration found; skipping email")
            email_sent = False
        elif not recipients:
            logger.warning("No email recipients configured; skipping email")
            email_sent = False
        else:
            email_sent = send_report(
                printer_results=results,
                smtp_cfg=smtp_cfg,
                recipients=recipients,
                subject=subject,
                threshold=threshold,
            )
    else:
        logger.info("Email sending skipped (--no-email flag)")

    # --- Summary ---
    failed = [r for r in results if r.get("error")]
    logger.info(
        "Done. Printers: %d total, %d failed. Email sent: %s",
        len(results),
        len(failed),
        email_sent,
    )

    if failed or not email_sent:
        return 1
    return 0


def main() -> None:
    args = parse_args()

    if args.test_smtp:
        cfg = load_config(args.config)
        log_cfg = cfg.get("logging", {})
        setup_logger(log_file=log_cfg.get("file", "printer_collector.log"), level="DEBUG")
        test_smtp_connection(cfg.get("smtp", {}))
        sys.exit(0)

    exit_code = run(
        config_path=args.config,
        send_email=not args.no_email,
        write_csv=not args.no_csv,
    )
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
