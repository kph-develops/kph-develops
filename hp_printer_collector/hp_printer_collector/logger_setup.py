"""
Logging configuration for HP Printer Collector.

Sets up a logger that writes to both a rotating file and the console,
so operators can monitor live runs while retaining history on disk.
"""

import logging
import logging.handlers
import os
import sys
from typing import Optional


def setup_logger(
    name: str = "hp_printer_collector",
    log_file: str = "printer_collector.log",
    level: str = "INFO",
    max_bytes: int = 5 * 1024 * 1024,  # 5 MB per file
    backup_count: int = 5,
) -> logging.Logger:
    """
    Configure and return the application logger.

    Creates a RotatingFileHandler so logs don't grow unbounded, plus a
    StreamHandler so output is visible when running interactively or
    in a Task Scheduler / cron console.

    Args:
        name:         Logger name (used by child modules via logging.getLogger(__name__)).
        log_file:     Path to the log file (absolute or relative to CWD).
        level:        Minimum log level string: DEBUG, INFO, WARNING, ERROR, CRITICAL.
        max_bytes:    Maximum size of a single log file before rotation.
        backup_count: Number of rotated backup files to keep.

    Returns:
        Configured logging.Logger instance.
    """
    numeric_level = getattr(logging, level.upper(), logging.INFO)

    logger = logging.getLogger(name)
    logger.setLevel(numeric_level)

    # Avoid adding duplicate handlers if setup_logger is called more than once
    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        fmt="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # --- Rotating file handler ---
    log_dir = os.path.dirname(log_file)
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir, exist_ok=True)

    file_handler = logging.handlers.RotatingFileHandler(
        log_file,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    file_handler.setLevel(numeric_level)

    # --- Console (stdout) handler ---
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    console_handler.setLevel(numeric_level)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger
