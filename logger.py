"""
logger.py
---------

Centralized logging configuration for the Test Bench Software.

This module is the ONLY place where logging is configured.
All other modules must obtain loggers via `get_logger(__name__)`.

Design principles:
- Single application-wide log file
- Safe to import multiple times (no duplicate handlers)
- No UI coupling
- No business logic inside logging
- Path Anchoring: Logs are stored relative to this file, not the CWD
"""

import sys
import logging
import os
from logging.handlers import RotatingFileHandler
from typing import Optional

# =============================================================================
# Configuration Constants
# =============================================================================

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LOG_DIR = os.path.join(BASE_DIR, "logs")
LOG_FILE_NAME = "testbench.log"
LOG_FILE_PATH = os.path.join(LOG_DIR, LOG_FILE_NAME)

DEFAULT_LOG_LEVEL = logging.INFO
MAX_LOG_SIZE_BYTES = 5 * 1024 * 1024   # 5 MB
BACKUP_COUNT = 3                       # testbench.log.1, .2, .3

# =============================================================================
# Internal State
# =============================================================================

_logger_initialized = False


# =============================================================================
# Logger Setup
# =============================================================================

def _initialize_logging(log_level: int = DEFAULT_LOG_LEVEL) -> None:
    """
    Initializes the global logging configuration.

    This function is intentionally private and guarded to ensure
    logging is configured exactly once.
    """
    global _logger_initialized

    if _logger_initialized:
        return

    # Ensure log directory exists
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
    except OSError as e:
        # Fallback to console only if file creation fails (e.g. permissions)
        print(f"CRITICAL: Failed to create log directory at {LOG_DIR}: {e}")
        # We continue so at least console logging works

    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)

    # Formatter (timestamped, readable, developer-focused)
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # File Handler (rotating)
    try:
        file_handler = RotatingFileHandler(
            LOG_FILE_PATH,
            maxBytes=MAX_LOG_SIZE_BYTES,
            backupCount=BACKUP_COUNT,
            encoding="utf-8"
        )
        file_handler.setLevel(log_level)
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
    except OSError as e:
         print(f"CRITICAL: Failed to initialize file logging at {LOG_FILE_PATH}: {e}")

    # Console Handler (developer convenience)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)

    # Attach handlers
    root_logger.addHandler(console_handler)

    _logger_initialized = True


# =============================================================================
# Public API
# =============================================================================

def get_logger(name: Optional[str] = None) -> logging.Logger:
    """
    Returns a configured logger instance.

    This is the ONLY function that should be used to obtain loggers
    throughout the application.

    Usage:
        logger = get_logger(__name__)
        logger.info("Something happened")

    :param name: Logger name (usually __name__)
    :return: logging.Logger
    """
    _initialize_logging()
    return logging.getLogger(name)