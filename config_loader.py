"""
config_loader.py
----------------
Centralized configuration loader, validator, and writer for the AVR Test Bench application.

Responsibilities:
- Load JSON configuration from disk
- Validate mandatory configuration structure
- Normalize paths and values
- Fail fast with clear errors for invalid configuration
- Persist runtime configuration changes back to disk (Path Anchored)

This module MUST NOT contain UI logic or application state.
"""

import sys
import json
import os
from typing import Dict

from logger import get_logger

logger = get_logger(__name__)

# =============================================================================
# Configuration Constants (Path Anchoring)
# =============================================================================
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_CONFIG_PATH = os.path.join(BASE_DIR, "config.json")


# =============================================================================
# Custom Exception
# =============================================================================
class ConfigError(Exception):
    """Raised when configuration is missing, invalid, or malformed."""
    pass


# =============================================================================
# Public API
# =============================================================================
def load_config(path: str = None) -> Dict:
    """
    Loads and validates the application configuration.

    :param path: Path to config.json. If None, defaults to the 'config.json'
                 located in the same directory as this script.
    :return: Validated configuration dictionary
    :raises ConfigError: if config is missing or invalid
    """
    if path is None:
        path = DEFAULT_CONFIG_PATH

    logger.info("Loading configuration file: %s", path)

    if not os.path.exists(path):
        logger.critical("Config file not found: %s", path)
        raise ConfigError(f"Config file not found: {path}")

    try:
        with open(path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except json.JSONDecodeError as e:
        logger.critical("Invalid JSON in config file: %s", e)
        raise ConfigError(f"Invalid JSON in config file: {e}")

    _validate_config(config)
    _normalize_config(config)

    logger.info("Configuration loaded successfully")
    return config


def save_config(config: Dict, path: str = None) -> None:
    """
    Saves the configuration dictionary to disk.
    
    This function should be called whenever settings are modified at runtime
    (e.g., via the Settings Dialog) to ensure persistence across restarts.

    :param config: The configuration dictionary to save.
    :param path: Path to config.json. If None, defaults to the 'config.json'
                 located in the same directory as this script.
    :raises IOError: If writing to the file fails.
    """
    if path is None:
        path = DEFAULT_CONFIG_PATH

    try:
        # Validate before saving to prevent writing corrupt config
        _validate_config(config)
        
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
        
        logger.info("Configuration saved successfully to %s", path)
        
    except Exception as e:
        logger.error("Failed to save configuration to %s: %s", path, e)
        raise


# =============================================================================
# Validation Helpers
# =============================================================================
def _validate_config(config: Dict) -> None:
    """
    Validates the structure of the configuration dictionary.
    Raises ConfigError on missing or invalid fields.
    """

    required_top_level_keys = [
        "app_name",
        "version",
        "site",
        "reports",
        "meter",
        "logging",
    ]

    for key in required_top_level_keys:
        if key not in config:
            raise ConfigError(f"Missing top-level config key: '{key}'")

    # ---- Site ----
    _require_keys(config["site"], ["site_id", "site_name"], "site")

    # ---- Reports ----
    _require_keys(
        config["reports"],
        ["default_output_dir"],
        "reports"
    )

    # ---- Meter ----
    _require_keys(
        config["meter"],
        ["ip", "port", "timeout_ms", "retry_count"],
        "meter"
    )

    # ---- Logging ----
    _require_keys(
        config["logging"],
        ["level"],
        "logging"
    )


def _require_keys(section: Dict, keys: list, section_name: str) -> None:
    """Utility validator for required keys in a config section."""
    for key in keys:
        if key not in section:
            raise ConfigError(
                f"Missing key '{key}' in config section '{section_name}'"
            )


# =============================================================================
# Normalization Helpers
# =============================================================================
def _normalize_config(config: Dict) -> None:
    """
    Normalizes configuration values (paths, types).
    Mutates config in-place.
    """

    # Normalize report output directory
    reports = config.get("reports", {})
    output_dir = reports.get("default_output_dir")

    if output_dir:
        normalized_path = os.path.abspath(
            os.path.expandvars(os.path.expanduser(output_dir))
        )
        reports["default_output_dir"] = normalized_path
        logger.debug(
            "Normalized report output directory: %s", normalized_path
        )

    # Ensure meter mock flag exists (default: False)
    meter = config.get("meter", {})
    meter.setdefault("mock", False)

    # Ensure AVR section exists (optional but expected)
    config.setdefault("avr", {})
    config["avr"].setdefault("rated_output_voltage", 230.0)