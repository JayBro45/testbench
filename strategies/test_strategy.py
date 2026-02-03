"""
strategies/test_strategy.py
===========================

Abstract Base Class (Interface) for Test Strategies.

This module defines the contract that all test modes (AVR, SMR, etc.) must implement.
The Main Window uses this interface to interact with the underlying logic dynamically,
without knowing the specific details of the test being performed.

Responsibilities
----------------
- Define UI configuration (Grid Headers, Live Reading Labels)
- Enforce implementation of Validation logic
- Enforce implementation of Report Generation logic
"""

from abc import ABC, abstractmethod
from typing import List, Dict, Any, Tuple


class TestStrategy(ABC):
    """
    Abstract Interface for Test Strategies.

    All specific test implementations (AVR, SMR) must inherit from this class
    and implement all abstract methods.
    """

    def __init__(self, config: Dict[str, Any]):
        """
        Initialize the strategy with the application configuration.

        :param config: The full application configuration dictionary.
        """
        self.config = config

    @property
    @abstractmethod
    def name(self) -> str:
        """
        Returns the display name of the test mode (e.g., 'AVR Test').
        Used in the UI selector.
        """
        pass

    @property
    @abstractmethod
    def grid_headers(self) -> List[str]:
        """
        Returns the list of column headers for the main results grid.
        These must match the keys expected by the validation engine.
        """
        pass

    @property
    @abstractmethod
    def live_readings_map(self) -> Dict[str, Tuple[str, str]]:
        """
        Maps internal meter data keys to UI display labels and units.

        Format:
            {
                "internal_key": ("Display Label", "Unit"),
                ...
            }

        Example:
            {
                "vin": ("Input Volts", "V"),
                "iin": ("Input Current", "A")
            }
        """
        pass

    @abstractmethod
    def create_row_data(self, reading: Dict[str, float]) -> List[str]:
        """
        Converts a raw reading dictionary into a list of strings 
        matching grid_headers order.
        """
        pass

    @abstractmethod
    def validate(self, rows: List[Dict[str, Any]]) -> Any:
        """
        Executes the acceptance logic for the collected data.

        :param rows: List of dictionaries representing the grid data.
        :return: A result object (e.g., AVRResult, SMRResult) containing
                 pass/fail status and detailed summaries.
        """
        pass

    @abstractmethod
    def generate_reports(self, rows: List[Dict[str, Any]], output_dir: str, prefix: str) -> None:
        """
        Generates the specific Excel reports for this test mode.

        :param rows: The data rows to report.
        :param output_dir: The directory where files should be saved.
        :param prefix: A timestamp/identifier prefix for the filenames.
        """
        pass