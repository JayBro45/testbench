"""
main.py
=======

Application entry point for the AVR Test Bench Software.

Responsibilities
----------------
- Load application configuration
- Initialize hardware abstraction (meter)
- Bootstrap the Qt application
- Inject dependencies into the main UI window

Design Notes
------------
- Keeps startup logic minimal and explicit
- No business or UI logic is implemented here
- Acts strictly as a composition root
"""

import sys
import os
from PySide6.QtWidgets import QApplication

from ui import MainWindow
from config_loader import load_config
from meter_hioki import HiokiPW3336


def main():
    """
    Application bootstrap sequence.

    Steps:
    1. Load configuration from config.json
    2. Initialize meter driver (mock or real)
    3. Create Qt application instance
    4. Construct and display the main window
    """


    # Load configuration
    try:
        config = load_config()
    except Exception as e:
        # Fallback if logger setup failed or config is missing/corrupt
        print(f"CRITICAL STARTUP ERROR: {e}")
        sys.exit(1)

    # Initialize meter (real or mock based on config)
    meter = HiokiPW3336(config)

    # Start Qt application
    app = QApplication(sys.argv)

    # Create and show main window
    window = MainWindow(
        meter=meter,
        app_config=config
    )
    window.show()

    # Enter Qt event loop
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
