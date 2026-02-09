"""
ui.py
=====

Main User Interface for the Industrial Test Bench.

Responsibilities
----------------
- Manage the main application window and layout.
- Handle user interactions (Start/Stop, Save, Export).
- Visualize real-time data from the MeterPollingWorker.
- Orchestrate the Strategy Pattern (Switching between AVR and SMR modes).
- Delegate hardware configuration and report generation to the active Strategy.

Design Notes
------------
- This module is "domain-agnostic" where possible. It asks the active
  Strategy object for headers, labels, and validation rules.
- Thread safety is managed via Qt Signals/Slots with the Worker.
"""

import os
import sys
import subprocess
from datetime import datetime

# =============================================================================
# GUI Imports
# =============================================================================
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QLabel, QPushButton, QComboBox,
    QTableWidget, QLineEdit, QFrame, QStatusBar, QHeaderView,
    QStyle, QSizePolicy, QTableWidgetItem, QMessageBox, QInputDialog, QMenu
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon

# =============================================================================
# Local Application Imports
# =============================================================================
from worker import MeterPollingWorker
from settings_dialog import SettingsDialog
from config_loader import save_config
from logger import get_logger

# Strategy Imports
from strategies.avr_strategy import AVRStrategy
from strategies.smr_strategy import SMRStrategy


# =============================================================================
# Asset Path (works when running from source and when frozen via PyInstaller)
# =============================================================================
def _logo_path():
    """Path to assets/logo.png next to the script or inside the frozen bundle."""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "assets", "logo.png")


# =============================================================================
# Custom Widget: Single Reading Display
# =============================================================================
class ReadingDisplay(QFrame):
    """
    A specific custom widget to display a single sensor reading.
    Consists of a small label (title) and a large value display.
    """
    def __init__(self, label: str, unit: str = ""):
        super().__init__()
        self.unit = unit
        
        # Frame Styling
        self.setFrameStyle(QFrame.Box)
        self.setStyleSheet("""
            QFrame { 
                background-color: #FFF8F0; 
                border: 1px solid #E0C8A0; 
            }
        """)

        # Layout Setup
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)

        # Label Component
        self.label = QLabel(label)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("font-size: 10px; color: #555;")

        # Value Component
        self.value = QLabel("--")
        self.value.setAlignment(Qt.AlignCenter)
        self.value.setStyleSheet("font-size: 18px; font-weight: bold; color: #C44;")

        layout.addWidget(self.label)
        layout.addWidget(self.value)

    def set_value(self, value: float | None, decimals: int = 2):
        """Updates the numeric display. Shows '--' if value is None."""
        if value is None:
            self.value.setText("--")
        else:
            self.value.setText(f"{value:.{decimals}f} {self.unit}")


# =============================================================================
# Custom Widget: Group of Readings
# =============================================================================
class ReadingsPanel(QGroupBox):
    """
    A container GroupBox that holds multiple ReadingDisplay widgets in a grid layout.
    """
    def __init__(self, title: str, items: list):
        """
        :param title: Title of the GroupBox
        :param items: List of tuples in format (key, label, unit)
        """
        super().__init__(title)
        
        # GroupBox Styling
        self.setStyleSheet("""
            QGroupBox { 
                font-weight: bold; 
                border: 2px solid #4A6A8A; 
                margin-top: 10px; 
            }
            QGroupBox::title {
                subcontrol-origin: margin; 
                subcontrol-position: top left;
                background-color: #4A6A8A; 
                color: white; 
                padding: 2px 8px;
            }
        """)

        layout = QGridLayout(self)
        layout.setSpacing(8)
        
        self.displays = {}

        # Dynamically create widgets based on 'items' list
        for i, (key, label, unit) in enumerate(items):
            widget = ReadingDisplay(label, unit)
            row, col = divmod(i, 2)
            layout.addWidget(widget, row, col)
            self.displays[key] = widget

    def update_value(self, key: str, value: float | None):
        """Updates a specific display widget by its key."""
        if key in self.displays:
            self.displays[key].set_value(value)


# =============================================================================
# Main Application Window
# =============================================================================
class MainWindow(QMainWindow):
    """
    Main application logic for the Test Bench software.
    Handles UI construction, device polling, strategy management, and reporting.
    """

    def __init__(self, meter, app_config):
        super().__init__()
        
        # 1. Configuration & Hardware Setup
        self.meter = meter
        self.config = app_config
        self.meter.mock = bool(self.config.get("meter", {}).get("mock", False))
        
        # 2. Logger Setup
        self.logger = get_logger("MainWindow") 
        self.logger.info("Application started")
      
        # 3. Threading State
        self.polling_worker = None
        self.is_polling = False
        self.latest_data = {}

        # 4. Strategy Initialization
        # Load available strategies
        self.strategies = {
            "AVR": AVRStrategy(self.config),
            "SMR": SMRStrategy(self.config)
        }

        # 5. Window Setup
        self.setWindowTitle(self.config.get("app_name", "Test Bench Software"))
        self.setMinimumSize(1000, 650)
        self.resize(1200, 750)
        logo_path = _logo_path()
        if os.path.isfile(logo_path):
            self.setWindowIcon(QIcon(logo_path))

        # 6. Build UI
        self._build_menus()
        self._build_ui()
        self._build_statusbar()
        
        # 7. Apply Defaults & Initial Strategy
        default_dir = self.config.get("reports", {}).get("default_output_dir")
        if default_dir:
            self.location_edit.setText(default_dir)

        # Apply the initial strategy from config (Sets headers, panels, meter mode)
        default_mode = self.config.get("default_test_mode", "AVR")
        if default_mode not in self.strategies:
            default_mode = "AVR"
        self.mode_selector.setCurrentText(default_mode)
        self.current_strategy = self.strategies[default_mode]
        self.apply_strategy(default_mode)
        
        # 8. Initial State
        self.save_btn.setEnabled(False)

    # =========================================================================
    # SECTION 1: Polling Logic & Thread Management
    # =========================================================================

    def toggle_polling(self):
        """Toggles the start/stop state of the meter polling."""
        if self.is_polling:
            self.stop_polling()
        else:
            self.start_polling()

    def start_polling(self):
        """Initiates the background worker thread to poll the meter."""
        # Validation
        if not self.config.get("meter", {}).get("ip") and not self.meter.mock:
            msg = "Meter IP not configured"
            self.statusbar.showMessage(msg)
            self.logger.warning(msg)
            return

        if self.is_polling:
            return

        self.logger.info(f"Starting test polling ({self.current_strategy.name})...") 

        # Lock UI Controls
        self.start_btn.setEnabled(False)
        self.mode_selector.setEnabled(False)

        # Create and Configure Worker
        self.polling_worker = MeterPollingWorker(self.meter, interval_sec=1.0)
        
        # Connect Signals
        self.polling_worker.data_ready.connect(self.update_live_readings)
        self.polling_worker.warning.connect(self.on_worker_warning)
        self.polling_worker.error.connect(self.on_worker_error)
        self.polling_worker.status.connect(self.on_status_message)
        
        # Lifecycle Management
        self.polling_worker.finished_polling.connect(self.on_polling_finished)
        self.polling_worker.finished.connect(self.polling_worker.deleteLater)

        self.polling_worker.start()

        # Update UI State
        self.start_btn.setText("Stop")
        self.start_btn.setIcon(self.style().standardIcon(QStyle.SP_MediaStop))
        self.start_btn.setEnabled(True)
        self.is_polling = True

    def stop_polling(self):
        """Signals the worker thread to stop."""
        if not self.polling_worker:
            return

        self.logger.info("Stopping test polling...") 
        self.statusbar.showMessage("Stopping...")
        self.start_btn.setEnabled(False)
        self.polling_worker.stop()

    def on_polling_finished(self):
        """Cleanup after polling stops."""
        self.polling_worker = None
        self.is_polling = False

        self.mode_selector.setEnabled(True)

        self.start_btn.setText("START TEST")
        self.start_btn.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        self.start_btn.setEnabled(True)

        self.statusbar.showMessage("Status: Idle")
        self.logger.info("Polling finished - Status: Idle") 

    def update_live_readings(self, data: dict):
        """
        Slot: Updates the UI panels with fresh data from the worker.
        Dynamically maps data to widgets based on the active strategy.
        """
        self.latest_data = data
        self.save_btn.setEnabled(True)
        
        # Get the mapping from the strategy (Key -> (Label, Unit))
        # But here we just need the Keys to match the widget IDs in ReadingsPanel
        mapping = self.current_strategy.live_readings_map
        
        # Iterate and update Input/Output panels if the key exists
        for key in mapping.keys():
            val = data.get(key)
            self.input_panel.update_value(key, val)
            self.output_panel.update_value(key, val)

    def on_worker_warning(self, msg):
        """Handle non-fatal warnings emitted by the background worker."""
        self.statusbar.showMessage(f"Warning: {msg}", 3000)

    def on_worker_error(self, msg):
        """Handle errors emitted by the background worker and stop polling."""
        self.logger.error(f"Worker error: {msg}")
        self.statusbar.showMessage(f"Error: {msg}")
        self.on_polling_finished()

    def on_status_message(self, msg):
        """Update the status bar with a generic informational message."""
        self.statusbar.showMessage(msg)

    # =========================================================================
    # SECTION 2: Strategy Management & Data Capture
    # =========================================================================

    def change_test_mode(self, mode_name: str):
        """
        Slot: Called when the user changes the Test Mode ComboBox.
        """
        if mode_name not in self.strategies:
            return
        
        # Prevent re-triggering if already on this mode
        if self.current_strategy == self.strategies[mode_name]:
            return

        self.logger.info(f"Switching strategy to: {mode_name}")
        
        # Define the actual switch logic as a closure/method
        def perform_switch():
            self.current_strategy = self.strategies[mode_name]
            self.apply_strategy(mode_name)

        # If running, stop first and wait for finish signal
        if self.is_polling:
            # We must disconnect existing connections to avoid double-firing
            try:
                self.polling_worker.finished_polling.disconnect(self.on_polling_finished)
            except RuntimeError:
                pass # logic to handle if not connected
            
            # Create a temporary slot to handle the sequence
            def on_stop_complete():
                self.on_polling_finished() # Perform standard cleanup
                perform_switch()           # Switch strategy
                # Restore standard connection for next time
                # (Note: on_polling_finished is usually connected in start_polling, 
                # but we intercepted the specific stop event here)

            self.polling_worker.finished_polling.connect(on_stop_complete)
            self.stop_polling()
        else:
            perform_switch()

    def apply_strategy(self, mode_name: str):
        """
        Applies settings for the selected strategy:
        1. Updates Grid Headers.
        2. Rebuilds Live Reading Panels.
        3. Configures Meter Hardware (AC/AC vs AC/DC).
        """
        # 1. Update Grid
        headers = self.current_strategy.grid_headers
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(0) # Clear previous test data
        
        # 2. Update Panels
        self._rebuild_readings_panels()

        # 3. Configure Hardware
        if mode_name == "AVR":
            self.meter.set_mode_avr()
        elif mode_name == "SMR":
            self.meter.set_mode_smr()
            
        self.logger.info(f"Strategy {mode_name} applied successfully")

    def _rebuild_readings_panels(self):
        """
        Destroys and recreates the Input/Output panels based on 
        the current strategy's 'live_readings_map'.
        """
        # Clear existing layout in the live_group wrapper
        # We need to access the layout where panels live.
        # This is 'self.panels_layout' defined in _build_ui. 
        # But strictly, we should remove the old widgets first.
        
        # Helper to clear a layout
        def clear_layout(layout):
            while layout.count():
                child = layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()

        clear_layout(self.panels_layout)

        # Get Mapping: {'vin': ('V (in)', 'V'), ...}
        full_map = self.current_strategy.live_readings_map
        
        # Split into Input and Output lists for display
        # Logic: Input keys usually end in 'in' or 'freq'/'pf'. Output in 'out'/'ripple'.
        input_items = []
        output_items = []

        # Explicitly define what goes to Output. Everything else is Input.
        output_keys = ["vout", "iout", "kwout", "pout", "ripple", "efficiency", "vthd_out", "vout_dc", "iout_dc"]
        
        for key, (label, unit) in full_map.items():
            if key in output_keys:
                output_items.append((key, label, unit))
            else:
                input_items.append((key, label, unit))

        # Re-instantiate Panels
        self.input_panel = ReadingsPanel("Input Readings", input_items)
        self.output_panel = ReadingsPanel("Output Readings", output_items)
        
        self.panels_layout.addWidget(self.input_panel)
        self.panels_layout.addWidget(self.output_panel)
        
        btn_layout = QVBoxLayout()
        btn_layout.setAlignment(Qt.AlignCenter)
        btn_layout.addWidget(self.start_btn)
        self.panels_layout.addLayout(btn_layout)

    def save_current_reading(self):
        """
        Snapshots current data and adds it to the grid.
        Adapts data formatting based on the active strategy.
        """
        if not hasattr(self, "latest_data") or not self.latest_data:
            self.statusbar.showMessage("No data to save")
            return
        
        # Calculate the NEXT row index (1-based for the logic)
        next_row_index = self.table.rowCount() + 1

        try:
            values = self.current_strategy.create_row_data(self.latest_data, row_index=next_row_index)
        except Exception as e:
            self.logger.error(f"Error creating row data: {e}")
            self.statusbar.showMessage("Error formatting data for grid")
            return
        
        # Populate Table
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col, val in enumerate(values):
            item = QTableWidgetItem(str(val))
            item.setTextAlignment(Qt.AlignCenter)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, col, item)
            self.table.scrollToBottom()

        self.statusbar.showMessage(f"Row {row + 1} saved")
        # Access name property safely
        mode = getattr(self.current_strategy, "name", "Unknown Mode")
        self.logger.info(f"Captured Row {row + 1} ({mode})")

    def delete_selected_rows(self):
        """Deletes currently selected rows from the grid."""
        selected_rows = sorted(
            set(index.row() for index in self.table.selectedIndexes()),
            reverse=True
        )

        if not selected_rows:
            self.statusbar.showMessage("No rows selected to delete")
            return

        for row in selected_rows:
            self.table.removeRow(row)

        msg = f"{len(selected_rows)} row(s) deleted"
        self.statusbar.showMessage(msg)
        self.logger.info(msg) 

    def clear_entire_grid(self):
        """Clears all data from the grid after confirmation."""
        if self.table.rowCount() == 0:
            self.statusbar.showMessage("Grid already empty")
            return

        reply = QMessageBox.warning(
            self,
            "Clear Entire Grid",
            "This will delete ALL saved readings.\n\nAre you sure?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            self.statusbar.showMessage("Clear grid cancelled")
            return

        self.table.setRowCount(0)
        self.statusbar.showMessage("All rows cleared")
        self.logger.info("Grid cleared by user") 

    # =========================================================================
    # SECTION 3: File Operations & Export
    # =========================================================================

    def export_excel(self):
        """
        Delegates report generation to the active strategy.
        Passes raw grid data to the strategy's generate_reports method.
        """
        row_count = self.table.rowCount()
        
        # Basic check
        if row_count == 0:
            QMessageBox.warning(self, "Empty Grid", "No data to export.")
            return

        base_dir = self.location_edit.text()
        if not base_dir or not os.path.isdir(base_dir):
            QMessageBox.warning(
                self,
                "Invalid Export Location",
                "Default export directory is not configured or does not exist.\n\n"
                "Please set it in Settings."
            )
            return

        # 1. Collect grid data based on current headers
        headers = self.current_strategy.grid_headers
        rows = []
        for r in range(row_count):
            row_data = {}
            for c, h in enumerate(headers):
                item = self.table.item(r, c)
                row_data[h] = item.text() if item else ""
            rows.append(row_data)

        # 2. Prepare Filename Prefix
        # Example: AVR_TEST_20231025_143000
        prefix = f"{self.current_strategy.name.split()[0]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        # 3. Ask for specific folder name (Optional, keeps organized)
        # UI is responsible only for selecting the base directory and a folder name.
        # Each strategy is responsible for creating any required directories.
        folder_name, ok = QInputDialog.getText(
            self,
            "Export Folder Name",
            "Enter folder name for this test:",
            text=prefix,
        )
        if not ok or not folder_name.strip():
            self.statusbar.showMessage("Export cancelled (no folder name)")
            return
        export_dir = os.path.join(base_dir, folder_name.strip())

        # 4. Delegate to Strategy
        try:
            actual_dir = self.current_strategy.generate_reports(rows, export_dir, prefix)

            msg = f"{self.current_strategy.name} reports generated successfully"
            QMessageBox.information(
                self,
                "Success",
                f"{msg}:\n\n{actual_dir}"
            )
            self.statusbar.showMessage("Reports generated")
            self.logger.info(f"Reports generated at {actual_dir}") 

        except Exception:
            # The exception is logged with full traceback for debugging,
            # while the UI shows a concise error message to the user.
            self.logger.error("Export failed while generating reports", exc_info=True)
            QMessageBox.critical(
                self,
                "Export Failed",
                f"Failed to generate reports:\n\n{e}"
            )

    def open_reports_folder(self):
        """Opens the OS file explorer at the default report directory."""
        folder = self.config.get("reports", {}).get("default_output_dir")
        if not folder:
            QMessageBox.warning(self, "No Folder", "Default report folder not configured.")
            return
        
        if sys.platform.startswith("win"):
            os.startfile(folder)
        elif sys.platform.startswith("darwin"):
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])

    # =========================================================================
    # SECTION 4: Menu & Window Events
    # =========================================================================

    def new_test(self):
        """Resets the UI for a fresh test session."""
        if self.polling_worker and self.is_polling:
            self.stop_polling()

        self.table.setRowCount(0)
        self.statusbar.showMessage("New test started")
        self.logger.info("New test initialized") 

    def open_settings(self):
        """Opens the configuration dialog."""
        dialog = SettingsDialog(self.config, self)
        if dialog.exec():
            default_dir = self.config.get("reports", {}).get("default_output_dir")
            if default_dir:
                self.location_edit.setText(default_dir)
            # Apply new default test mode to current session if not polling
            new_mode = self.config.get("default_test_mode", "AVR")
            if new_mode in self.strategies and not self.is_polling:
                self.mode_selector.blockSignals(True)
                self.mode_selector.setCurrentText(new_mode)
                self.mode_selector.blockSignals(False)
                self.current_strategy = self.strategies[new_mode]
                self.apply_strategy(new_mode)

            try:
                save_config(self.config)
                self.statusbar.showMessage("Settings updated and saved")
                self.logger.info("Settings updated") 
            except Exception as e:
                QMessageBox.warning(
                    self, 
                    "Save Error", 
                    f"Settings updated in memory but failed to save to disk:\n{e}"
                )
                self.statusbar.showMessage("Error saving settings")
                self.logger.error(f"Error saving settings: {e}") 

    def show_help(self):
        """Displays usage instructions."""
        QMessageBox.information(
            self,
            "How to Use",
            "1. Select Test Mode (AVR / SMR)\n"
            "2. Start Test to monitor readings\n"
            "3. Save Readings to the grid\n"
            "4. Export Excel to generate reports\n\n"
            "Note: Hardware mode switches automatically."
        )

    def keyPressEvent(self, event):
        """Global key handler for the Main Window."""
        if event.key() == Qt.Key_Delete:
            self.delete_selected_rows()
        else:
            super().keyPressEvent(event)

    def _show_context_menu(self, position):
        """Show a context menu on right-click to allow deleting selected rows."""
        item = self.table.itemAt(position)
        if not item:
            return

        menu = QMenu(self)
        delete_action = menu.addAction("Delete")
        action = menu.exec(self.table.viewport().mapToGlobal(position))
        if action == delete_action:
            self.delete_selected_rows()

    def closeEvent(self, event):
        """Ensures threads are cleaned up before closing."""
        self.logger.info("Application closing...") 
        if self.polling_worker and self.polling_worker.isRunning():
            self.polling_worker.stop()
            # Calculate wait time based on meter config + 1 second buffer.
            meter_timeout = self.config.get("meter", {}).get("timeout_ms", 3000)
            wait_time = meter_timeout + 1000
            if not self.polling_worker.wait(wait_time):
                msg = "Forcing thread termination..."
                self.statusbar.showMessage(msg)
                self.logger.warning(msg) 
                self.polling_worker.terminate()
                
        event.accept()

    # =========================================================================
    # SECTION 5: UI Construction & Styling
    # =========================================================================

    def _build_menus(self):
        """Constructs the Menu Bar."""
        menubar = self.menuBar()

        # File Menu
        file_menu = menubar.addMenu("File")
        new_test_action = file_menu.addAction("New Test")
        new_test_action.triggered.connect(self.new_test)
        open_reports_action = file_menu.addAction("Open Reports Folder")
        open_reports_action.triggered.connect(self.open_reports_folder)
        file_menu.addSeparator()
        exit_action = file_menu.addAction("Exit")
        exit_action.triggered.connect(self.close)

        # Configuration Menu
        config_menu = menubar.addMenu("Configuration")
        settings_action = config_menu.addAction("Settings")
        settings_action.triggered.connect(self.open_settings)

        # Help Menu
        help_menu = menubar.addMenu("Help")
        help_action = help_menu.addAction("Usage Instructions")
        help_action.triggered.connect(self.show_help)

    def _build_ui(self):
        """Constructs the central widget content."""
        central = QWidget()
        central.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(15)

        # --- Top Section Container ---
        top_layout = QHBoxLayout()
        top_layout.setSpacing(16)

        # 1. Test Mode Selection (Refactored from Radio Buttons to ComboBox)
        test_group = QGroupBox("1. Select Test Mode")
        test_group.setMinimumWidth(240)
        test_group.setMaximumWidth(280)
        test_layout = QVBoxLayout(test_group)
        
        self.mode_selector = QComboBox()
        self.mode_selector.addItems(["AVR", "SMR"])
        self.mode_selector.setMinimumHeight(40)
        self.mode_selector.setStyleSheet("font-size: 14px; font-weight: bold; padding: 5px;")
        
        # Connect change event
        self.mode_selector.currentTextChanged.connect(self.change_test_mode)
        
        test_layout.addWidget(self.mode_selector)
        top_layout.addWidget(test_group)

        # 2. Live Readings Section (fixed height so AVR/SMR switch doesn't resize window)
        live_group = QGroupBox("2. Observe Real-Time Readings")
        live_group.setMinimumHeight(220)
        live_layout = QVBoxLayout(live_group)
        self.panels_layout = QHBoxLayout()
        self.panels_layout.setSpacing(12)
        
        # Panels will be injected here dynamically by apply_strategy()
        
        # Start Button (Created once, re-added dynamically)
        self.start_btn = QPushButton("START TEST")
        self._style_button(self.start_btn, "#2E7D32", QStyle.SP_MediaPlay)
        self.start_btn.clicked.connect(self.toggle_polling)
        
        live_layout.addLayout(self.panels_layout)
        top_layout.addWidget(live_group, stretch=1)
        main_layout.addLayout(top_layout)

        # 3. Grid Section
        grid_group = QGroupBox("3. Save Readings")
        grid_layout = QVBoxLayout(grid_group)
        
        # Grid columns are set dynamically by strategy
        self.table = QTableWidget()
        self.table.setMinimumHeight(250)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        grid_layout.addWidget(self.table)

        # Grid Action Buttons
        grid_btn_sep = QFrame()
        grid_btn_sep.setFrameShape(QFrame.HLine)
        grid_btn_sep.setFrameShadow(QFrame.Sunken)
        grid_btn_sep.setLineWidth(1)
        grid_btn_sep.setStyleSheet("QFrame { margin: 6px 0; }")
        grid_layout.addWidget(grid_btn_sep)

        # Grid Action Buttons (dedicated row with spacing)
        grid_btns_layout = QHBoxLayout()
        grid_btns_layout.setSpacing(12)
        grid_btns_layout.setContentsMargins(0, 8, 0, 4)
        self.clear_btn = QPushButton("CLEAR GRID")
        self._style_button(self.clear_btn, "#C62828", QStyle.SP_BrowserReload)
        self.clear_btn.clicked.connect(self.clear_entire_grid)
        
        self.save_btn = QPushButton("SAVE READING")
        self._style_button(self.save_btn, "#1565C0", QStyle.SP_DialogSaveButton)
        self.save_btn.clicked.connect(self.save_current_reading)
        
        grid_btns_layout.addWidget(self.clear_btn)
        grid_btns_layout.addStretch()
        grid_btns_layout.addWidget(self.save_btn)
        grid_layout.addLayout(grid_btns_layout)
        main_layout.addWidget(grid_group)

        # 4. Export Section
        export_group = QGroupBox("4. Export Report")
        export_layout = QVBoxLayout(export_group)
        fields = QHBoxLayout()
        
        self.location_edit = QLineEdit()
        self.location_edit.setPlaceholderText("D:/TEST_REPORTS")
        self.location_edit.setMinimumHeight(35)

        # Lock editing
        self.location_edit.setReadOnly(True)
        self.location_edit.setStyleSheet("""
            QLineEdit {
                background-color: #F0F0F0;
                color: #333;
            }
        """)
        
        self.export_btn = QPushButton("EXPORT EXCEL")
        self._style_button(self.export_btn, "#00695C", QStyle.SP_DirIcon)
        self.export_btn.clicked.connect(self.export_excel)
        
        fields.addWidget(QLabel("Location:"))
        fields.addWidget(self.location_edit)
        fields.addWidget(self.export_btn)
        
        export_layout.addLayout(fields)
        main_layout.addWidget(export_group)

        self.setCentralWidget(central)

    def _style_button(self, button: QPushButton, color: str, icon_std):
        """Applies a consistent modern style to buttons."""
        icon = self.style().standardIcon(icon_std)
        button.setIcon(icon)
        button.setIconSize(QSize(24, 24))
        button.setMinimumHeight(45)
        button.setMinimumWidth(130)
        button.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        button.setCursor(Qt.PointingHandCursor)
        
        button.setStyleSheet(f"""
            QPushButton {{
                background-color: {color}; 
                color: white;
                font-size: 14px; 
                font-weight: bold; 
                border: none;
                border-radius: 6px; 
                padding: 5px 15px; 
                text-align: center;
            }}
            QPushButton:hover {{
                background-color: {color}DD; 
                border: 2px solid #FFFFFF;
            }}
            QPushButton:pressed {{ 
                background-color: #333333; 
            }}
        """)

    def _build_statusbar(self):
        """Constructs the Status Bar."""
        self.statusbar = QStatusBar()
        self.setStatusBar(self.statusbar)
        self.statusbar.setStyleSheet("""
            QStatusBar { 
                background-color: #ECECEC; 
                color: #333; 
                font-weight: bold; 
            }
        """)
        
        mode = "MOCK" if self.meter.mock else "REAL"
        site_id = self.config.get("site", {}).get("site_id", "UNKNOWN")
        self.statusbar.showMessage(f"Site: {site_id} | Mode: {mode} | Status: Idle")