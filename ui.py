import os
import sys
import subprocess
from datetime import datetime

# =============================================================================
# GUI Imports
# =============================================================================
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QLabel, QPushButton, QRadioButton, QButtonGroup,
    QTableWidget, QLineEdit, QFrame, QStatusBar, QHeaderView,
    QStyle, QSizePolicy, QTableWidgetItem, QMessageBox, QInputDialog, QMenu
)
from PySide6.QtCore import Qt, QSize

# =============================================================================
# Local Application Imports
# =============================================================================
from worker import MeterPollingWorker
from avr_excel_report import generate_avr_excel_report
from avr_submission_report import generate_avr_submission_excel
from settings_dialog import SettingsDialog
from config_loader import save_config
from logger import get_logger 

AVR_REQUIRED_ROWS = 6


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
    Main application logic for the AVR Test Bench software.
    Handles UI construction, device polling, data aggregation, and report generation.
    """

    def __init__(self, meter, app_config):
        super().__init__()
        
        # 1. Configuration & Hardware Setup
        self.meter = meter
        self.config = app_config
        self.meter.mock = bool(self.config.get("meter", {}).get("mock", False))
        self.rated_output_voltage = self.config.get("avr", {}).get("rated_output_voltage", 230.0)
        
        # 2. Logger Setup
        self.logger = get_logger("MainWindow") 
        self.logger.info("Application started")
      
        # 3. Threading State
        self.polling_worker = None
        self.is_polling = False
        self.latest_data = {}

        # 4. Window Setup
        self.setWindowTitle(self.config.get("app_name", "Test Bench Software"))
        self.resize(1200, 750)

        # 5. Build UI
        self._build_menus()
        self._build_ui()
        self._build_statusbar()
        
        # 6. Apply Defaults
        default_dir = self.config.get("reports", {}).get("default_output_dir")
        if default_dir:
            self.location_edit.setText(default_dir)

        # 7. Initial State
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

        self.logger.info("Starting test polling...") 

        # Lock UI Controls
        self.start_btn.setEnabled(False)
        self.radio_1p.setEnabled(False)
        self.radio_3p.setEnabled(False)

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

        self.radio_1p.setEnabled(True)
        self.radio_3p.setEnabled(False)

        self.start_btn.setText("START TEST")
        self.start_btn.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        self.start_btn.setEnabled(True)

        self.statusbar.showMessage("Status: Idle")
        self.logger.info("Polling finished - Status: Idle") 

    def update_live_readings(self, data: dict):
        """Slot: Updates the UI panels with fresh data from the worker."""
        self.latest_data = data
        self.save_btn.setEnabled(True)
        
        # Update Input Panel
        self.input_panel.update_value("vin", data.get("vin"))
        self.input_panel.update_value("iin", data.get("iin"))
        self.input_panel.update_value("kwin", data.get("kwin"))
        self.input_panel.update_value("freq", data.get("frequency"))
        
        # Update Output Panel
        self.output_panel.update_value("vout", data.get("vout"))
        self.output_panel.update_value("iout", data.get("iout"))
        self.output_panel.update_value("kwout", data.get("kwout"))
        self.output_panel.update_value("vthd", data.get("vthd_out"))

    def on_worker_warning(self, msg):
        self.statusbar.showMessage(f"Warning: {msg}", 3000)

    def on_worker_error(self, msg):
        self.statusbar.showMessage(f"Error: {msg}")
        self.on_polling_finished()

    def on_status_message(self, msg):
        self.statusbar.showMessage(msg)

    # =========================================================================
    # SECTION 2: Data Management (Grid & Math)
    # =========================================================================

    def _derive_load_line(self, row_index: int, vout: float | None):
        """
        Calculates Load/Line regulation based on row position.
        :param row_index: 1-based index (operator view)
        :return: Tuple (load_str, line_str)
        """
        if vout is None:
            return "--", "--"

        value = round((1 - vout / self.rated_output_voltage) * 100, 2)

        # Hard-coded business logic for AVR reporting
        if row_index in (2, 3, 6):
            return f"{value:.2f}", "--"

        if row_index in (4, 5):
            return "--", f"{value:.2f}"

        return "--", "--"
    
    def save_current_reading(self):
        """Snapshots the current live reading and appends it to the grid."""
        if not hasattr(self, "latest_data") or not self.latest_data:
            self.statusbar.showMessage("No data to save")
            return

        row = self.table.rowCount()
        self.table.insertRow(row)
        row_number = row + 1  # 1-based logic for calc

        d = self.latest_data

        # Extract & format data
        freq = round(d.get("frequency", 0), 2)
        vin = round(d.get("vin", 0), 1)
        iin = round(d.get("iin", 0), 2)
        kwin = round(d.get("kwin", 0), 2)

        vout = round(d.get("vout", 0), 1)
        iout = round(d.get("iout", 0), 2)
        kwout = round(abs(d.get("kwout", 0)), 2)

        vthd = round(d.get("vthd_out", 0), 1)
        eff = round(d.get("efficiency", 0), 2)

        load, line = self._derive_load_line(row_number, vout)

        values = [
            freq, vin, iin, kwin,
            vout, iout, kwout,
            vthd, eff, load, line
        ]

        for col, val in enumerate(values):
            item = QTableWidgetItem(str(val))
            item.setTextAlignment(Qt.AlignCenter)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, col, item)

        self.statusbar.showMessage(f"Row {row_number} saved")
        self.logger.info(f"Captured Row {row_number}") 
    
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
        """Generates the AVR Engineering and Submission reports."""
        row_count = self.table.rowCount()

        # Validation: Hard AVR rule
        if row_count != AVR_REQUIRED_ROWS:
            QMessageBox.warning(
                self,
                "Invalid Row Count",
                f"AVR report requires exactly {AVR_REQUIRED_ROWS} rows in the grid.\n\n"
                f"Current rows: {row_count}"
            )
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

        # Collect grid data
        headers = [
            "Frequency", "V (in)", "I (in)", "kW (in)",
            "V (out)", "I (out)", "kW (out)", "VTHD (out)",
            "Efficiency", "Load", "Line"
        ]

        rows = []
        for r in range(row_count):
            row = {}
            for c, header in enumerate(headers):
                item = self.table.item(r, c)
                row[header] = item.text() if item else ""
            rows.append(row)

        # Generate Paths
        base_name = f"AVR_TEST_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        export_dir = self.get_or_create_export_folder(base_dir)
        
        if not export_dir:
            return

        engineering_path = os.path.join(export_dir, f"{base_name}_RESULT.xlsx")
        submission_path = os.path.join(export_dir, f"{base_name}_SUBMISSION.xlsx")

        try:
            generate_avr_excel_report(rows, engineering_path, self.rated_output_voltage)
            generate_avr_submission_excel(rows, submission_path)

            msg = "Both AVR reports were generated successfully"
            QMessageBox.information(
                self,
                "Reports Generated Successfully",
                f"{msg}:\n\n{export_dir}"
            )
            self.statusbar.showMessage(msg)
            self.logger.info(f"{msg} at {export_dir}") 

        except Exception as e:
            self.logger.error("Export Failed", exc_info=True) 
            QMessageBox.critical(
                self,
                "Export Failed",
                f"Failed to generate AVR reports:\n\n{e}"
            )

    def get_or_create_export_folder(self, base_dir: str) -> str | None:
        """Prompts user for a specific test folder name."""
        default_name = f"AVR_{datetime.now().strftime('%d-%m-%Y_%H%M')}"

        folder_name, ok = QInputDialog.getText(
            self,
            "Export Folder Name",
            "Enter folder name for this test:",
            text=default_name
        )

        if not ok or not folder_name.strip():
            self.statusbar.showMessage("Export cancelled (no folder name)")
            return None

        final_path = os.path.join(base_dir, folder_name.strip())
        os.makedirs(final_path, exist_ok=True)
        return final_path

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

            self.rated_output_voltage = self.config.get(
                "avr", {}
            ).get("rated_output_voltage", 230.0)

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
            "1. Start Test\n"
            f"2. Save {AVR_REQUIRED_ROWS} readings\n"
            "3. Export Excel\n\n"
            "Delete rows using DEL key."
        )

    def keyPressEvent(self, event):
        """Global key handler for the Main Window."""
        if event.key() == Qt.Key_Delete:
            self.delete_selected_rows()
        else:
            super().keyPressEvent(event)

    def _show_context_menu(self, position):
        """Triggered on right-click; shows a 'Delete' option if rows are selected."""
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
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(15)

        # --- Top Section Container ---
        top_layout = QHBoxLayout()
        top_layout.setSpacing(16)

        # 1. Test Type Selection
        test_group = QGroupBox("1. Choose the type of test")
        test_group.setFixedWidth(260)
        test_layout = QVBoxLayout(test_group)
        
        self.radio_1p = QRadioButton("1 Phase")
        self.radio_3p = QRadioButton("3 Phase")
        self.radio_1p.setChecked(True)
        self.radio_3p.setEnabled(False)
        self.radio_3p.setToolTip("AVR tests are single-phase only")
        
        # Radio Styling
        self.radio_1p.setStyleSheet("font-size: 13px; padding: 5px;")
        self.radio_3p.setStyleSheet("font-size: 13px; padding: 5px;")
        
        self.phase_group = QButtonGroup()
        self.phase_group.addButton(self.radio_1p)
        self.phase_group.addButton(self.radio_3p)
        
        test_layout.addWidget(self.radio_1p)
        test_layout.addWidget(self.radio_3p)
        test_layout.addStretch()
        top_layout.addWidget(test_group)

        # 2. Live Readings Section
        live_group = QGroupBox("2. Observe the real-time readings by clicking Start")
        live_layout = QVBoxLayout(live_group)
        panels_layout = QHBoxLayout()
        panels_layout.setSpacing(12)

        self.input_panel = ReadingsPanel("Input Side Readings", [
            ("vin", "V Ph", "V"), ("iin", "I Ph", "A"),
            ("kwin", "1-Ph kW", "kW"), ("freq", "Frequency", "Hz"),
        ])
        
        self.output_panel = ReadingsPanel("Output Side Readings", [
            ("vout", "V Ph", "V"), ("iout", "I Ph", "A"),
            ("kwout", "1-Ph kW", "kW"), ("vthd", "V THD", "%"),
        ])
        
        panels_layout.addWidget(self.input_panel)
        panels_layout.addWidget(self.output_panel)

        # Start Button
        btn_layout = QVBoxLayout()
        btn_layout.setAlignment(Qt.AlignCenter)
        self.start_btn = QPushButton("START TEST")
        self._style_button(self.start_btn, "#2E7D32", QStyle.SP_MediaPlay)
        self.start_btn.clicked.connect(self.toggle_polling)
        
        btn_layout.addWidget(self.start_btn)
        panels_layout.addLayout(btn_layout)
        
        live_layout.addLayout(panels_layout)
        top_layout.addWidget(live_group, stretch=1)
        main_layout.addLayout(top_layout)

        # 3. Grid Section
        grid_group = QGroupBox("3. Save the required readings into the grid")
        grid_layout = QVBoxLayout(grid_group)
        
        self.table = QTableWidget(0, 11)
        self.table.setHorizontalHeaderLabels([
            "Frequency", "V (in)", "I (in)", "kW (in)",
            "V (out)", "I (out)", "kW (out)", "VTHD (out)", 
            "Efficiency", "Load", "Line"
        ])
        self.table.setMinimumHeight(250)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        grid_layout.addWidget(self.table)

        # Grid Action Buttons
        grid_btns_layout = QHBoxLayout()
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
        export_group = QGroupBox("4. Export the readings in the grid to Excel")
        export_layout = QVBoxLayout(export_group)
        fields = QHBoxLayout()
        
        self.location_edit = QLineEdit()
        self.location_edit.setPlaceholderText("D:/AVR_TEST_REPORTS")
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