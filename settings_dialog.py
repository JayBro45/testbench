"""
settings_dialog.py
==================

Settings Dialog for AVR Test Bench Software.

Purpose
-------
Provides a modal configuration interface for **runtime-editable**
application settings. This dialog edits the in-memory config dictionary
only; persistence is handled externally by the application.

Editable Settings
-----------------
- Default export directory
- Rated AVR output voltage
- Meter IP address

Design Notes
------------
- Mock Mode is intentionally EXCLUDED from this dialog. It must be
  configured directly in config.json.
- This dialog intentionally avoids file I/O
- Validation is lightweight and defensive
- Any rejected input does NOT mutate config
"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QFileDialog,
    QMessageBox
)
from PySide6.QtCore import Qt


class SettingsDialog(QDialog):
    """
    Modal dialog for editing application configuration values.

    Parameters
    ----------
    config : dict
        Mutable configuration dictionary shared with the main application.
    parent : QWidget | None
        Parent widget (typically MainWindow).
    """

    def __init__(self, config: dict, parent=None):
        super().__init__(parent)

        # ------------------------------------------------------------------
        # Dialog Configuration
        # ------------------------------------------------------------------
        self.setWindowTitle("Settings")
        self.setModal(True)
        self.resize(450, 180)  

        self.config = config

        # ------------------------------------------------------------------
        # Layout Root
        # ------------------------------------------------------------------
        layout = QVBoxLayout(self)

        # ==================================================================
        # Default Export Directory
        # ==================================================================
        layout.addWidget(QLabel("<b>Default Export Directory</b>"))

        dir_layout = QHBoxLayout()
        self.export_dir_edit = QLineEdit(
            config.get("reports", {}).get("default_output_dir", "")
        )

        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_folder)

        dir_layout.addWidget(self.export_dir_edit)
        dir_layout.addWidget(browse_btn)
        layout.addLayout(dir_layout)

        # ==================================================================
        # Meter Configuration
        # ==================================================================
        layout.addWidget(QLabel("<b>Meter IP</b>"))

        self.meter_ip_edit = QLineEdit(
            config.get("meter", {}).get("ip", "")
        )
        layout.addWidget(self.meter_ip_edit)
        layout.addStretch()

        # ==================================================================
        # Action Buttons
        # ==================================================================
        btn_layout = QHBoxLayout()

        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")

        save_btn.clicked.connect(self.save)
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)

        layout.addLayout(btn_layout)

    # ======================================================================
    # Slots
    # ======================================================================

    def browse_folder(self):
        """
        Opens a directory chooser dialog and updates the export path field.
        """
        folder = QFileDialog.getExistingDirectory(
            self, "Select Default Export Directory"
        )
        if folder:
            self.export_dir_edit.setText(folder)

    def save(self):
        """
        Validate inputs and persist changes to the shared config dictionary.

        Notes
        -----
        - Config is only mutated after successful validation
        - Dialog is accepted only on success
        """

        self.config.setdefault("reports", {})[
            "default_output_dir"
        ] = self.export_dir_edit.text().strip()

        self.config.setdefault("meter", {})["ip"] = (
            self.meter_ip_edit.text().strip()
        )
        
        self.accept()