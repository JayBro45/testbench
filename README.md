# IPS Testing Software

Industrial Test Bench application for **AVR** (Automatic Voltage Regulator) and **SMR** (Switched Mode Rectifier) testing. Connects to a Hioki power meter, captures live readings, and exports Excel reports with acceptance checks.

**Author:** Jayant

---

## Features

- **Dual test modes:** AVR (AC output) and SMR (DC output) with automatic hardware mode switching
- **Live readings:** Real-time input/output values from the meter while a test is running
- **Data grid:** Save multiple readings per test; clear or delete rows as needed
- **Excel export:** Generate asubmission workbook per test (AVR- or SMR-specific, with pass/fail result) to a configurable folder
- **Configuration:** Default export directory, meter IP, and default test mode via **Configuration → Settings** or by editing `config.json`
- **Mock mode:** Run without hardware by setting `meter.mock: true` in `config.json`

---

## Requirements

- **Python 3.10+**
- Windows 10/11 (for running or building the executable)
- For real meter use: Hioki meter on the network; drivers (e.g. VISA) if required

---

## Running from source

1. Clone or extract the project and open a terminal in the project root.
2. Create a virtual environment (recommended) and install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Ensure `config.json` is in the project root (see **Configuration** below).
4. Run the application:
   ```bash
   python main.py
   ```
   On Windows you can use `run.bat` if available.

---

## Configuration

- **Settings dialog:** Use **Configuration → Settings** to set:
  - Default export directory (where Excel reports are saved)
  - Meter IP address
  - Default test mode (AVR or SMR)

  Changes are saved to `config.json` and persist across restarts.

- **config.json:** Must be in the same folder as the script (or the built `.exe`). You can edit it directly for:
  - `app_name`, `site` (site_id, site_name)
  - `reports.default_output_dir`
  - `meter.ip`, `meter.port`, `meter.timeout_ms`, `meter.retry_count`, `meter.mock`
  - `default_test_mode` ("AVR" or "SMR")
  - `logging.level`

  Restart the application after editing the file.

---

## Building a standalone executable

To build a single Windows executable (no Python required on target PCs):

1. Install dependencies and PyInstaller:  
   `pip install -r requirements.txt` (PyInstaller is in requirements).
2. Run **build.bat** (or `pyinstaller testbench.spec`).
3. Output: `dist\IPS_Testing_Software.exe`.

For distribution, copy the `.exe`, `config.json`, and (optionally) `docs\USER_MANUAL.txt` into a folder and zip it. See **docs/BUILD_AND_DISTRIBUTE.md** for full details.

---

## Project structure (overview)

| Path | Description |
|------|-------------|
| `main.py` | Application entry point |
| `ui.py` | Main window, menus, grid, live panels |
| `config_loader.py` | Load/save `config.json` |
| `meter_hioki.py` | Hioki meter communication |
| `worker.py` | Background polling worker |
| `strategies/` | AVR and SMR strategy implementations |
| `avr_*` / `smr_*` | Acceptance engines and Excel report generators |
| `settings_dialog.py` | Configuration dialog |
| `config.json` | Application configuration (required at runtime) |
| `docs/` | User manual, build instructions, acceptance check docs |

---

## Documentation

- **docs/USER_MANUAL.txt** — End-user guide (installation, usage, menus, troubleshooting)
- **docs/BUILD_AND_DISTRIBUTE.md** — Building and distributing the executable
- **docs/AVR_ACCEPTANCE_CHECKS.md** — AVR acceptance criteria
- **docs/SMR_ACCEPTANCE_CHECKS.md** — SMR acceptance criteria

---

## Logs

When the application runs, it creates a `logs` folder next to the executable (or script) and writes **testbench.log** there. Use it for troubleshooting meter or export issues.
