# Building and Distributing IPS Testing Software

This document describes how to build a standalone executable with PyInstaller so the application can run on other Windows systems without installing Python.

## Prerequisites

- **Python 3.10+** with pip
- All project dependencies installed: `pip install -r requirements.txt`

## Quick build (Windows)

1. Open a command prompt in the project root.
2. Run:
   ```bat
   build.bat
   ```
3. The executable is created at:
   ```
   dist\IPS_Testing_Software.exe
   ```

## Manual build

1. Install dependencies and PyInstaller:
   ```bat
   pip install -r requirements.txt
   pip install pyinstaller
   ```

2. Run PyInstaller using the provided spec file:
   ```bat
   pyinstaller testbench.spec
   ```

3. Output:
   - **Single executable:** `dist\IPS_Testing_Software.exe` (one-file bundle; no console window).

## Creating a distribution zip

To ship the app as a single zip file for end users:

1. Build the app (run `build.bat`). You will have `dist\IPS_Testing_Software.exe`.
2. Create a folder (e.g. `IPS_Testing_Software_Release`) and put in it:
   - `IPS_Testing_Software.exe` (from `dist`)
   - `config.json` (copy from project root; edit if needed for default site/meter)
   - `USER_MANUAL.txt` (copy from `docs\USER_MANUAL.txt`)
3. Zip this folder and send it. Recipients unzip and follow the **User Manual** to install and use the software.

---

## Distributing to other systems

1. **Copy the executable and config**
   - Copy `dist\IPS_Testing_Software.exe` and a `config.json` to the target machine.
   - Place **config.json in the same folder** as the exe. The app reads and writes only this one config file (no bundled fallback).
   - The **company logo** (window icon and header) is bundled from `assets/logo.png`; add that file before building if you want it.

2. **Logs**
   - At runtime the app creates a `logs` folder next to the exe and writes `testbench.log` there.

3. **Reports**
   - Excel reports are saved to the path set in config (`reports.default_output_dir`) or the folder chosen in the UI.

4. **Target system**
   - Windows 10/11.
   - No Python or extra runtimes required.
   - For real meter use: correct network/hardware and drivers (e.g. VISA if used) must be installed on the target PC.

## Build options (editing `testbench.spec`)

- **Console window:** Set `console=True` in the `EXE()` block to show a console for debugging.
- **One-folder build:** To produce a folder with exe + DLLs instead of one file, change the spec to use `COLLECT` and point users to the exe inside the collected folder (see PyInstaller docs).
- **Icon:** The window/taskbar icon comes from `assets/logo.png`. To set the **.exe file icon** in Explorer, add `assets/logo.ico` and set `icon='assets/logo.ico'` in the `EXE()` block.

## Troubleshooting

- **“Config file not found”**  
  Place a valid `config.json` in the same directory as the exe. You can copy it from the project root when distributing.

- **Missing DLL or import errors**  
  Add the missing module to `hiddenimports` in `testbench.spec` and rebuild.

- **Antivirus warning**  
  PyInstaller one-file executables are sometimes flagged. You can use a one-folder build or sign the exe for production.
