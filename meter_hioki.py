"""
meter_hioki.py
==============

Hardware abstraction layer for the Hioki PW3336 Power Analyzer.

This module provides a stable interface for:
- Establishing and managing communication with the Hioki PW3336
- Reading electrical parameters required for AVR (AC) and SMR (DC) testing
- Supporting both REAL (PyVISA-based) and MOCK (simulation) modes
- Providing resilient read operations with retry and logging

Design Principles:
- Configuration-driven behavior (via config.json)
- No business logic or acceptance rules
- Deterministic public API used by worker/UI layers
- Power values exposed in kW (internally fetched in watts)
- Input Power values taken as Absolute

This module must remain backward-compatible with existing callers.
"""

import pyvisa
import time
import random
from logger import get_logger


class HiokiPW3336:
    """
    Driver class for the Hioki PW3336 Power Analyzer.

    Responsibilities:
    - Manage meter connection lifecycle
    - Configure measurement modes (AC/AC for AVR, AC/DC for SMR)
    - Read electrical parameters from the device
    - Provide mock readings when operating without hardware
    - Expose a consistent data contract to the worker layer

    Supported Operating Modes:
    - REAL: Communicates with physical hardware via PyVISA
    - MOCK: Generates simulated readings for development/testing

    Configuration Dependencies (config.json):
    - meter: connection and retry parameters
    - parameters: controls which values are read in bulk operations
    - logging: logging verbosity

    This class performs **no validation or acceptance logic**.
    """

    # ------------------------------------------------------------------
    # Initialization & Configuration
    # ------------------------------------------------------------------

    def __init__(self, config: dict):
        """
        Initialize the Hioki PW3336 driver.

        Parameters
        ----------
        config : dict
            Parsed application configuration dictionary.
            Must contain a 'meter' section.

        Raises
        ------
        ValueError
            If required configuration is missing or invalid.
        """
        if "meter" not in config:
            raise ValueError("Missing 'meter' section in config.json")

        meter_cfg = config["meter"]

        self.ip = meter_cfg.get("ip")
        self.port = meter_cfg.get("port")
        self.timeout_ms = meter_cfg.get("timeout_ms", 5000)
        self.retry_count = meter_cfg.get("retry_count", 2)
        self.mock = meter_cfg.get("mock", False)

        if not self.mock and not self.ip:
            raise ValueError("Meter IP must be provided when mock=false")

        self.resource_string = (
            f"TCPIP::{self.ip}::{self.port}::SOCKET"
            if not self.mock else None
        )

        self.rm = None
        self.inst = None

        self.logger = get_logger("HiokiPW3336")

        self.logger.info(
            f"Meter initialized | mode={'MOCK' if self.mock else 'REAL'} | ip={self.ip}"
        )

        # Parameter read mask (non-enforcing)
        self._enabled_params = config.get("parameters", {})

        # ------------------------------------------------------------------
        # Mock State (internal values in watts where applicable)
        # ------------------------------------------------------------------
        self._mock_state = {
            # Common
            "vin": 230.0,
            "iin": 5.0,
            "kwin": 1.15*1000,   # kW
            "vout": 230.0,
            "iout": 4.8,
            "kwout": 1.10*1000,  # kW
            "efficiency": 95.0,
            
            # AVR Specific
            "frequency": 50.0,
            "vthd_out": 3.5,
            
            # SMR Specific
            "pf": 0.99,
            "ripple": 0.05,
            "vthd_in": 2.5,
            "ithd_in": 3.0,
            "vout_dc": 48.0,
            "iout_dc": 10.0
        }

    # ------------------------------------------------------------------
    # Connection Management
    # ------------------------------------------------------------------

    def connect(self):
        """
        Establish connection to the meter.

        Returns
        -------
        str
            Identification string from the meter (*IDN?).

        Notes
        -----
        In MOCK mode, no hardware connection is made and a
        simulated ID string is returned.
        """
        if self.mock:
            self.logger.info("Mock mode enabled — skipping hardware connection")
            return "HIOKI,PW3336,MOCK,0.0,SIMULATED"

        try:
            self.logger.info(f"Connecting to {self.resource_string}")

            self.rm = pyvisa.ResourceManager()
            self.inst = self.rm.open_resource(self.resource_string)

            self.inst.read_termination = "\r\n"
            self.inst.write_termination = "\r\n"
            self.inst.timeout = self.timeout_ms

            self.inst.write(":HEADer OFF")

            idn = self.inst.query("*IDN?").strip()
            self.logger.info(f"Connected successfully | IDN={idn}")
            return idn

        except Exception:
            self.logger.error("Failed to connect to meter", exc_info=True)
            raise

    def disconnect(self):
        """
        Disconnect from the meter and release VISA resources.

        Safe to call multiple times.
        """
        self.logger.info("Disconnecting meter")

        if self.inst:
            self.inst.close()
            self.inst = None

        if self.rm:
            self.rm.close()
            self.rm = None

    def is_connected(self) -> bool:
        """
        Check connection status.

        Returns
        -------
        bool
            True if connected to hardware or operating in mock mode.
        """
        return self.mock or self.inst is not None

    # ------------------------------------------------------------------
    # Mode Configuration (REQUIRED FOR UI.PY)
    # ------------------------------------------------------------------

    def set_mode_avr(self):
        """
        Configure meter for AVR Testing.
        PW3336 determines mode by the query (U1/U2 for RMS), so no global setup is required.
        """
        if self.mock: 
            self.logger.info("MOCK: Meter configured for AVR")
            return
        self.logger.info("Meter configured for AVR (Logic handled by specific queries)")

    def set_mode_smr(self):
        if self.mock:
            self.logger.info("MOCK: Meter configured for SMR")
            return

        if not self.inst:
            raise RuntimeError("Meter not connected")

        # Set DC mode
        self.inst.write(":WIRing TYPE1")

        # Clear all measurement items
        self.inst.write(":MEASure:ITEM:ALLClear")

        # Enable required DC outputs
        self.inst.write(":MEASure:ITEM:UDC:CH2 ON")   # DC voltage
        self.inst.write(":MEASure:ITEM:IDC:CH2 ON")   # DC current
        self.inst.write(":MEASure:ITEM:PDC:CH2 ON")   # DC power
        self.inst.write(":MEASure:ITEM:UAC:CH2 ON")   # ripple component

        # Wait for stabilization (2 cycles)
        time.sleep(0.5)

        self.logger.info("Meter configured for SMR")

    # ------------------------------------------------------------------
    # Internal Helpers
    # ------------------------------------------------------------------

    def _mock_read(self, key: str) -> float:
        """
        Generate a simulated reading for the given parameter.
        """
        base = self._mock_state.get(key, 0.0)
        noise = random.uniform(-0.05, 0.05)
        
        # Ensure non-negative for magnitudes
        return round(max(0.0, base + noise), 3)

    def _query_float(self, command: str, key: str) -> float:
        """
        Query a floating-point value from the meter with retry logic.

        Parameters
        ----------
        command : str
            SCPI command to execute.
        key : str
            Logical parameter name (for logging/error reporting).

        Returns
        -------
        float
            Parsed numeric value from the meter.

        Raises
        ------
        RuntimeError
            If all retry attempts fail.
        """
        if self.mock:
            return self._mock_read(key)

        if not self.inst:
            raise RuntimeError("Meter not connected")

        last_error = None

        for attempt in range(1, self.retry_count + 1):
            try:
                resp = self.inst.query(command).strip()
                value = float(resp)
                # self.logger.debug(f"READ OK | {command} -> {value}")
                return value

            except Exception as e:
                last_error = e
                self.logger.warning(
                    f"Read failed (attempt {attempt}) | {command}"
                )
                time.sleep(0.1)

        self.logger.error(f"READ FAILED | {command}", exc_info=True)
        raise RuntimeError(f"Failed to read {key}") from last_error

    # ------------------------------------------------------------------
    # Measurement APIs (Single Reads)
    # ------------------------------------------------------------------

    # Input (CH1 - AC Source)
    def read_voltage_in(self):              return self._query_float(":MEASure? U1", "vin")
    def read_current_in(self):              return self._query_float(":MEASure? I1", "iin")
    def read_power_in(self):                return abs(self._query_float(":MEASure? P1", "kwin") / 1000)
    def read_frequency(self):               return self._query_float(":MEASure? FREQU1", "frequency")

    # New Input Params for SMR          
    def read_pf_in(self):                   return abs(self._query_float(":MEASure? PF1", "pf"))
    def read_vthd_in(self):                 return self._query_float(":MEASure? UTHD1", "vthd_in")
    def read_ithd_in(self):                 return self._query_float(":MEASure? ITHD1", "ithd_in")
    def read_power_in_watts(self):          return abs(self._query_float(":MEASure? P1", "pin"))

    # Output (CH2 - AC for AVR, DC  for SMR)
    def read_voltage_out(self):             return self._query_float(":MEASure? U2", "vout")
    def read_current_out(self):             return self._query_float(":MEASure? I2", "iout")
    def read_power_out(self):               return self._query_float(":MEASure? P2", "kwout") / 1000
    def read_vthd_out(self):                return self._query_float(":MEASure? UTHD2", "vthd_out")
    def read_efficiency(self):              return self._query_float(":MEASure? EFF1", "efficiency")
    def read_power_out_dc_watts(self):      return self._query_float(":MEASure? PDC2", "pout")
    def read_ripple(self):                  return self._query_float(":MEASure? URF2", "ripple") * 100
    # Add explicit DC read methods      
    def read_voltage_out_dc(self):          return self._query_float(":MEASure? UDC2", "vout_dc")
    def read_current_out_dc(self):          return self._query_float(":MEASure? IDC2", "iout_dc")

    # ------------------------------------------------------------------
    # Bulk Read
    # ------------------------------------------------------------------

    def read_all(self) -> dict:
        """
        Perform a bulk read of ALL available parameters.
        Each parameter read is isolated so one failure
        does not break the entire bulk read.

        Returns
        -------
        dict
            Dictionary of all parameter names mapped to numeric values
            or None if the read failed.
        """

        data = {}

        def safe_read(key: str, reader):
            try:
                data[key] = reader()
            except Exception as e:
                self.logger.error(f"Read failed for '{key}': {e}")
                data[key] = None   # or float("nan")

        # ----------------------
        # Common Parameters
        # ----------------------
        safe_read("vin", self.read_voltage_in)
        safe_read("iin", self.read_current_in)
        safe_read("kwin", self.read_power_in)
        safe_read("vout", self.read_voltage_out)
        safe_read("iout", self.read_current_out)
        safe_read("kwout", self.read_power_out)
        safe_read("efficiency", self.read_efficiency)

        # ----------------------
        # AVR Specific
        # ----------------------
        safe_read("frequency", self.read_frequency)
        safe_read("vthd_out", self.read_vthd_out)

        # ----------------------
        # SMR Specific
        # ----------------------
        safe_read("pf", self.read_pf_in)
        safe_read("vthd_in", self.read_vthd_in)
        safe_read("ithd_in", self.read_ithd_in)
        safe_read("ripple", self.read_ripple)
        safe_read("pin", self.read_power_in_watts)
        safe_read("pout", self.read_power_out_dc_watts)

        # Fetch DC specific values
        safe_read("vout_dc", self.read_voltage_out_dc)
        safe_read("iout_dc", self.read_current_out_dc)

        return data

    # ------------------------------------------------------------------
    # Diagnostics
    # ------------------------------------------------------------------

    def health_check(self) -> dict:
        """
        Perform a basic health check on the meter connection.

        Returns
        -------
        dict
            Status information including mode, connection state, and IDN.
        """
        status = {
            "mode": "MOCK" if self.mock else "REAL",
            "connected": False
        }

        if self.mock:
            status["connected"] = True
            status["idn"] = "SIMULATED"
            return status

        try:
            status["idn"] = self.inst.query("*IDN?").strip()
            status["connected"] = True
        except Exception:
            status["idn"] = "UNAVAILABLE"

        return status
    
    def smr_health_check(self) -> dict:
        """
        Perform an automatic SMR-mode health check.

        Returns
        -------
        dict
            Diagnostic results with status flags.
        """

        report = {
            "mode": None,
            "connected": False,
            "dc_voltage_valid": False,
            "dc_current_valid": False,
            "ripple_valid": False,
            "errors": [],
            "warnings": [],
            "status": "UNKNOWN"
        }

        try:
            # -------------------------------------------------
            # 1. Connection check
            # -------------------------------------------------
            if not self.is_connected():
                self.connect()

            report["connected"] = True

            # -------------------------------------------------
            # 2. Read wiring mode (actual configuration)
            # -------------------------------------------------
            if not self.mock:
                try:
                    wiring = self.inst.query(":WIRing?").strip()
                    report["mode"] = wiring
                except Exception:
                    report["warnings"].append("Unable to read wiring mode")
            else:
                report["mode"] = "MOCK"

            # -------------------------------------------------
            # 3. Wait for stabilization
            # (manual: up to 200 ms measurement cycle)
            # -------------------------------------------------
            time.sleep(0.5)

            # -------------------------------------------------
            # 4. Check event registers for errors
            # -------------------------------------------------
            if not self.mock:
                try:
                    esr1 = int(self.inst.query(":ESR1?"))
                    esr2 = int(self.inst.query(":ESR2?"))

                    if esr1 != 0:
                        report["warnings"].append(f"CH1 event flags: {esr1}")

                    if esr2 != 0:
                        report["warnings"].append(f"CH2 event flags: {esr2}")

                except Exception:
                    report["warnings"].append("Could not read event registers")

            # -------------------------------------------------
            # 5. Read DC output values
            # -------------------------------------------------
            try:
                vdc = self.read_voltage_out_dc()
                idc = self.read_current_out_dc()
                ripple = self.read_ripple()

                # Voltage check
                if vdc is not None and vdc > 1:
                    report["dc_voltage_valid"] = True
                else:
                    report["errors"].append("Invalid DC output voltage")

                # Current check
                if idc is not None and idc > 0.1:
                    report["dc_current_valid"] = True
                else:
                    report["warnings"].append("Low or zero DC output current")

                # Ripple check
                if ripple is not None and ripple > 0.001:
                    report["ripple_valid"] = True
                else:
                    report["warnings"].append("Ripple reading is zero or invalid")

            except Exception:
                report["errors"].append("Failed to read DC output values")

            # -------------------------------------------------
            # 6. Final status decision
            # -------------------------------------------------
            if report["errors"]:
                report["status"] = "FAIL"
            elif report["warnings"]:
                report["status"] = "WARN"
            else:
                report["status"] = "PASS"

        except Exception as e:
            report["errors"].append(str(e))
            report["status"] = "FAIL"

        return report

