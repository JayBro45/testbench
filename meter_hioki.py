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
        self.port = meter_cfg.get("port", 5025)
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
            "kwin": 1.15,   # kW
            "vout": 230.0,
            "iout": 4.8,
            "kwout": 1.10,  # kW
            "efficiency": 95.0,
            
            # AVR Specific
            "frequency": 50.0,
            "vthd_out": 3.5,
            
            # SMR Specific
            "pf": 0.99,
            "ripple": 0.05,
            "vthd_in": 2.5,
            "ithd_in": 3.0
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
        Configure meter for AVR Testing (AC Input / AC Output).
        
        Commands:
        - Rectifier Mode Ch1: RMS
        - Rectifier Mode Ch2: RMS
        """
        if self.mock: 
            self.logger.info("MOCK: Meter configured for AVR (AC/AC)")
            return

        try:
            self.inst.write(":RECTifier:MODE 1, RMS") 
            self.inst.write(":RECTifier:MODE 2, RMS")
            self.logger.info("Meter configured for AVR (AC/AC)")
        except Exception as e:
            self.logger.error(f"Failed to set AVR mode: {e}")

    def set_mode_smr(self):
        """
        Configure meter for SMR Testing (AC Input / DC Output).
        
        Commands:
        - Rectifier Mode Ch1: RMS (AC Input)
        - Rectifier Mode Ch2: MEAN (DC Output)
        """
        if self.mock: 
            self.logger.info("MOCK: Meter configured for SMR (AC/DC)")
            return

        try:
            self.inst.write(":RECTifier:MODE 1, RMS")
            self.inst.write(":RECTifier:MODE 2, MEAN") 
            self.logger.info("Meter configured for SMR (AC/DC)")
        except Exception as e:
            self.logger.error(f"Failed to set SMR mode: {e}")

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
    def read_voltage_in(self):   return self._query_float(":MEASure? U1", "vin")
    def read_current_in(self):   return self._query_float(":MEASure? I1", "iin")
    def read_power_in(self):     return abs(self._query_float(":MEASure? P1", "kwin") / 1000)
    def read_frequency(self):    return self._query_float(":MEASure? FREQU1", "frequency")
    
    # New Input Params for SMR (REQUIRED FOR SMR STRATEGY)
    def read_pf_in(self):        return self._query_float(":MEASure? PF1", "pf")
    def read_vthd_in(self):      return self._query_float(":MEASure? UTHD1", "vthd_in")
    def read_ithd_in(self):      return self._query_float(":MEASure? ITHD1", "ithd_in")

    # Output (CH2 - AC for AVR, DC for SMR)
    def read_voltage_out(self):  return self._query_float(":MEASure? U2", "vout")
    def read_current_out(self):  return self._query_float(":MEASure? I2", "iout")
    def read_power_out(self):    return self._query_float(":MEASure? P2", "kwout") / 1000
    def read_vthd_out(self):     return self._query_float(":MEASure? UTHD2", "vthd_out")
    def read_efficiency(self):   return self._query_float(":MEASure? EFF1", "efficiency")

    def read_ripple(self):
        """
        Read Ripple Voltage (AC component on DC line).
        Assumes Ch2 is Output.
        """
        if self.mock: return self._mock_read("ripple")
        # Read AC component of Ch2 Voltage
        return self._query_float(":MEASure? UAC2", "ripple")

    # ------------------------------------------------------------------
    # Bulk Read
    # ------------------------------------------------------------------

    def read_all(self) -> dict:
        """
        Perform a bulk read of ALL available parameters.
        The UI/Strategy layer decides which values to display/store.

        Returns
        -------
        dict
            Dictionary of all parameter names mapped to numeric values.
        """
        # self.logger.debug("Bulk read started")
        data = {}

        try:
            # Common Parameters
            data["vin"] = self.read_voltage_in()
            data["iin"] = self.read_current_in()
            data["kwin"] = self.read_power_in()
            data["vout"] = self.read_voltage_out()
            data["iout"] = self.read_current_out()
            data["kwout"] = self.read_power_out()
            data["efficiency"] = self.read_efficiency()
            
            # AVR Specific
            data["frequency"] = self.read_frequency()
            data["vthd_out"] = self.read_vthd_out()
            
            # SMR Specific (CRITICAL: Added missing fields)
            data["pf"] = self.read_pf_in()
            data["vthd_in"] = self.read_vthd_in()
            data["ithd_in"] = self.read_ithd_in()
            data["ripple"] = self.read_ripple()
            
        except Exception as e:
            self.logger.error(f"Bulk read error: {e}")
            # In a real scenario, you might want to raise this or return partial data
            # For now, we return what we have to keep the UI responsive
            
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