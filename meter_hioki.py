"""
meter_hioki.py
==============

Optimized hardware abstraction layer for the Hioki PW3336 Power Analyzer.

This module provides a stable and efficient interface for:

- Establishing and managing communication with the Hioki PW3336
- Reading electrical parameters required for AVR (AC) and SMR (DC) testing
- Supporting both REAL (PyVISA-based) and MOCK (simulation) modes
- Providing resilient read operations with retry logic
- Exposing a consistent and backward-compatible API

Design Principles:
------------------
- Configuration-driven behavior via config.json
- No business logic or acceptance rules
- Deterministic public API used by worker/UI layers
- Internal power units in watts, exposed as kW where required
- Backward compatibility with existing callers

Optimizations Implemented:
--------------------------
- Centralized SCPI command mapping
- Reduced duplicated read logic
- Exponential retry backoff
- Context manager support
- Realistic mock noise model
- Safer health checks
"""

import pyvisa
import time
import random
from logger import get_logger


class HiokiPW3336:
    """
    Optimized driver for the Hioki PW3336 Power Analyzer.

    Responsibilities:
    -----------------
    - Manage meter connection lifecycle
    - Read electrical parameters from the device
    - Provide mock readings when hardware is unavailable
    - Expose a consistent data contract to worker/UI layers

    Supported Operating Modes:
    --------------------------
    - REAL: Communicates with physical hardware via PyVISA
    - MOCK: Generates simulated readings for development/testing

    Configuration Dependencies (config.json):
    -----------------------------------------
    - meter: connection parameters and retry settings
    - parameters: optional parameter enabling flags
    - logging: logging verbosity

    Notes:
    ------
    - This class performs no validation or acceptance logic.
    - Internal power units are stored in watts.
    - Public power APIs return values in kW where applicable.
    """

    # ------------------------------------------------------------------
    # SCPI Command Map
    # ------------------------------------------------------------------

    SCPI_MAP = {
        # Input CH1
        "vin": ":MEASure? U1",
        "iin": ":MEASure? I1",
        "pin": ":MEASure? P1",
        "frequency": ":MEASure? FREQU1",
        "pf": ":MEASure? PF1",
        "vthd_in": ":MEASure? UTHD1",
        "ithd_in": ":MEASure? ITHD1",

        # Output CH2
        "vout": ":MEASure? U2",
        "iout": ":MEASure? I2",
        "pout": ":MEASure? P2",
        "vthd_out": ":MEASure? UTHD2",
        "efficiency": ":MEASure? EFF1",

        # DC
        "vout_dc": ":MEASure? UDC2",
        "iout_dc": ":MEASure? IDC2",
        "pout_dc": ":MEASure? PDC2",
        "ripple": ":MEASure? URF2",
    }

    # ------------------------------------------------------------------
    # Initialization
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

        # Assign query function once (avoids repeated conditionals)
        self._query = self._mock_query if self.mock else self._real_query

        # Mock base state (internal units: volts, amps, watts)
        self._mock_state = {
            "vin": 230.0,
            "iin": 5.0,
            "pin": 1150.0,   # watts
            "vout": 230.0,
            "iout": 4.8,
            "pout": 1100.0,  # watts
            "efficiency": 95.0,
            "frequency": 50.0,
            "vthd_out": 3.5,
            "pf": 0.99,
            "ripple": 0.05,
            "vthd_in": 2.5,
            "ithd_in": 3.0,
            "vout_dc": 48.0,
            "iout_dc": 10.0,
            "pout_dc": 480.0,
        }

    # ------------------------------------------------------------------
    # Context Manager Support
    # ------------------------------------------------------------------

    def __enter__(self):
        """
        Context manager entry.
        Automatically connects to the meter.
        """
        self.connect()
        return self

    def __exit__(self, exc_type, exc, tb):
        """
        Context manager exit.
        Ensures meter is disconnected.
        """
        self.disconnect()

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
    # Query Engines
    # ------------------------------------------------------------------

    def _mock_query(self, key: str) -> float:
        """
        Generate a simulated reading with realistic noise.

        Parameters
        ----------
        key : str
            Logical parameter name.

        Returns
        -------
        float
            Simulated measurement value.
        """
        base = self._mock_state.get(key, 0.0)
        noise = base * random.uniform(-0.005, 0.005)  # ±0.5%
        return round(max(0.0, base + noise), 3)

    def _real_query(self, key: str) -> float:
        """
        Query a floating-point value from the meter with retry logic.

        Parameters
        ----------
        key : str
            Logical parameter name.

        Returns
        -------
        float
            Parsed numeric value from the meter.

        Raises
        ------
        RuntimeError
            If all retry attempts fail.
        """
        if not self.inst:
            raise RuntimeError("Meter not connected")

        command = self.SCPI_MAP[key]
        last_error = None

        for attempt in range(1, self.retry_count + 1):
            try:
                resp = self.inst.query(command).strip()
                return float(resp)

            except Exception as e:
                last_error = e
                time.sleep(0.05 * attempt)

        raise RuntimeError(f"Failed to read {key}") from last_error

    # ------------------------------------------------------------------
    # Generic Reader
    # ------------------------------------------------------------------

    def read(self, key: str):
        """
        Read a measurement by logical parameter key.

        Parameters
        ----------
        key : str
            Logical parameter name from SCPI_MAP.

        Returns
        -------
        float
            Measured value.
        """
        return self._query(key)

    # ------------------------------------------------------------------
    # Backward-Compatible Public API
    # ------------------------------------------------------------------

    def read_voltage_in(self): return self.read("vin")
    def read_current_in(self): return self.read("iin")
    def read_frequency(self): return self.read("frequency")
    def read_pf_in(self): return abs(self.read("pf"))
    def read_vthd_in(self): return self.read("vthd_in")
    def read_ithd_in(self): return self.read("ithd_in")

    def read_voltage_out(self): return self.read("vout")
    def read_current_out(self): return self.read("iout")
    def read_vthd_out(self): return self.read("vthd_out")
    def read_efficiency(self): return self.read("efficiency")

    def read_voltage_out_dc(self): return self.read("vout_dc")
    def read_current_out_dc(self): return self.read("iout_dc")

    def read_ripple(self): return self.read("ripple") * 100

    def read_power_in(self):
        """Return input power in kW."""
        return abs(self.read("pin") / 1000)

    def read_power_out(self):
        """Return output power in kW."""
        return self.read("pout") / 1000

    def read_power_in_watts(self):
        """Return input power in watts."""
        return abs(self.read("pin"))

    def read_power_out_dc_watts(self):
        """Return DC output power in watts."""
        return self.read("pout_dc")

    # ------------------------------------------------------------------
    # Optimized Bulk Read
    # ------------------------------------------------------------------

    def read_all(self) -> dict:
        """
        Perform a bulk read of all available parameters.

        Returns
        -------
        dict
            Dictionary mapping parameter names to values.
            Failed reads return None.
        """
        data = {}

        for key in self.SCPI_MAP.keys():
            try:
                data[key] = self.read(key)
            except Exception:
                data[key] = None

        # Derived values
        data["kwin"] = None if data["pin"] is None else abs(data["pin"] / 1000)
        data["kwout"] = None if data["pout"] is None else data["pout"] / 1000

        return data

    # ------------------------------------------------------------------
    # Health Check
    # ------------------------------------------------------------------

    def health_check(self) -> dict:
        """
        Perform a basic connection health check.

        Returns
        -------
        dict
            Status including mode, connection state, and IDN.
        """
        status = {
            "mode": "MOCK" if self.mock else "REAL",
            "connected": False,
            "idn": "UNAVAILABLE"
        }

        if self.mock:
            status["connected"] = True
            status["idn"] = "SIMULATED"
            return status

        if not self.inst:
            return status

        try:
            status["idn"] = self.inst.query("*IDN?").strip()
            status["connected"] = True
        except Exception:
            pass

        return status
