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

Backward Compatibility:
- All original function names preserved
- No public API changes
"""

import pyvisa
import time
import random
from logger import get_logger


class HiokiPW3336:
    """
    Driver class for the Hioki PW3336 Power Analyzer.
    """

    # Central SCPI command map (internal optimization)
    SCPI_MAP = {
        "vin": ":MEASure? U1",
        "iin": ":MEASure? I1",
        "pin": ":MEASure? P1",
        "frequency": ":MEASure? FREQU1",
        "pf": ":MEASure? PF1",
        "vthd_in": ":MEASure? UTHD1",
        "ithd_in": ":MEASure? ITHD1",
        "vout": ":MEASure? U2",
        "iout": ":MEASure? I2",
        "pout": ":MEASure? P2",
        "vthd_out": ":MEASure? UTHD2",
        "efficiency": ":MEASure? EFF1",
        "vout_dc": ":MEASure? UDC2",
        "iout_dc": ":MEASure? IDC2",
        "pout_dc": ":MEASure? PDC2",
        "ripple": ":MEASure? URF2",
    }

    # ------------------------------------------------------------------
    # Initialization
    # ------------------------------------------------------------------

    def __init__(self, config: dict):
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

        # Mock base state (internal units: volts, amps, watts)
        self._mock_state = {
            "vin": 230.0,
            "iin": 5.0,
            "pin": 1150.0,
            "vout": 230.0,
            "iout": 4.8,
            "pout": 1100.0,
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
        self.connect()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.disconnect()

    # ------------------------------------------------------------------
    # Connection Management
    # ------------------------------------------------------------------

    def connect(self):
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
        self.logger.info("Disconnecting meter")

        if self.inst:
            self.inst.close()
            self.inst = None

        if self.rm:
            self.rm.close()
            self.rm = None

    def is_connected(self) -> bool:
        return self.mock or self.inst is not None

    # ------------------------------------------------------------------
    # Mode Configuration (unchanged)
    # ------------------------------------------------------------------

    def set_mode_avr(self):
        if self.mock:
            self.logger.info("MOCK: Meter configured for AVR")
            return
        self.logger.info("Meter configured for AVR")

    def set_mode_smr(self):
        if self.mock:
            self.logger.info("MOCK: Meter configured for SMR")
            return

        if not self.inst:
            raise RuntimeError("Meter not connected")

        self.inst.write(":WIRing TYPE1")
        self.inst.write(":MEASure:ITEM:ALLClear")
        self.inst.write(":MEASure:ITEM:UDC:CH2 ON")
        self.inst.write(":MEASure:ITEM:IDC:CH2 ON")
        self.inst.write(":MEASure:ITEM:PDC:CH2 ON")
        self.inst.write(":MEASure:ITEM:UAC:CH2 ON")

        time.sleep(0.5)
        self.logger.info("Meter configured for SMR")

    # ------------------------------------------------------------------
    # Internal Helpers
    # ------------------------------------------------------------------

    def _mock_read(self, key: str) -> float:
        base = self._mock_state.get(key, 0.0)
        noise = base * random.uniform(-0.005, 0.005)
        return round(max(0.0, base + noise), 3)

    def _query_float(self, command: str, key: str) -> float:
        if self.mock:
            return self._mock_read(key)

        if not self.inst:
            raise RuntimeError("Meter not connected")

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
    # Measurement APIs (unchanged)
    # ------------------------------------------------------------------

    def read_voltage_in(self): return self._query_float(self.SCPI_MAP["vin"], "vin")
    def read_current_in(self): return self._query_float(self.SCPI_MAP["iin"], "iin")
    def read_power_in(self): return abs(self._query_float(self.SCPI_MAP["pin"], "pin") / 1000)
    def read_frequency(self): return self._query_float(self.SCPI_MAP["frequency"], "frequency")

    def read_pf_in(self): return abs(self._query_float(self.SCPI_MAP["pf"], "pf"))
    def read_vthd_in(self): return self._query_float(self.SCPI_MAP["vthd_in"], "vthd_in")
    def read_ithd_in(self): return self._query_float(self.SCPI_MAP["ithd_in"], "ithd_in")
    def read_power_in_watts(self): return abs(self._query_float(self.SCPI_MAP["pin"], "pin"))

    def read_voltage_out(self): return self._query_float(self.SCPI_MAP["vout"], "vout")
    def read_current_out(self): return self._query_float(self.SCPI_MAP["iout"], "iout")
    def read_power_out(self): return self._query_float(self.SCPI_MAP["pout"], "pout") / 1000
    def read_vthd_out(self): return self._query_float(self.SCPI_MAP["vthd_out"], "vthd_out")
    def read_efficiency(self): return self._query_float(self.SCPI_MAP["efficiency"], "efficiency")
    def read_power_out_dc_watts(self): return self._query_float(self.SCPI_MAP["pout_dc"], "pout_dc")
    def read_ripple(self): return self._query_float(self.SCPI_MAP["ripple"], "ripple") * 100
    def read_voltage_out_dc(self): return self._query_float(self.SCPI_MAP["vout_dc"], "vout_dc")
    def read_current_out_dc(self): return self._query_float(self.SCPI_MAP["iout_dc"], "iout_dc")

    # ------------------------------------------------------------------
    # Bulk Read (backward-compatible behavior)
    # ------------------------------------------------------------------

    def read_all(self) -> dict:
        data = {}

        def safe_read(key: str, reader):
            try:
                data[key] = reader()
            except Exception:
                data[key] = None

        safe_read("vin", self.read_voltage_in)
        safe_read("iin", self.read_current_in)
        safe_read("kwin", self.read_power_in)
        safe_read("vout", self.read_voltage_out)
        safe_read("iout", self.read_current_out)
        safe_read("kwout", self.read_power_out)
        safe_read("efficiency", self.read_efficiency)

        safe_read("frequency", self.read_frequency)
        safe_read("vthd_out", self.read_vthd_out)

        safe_read("pf", self.read_pf_in)
        safe_read("vthd_in", self.read_vthd_in)
        safe_read("ithd_in", self.read_ithd_in)
        safe_read("ripple", self.read_ripple)
        safe_read("pin", self.read_power_in_watts)
        safe_read("pout", self.read_power_out_dc_watts)

        safe_read("vout_dc", self.read_voltage_out_dc)
        safe_read("iout_dc", self.read_current_out_dc)

        return data

    # ------------------------------------------------------------------
    # Diagnostics (unchanged)
    # ------------------------------------------------------------------

    def health_check(self) -> dict:
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
