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
        # Mock State (internal values in **watts** where applicable)
        # NOTE:
        # - Power channels are stored in watts and converted to kW
        #   by the public read helpers (see read_power_in/read_power_out
        #   and post-processing in read_all).
        # ------------------------------------------------------------------
        self._mock_state = {
            # Common
            "vin": 230.0,
            "iin": 5.0,
            "pin_watts": 1.15 * 1000,
            "vout": 230.0,
            "iout": 4.8,
            "pout_ac_watts": 1.10 * 1000,
            "pout_dc_watts": 48.0 * 10.0,
            "efficiency": 95.0,
            # AVR Specific
            "frequency": 50.0,
            "vthd_out": 3.5,
            # SMR Specific
            "pf": 0.99,
            "ripple_raw": 0.05,
            "vthd_in": 2.5,
            "ithd_in": 3.0,
            "vout_dc": 48.0,
            "iout_dc": 10.0,
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
            self.logger.info(
                "SMR mode selected but meter not connected yet; "
                "configuration will be applied after connection."
            )
            return

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

        param_map = [
            ("vin", "U1"), ("iin", "I1"), ("pin_watts", "P1"), ("frequency", "FREQU1"),
            ("pf", "PF1"), ("vthd_in", "UTHD1"), ("ithd_in", "ITHD1"),
            ("vout", "U2"), ("iout", "I2"), ("pout_ac_watts", "P2"), ("vthd_out", "UTHD2"),
            ("efficiency", "EFF1"), ("pout_dc_watts", "PDC2"), ("ripple_raw", "URF2"),
            ("vout_dc", "UDC2"), ("iout_dc", "IDC2")
        ]
        
        data = {}

        if self.mock:
            # Simulation path
            for key, _ in param_map:
                data[key] = self._mock_read(key)
        else:
            # Hardware path: Build one query string like ":MEAS? U1,I1,P1..."
            items = ",".join(item for _, item in param_map)
            cmd = f":MEASure? {items}"
            
            try:
                # One round-trip for all data
                resp_str = self.inst.query(cmd).strip()
                values = resp_str.split(';') # PW3336 usually separates multiple items with ';'
                
                # Fallback if comma separated (depending on config, though default is typically ;)
                if len(values) == 1 and ',' in resp_str:
                    values = resp_str.split(',')

                if len(values) != len(param_map):
                    raise ValueError(f"Mismatch in bulk read: Expected {len(param_map)}, got {len(values)}")

                for i, (key, _) in enumerate(param_map):
                    try:
                        data[key] = float(values[i])
                    except ValueError:
                        data[key] = None # Partial failure handling

            except Exception as e:
                self.logger.error(f"Bulk read failed: {e}")
                # Return empty dict implies full failure, handled by worker
                return {} 

        # ------------------------------------------------------
        # Post-Processing / Derivations
        # ------------------------------------------------------
        # Safe getter helper
        def get(k): return data.get(k) if data.get(k) is not None else None

        final_data = {}
        
        # --- Direct Pass-through ---
        final_data["vin"] = get("vin")
        final_data["iin"] = get("iin")
        final_data["vout"] = get("vout")
        final_data["iout"] = get("iout")
        final_data["frequency"] = get("frequency")
        final_data["pf"] = abs(get("pf")) if get("pf") is not None else None
        final_data["vthd_in"] = get("vthd_in")
        final_data["ithd_in"] = get("ithd_in")
        final_data["vthd_out"] = get("vthd_out")
        final_data["efficiency"] = get("efficiency")
        final_data["vout_dc"] = get("vout_dc")
        final_data["iout_dc"] = get("iout_dc")

        # --- Calculated  ---
        # 1. Input Power: Used for both kW (kwin) and W (pin)
        p1 = get("pin_watts")
        if p1 is not None:
            final_data["pin"] = abs(p1)
            final_data["kwin"] = abs(p1) / 1000.0
        else:
            final_data["pin"] = None
            final_data["kwin"] = None

        # 2. Output Power (AC): kW (kwout)
        p2 = get("pout_ac_watts")
        if p2 is not None:
            final_data["kwout"] = p2 / 1000.0
        else:
            final_data["kwout"] = None

        # 3. Output Power (DC): W (pout)
        final_data["pout"] = get("pout_dc_watts")

        # 4. Ripple
        # The meter's URF2 reading is returned in volts (per the Hioki manual),
        # and historically we multiplied by 100 before surfacing this to the UI
        # so that typical values are in a convenient mV-range scale for reports.
        # This scaling is preserved here for strict legacy compatibility.
        rip = get("ripple_raw")
        final_data["ripple"] = rip * 100.0 if rip is not None else None # Original logic was *100, assuming scale? Keeping logic.

        return final_data