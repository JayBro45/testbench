#delete this file

import time
import serial
from typing import Dict, Any
from threading import Lock
from logger import get_logger


class RishMultiMeter:
    def __init__(self, config: Dict[str, Any]):
        self.logger = get_logger("RishMultiMeter")

        meter_cfg = config.get("meter", {})
        rish_cfg = meter_cfg.get("rish", {})

        self.port = rish_cfg.get("port", "COM3")
        self.baud = rish_cfg.get("baud", 9600)
        self.timeout = rish_cfg.get("timeout", 1.0)

        self.ser = None
        self.lock = Lock()

        self.cache = {
            "vin": 0.0,
            "iin": 0.0,
            "vout": 0.0,
            "iout": 0.0,
            "freq": 0.0
        }

    #region agent log
    def _agent_debug_log(self, run_id: str, hypothesis_id: str, location: str, message: str, data: Dict[str, Any]):
        """
        Lightweight NDJSON logger for debug-mode hypotheses.
        """
        try:
            import json as _json
            import time as _time

            log_path = r"d:\Jayant\Testing Automation\Testbench Software\.cursor\debug.log"
            ts_ms = int(_time.time() * 1000)
            entry = {
                "id": f"log_{ts_ms}",
                "timestamp": ts_ms,
                "location": location,
                "message": message,
                "data": data,
                "runId": run_id,
                "hypothesisId": hypothesis_id,
            }
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(_json.dumps(entry) + "\n")
        except Exception:
            # Instrumentation must never break main logic
            pass
    #endregion

    # -----------------------------------------------------
    # Connection
    # -----------------------------------------------------
    def connect(self):
        if self.ser and self.ser.is_open:
            return

        try:
            self.logger.info(f"Connecting to Rishabh on {self.port}...")
            self.ser = serial.Serial(
                port=self.port,
                baudrate=self.baud,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=self.timeout
            )

            self.ser.dtr = True
            self.ser.rts = True
            time.sleep(1.0)

            self.logger.info("Connected.")

        except Exception as e:
            self.logger.error(f"Connection Failed: {e}")
            self.ser = None

    def disconnect(self):
        if self.ser:
            try:
                self.ser.close()
            except Exception:
                pass
            self.ser = None

    # -----------------------------------------------------
    # Public API
    # -----------------------------------------------------
    def read_all(self) -> Dict[str, float]:
        if not self.ser or not self.ser.is_open:
            return self.cache

        try:
            raw_data = self.ser.read(self.ser.in_waiting or 100)
            if raw_data:
                self._parse_stream(raw_data)

        except serial.SerialException:
            self.logger.error("Serial lost. Reconnecting...")
            self.disconnect()
            self.connect()

        except Exception as e:
            self.logger.error(f"Read Error: {e}")

        return self.cache

    # -----------------------------------------------------
    # Parser
    # -----------------------------------------------------
    def _parse_stream(self, data: bytes):
        #region agent log
        self._agent_debug_log(
            run_id="initial",
            hypothesis_id="H1",
            location="meter_rish_multi.py:_parse_stream",
            message="parse_stream entry",
            data={
                "length": len(data),
                "first_bytes": list(data[:8]),
            },
        )
        #endregion

        i = 0
        length = len(data)

        while i < length:
            byte = data[i]

            if 0xC0 <= byte <= 0xDF:
                val_str = ""
                j = i + 1
                digit_count = 0

                while j < length and digit_count < 8:
                    b = data[j]

                    if 0xF0 <= b <= 0xF9:
                        val_str += str(b & 0x0F)
                        digit_count += 1

                    elif 0xE0 <= b <= 0xEF:
                        pass

                    elif 0xC0 <= b <= 0xDF:
                        break

                    j += 1

                if val_str:
                    #region agent log
                    self._agent_debug_log(
                        run_id="initial",
                        hypothesis_id="H2",
                        location="meter_rish_multi.py:_parse_stream",
                        message="digit sequence extracted",
                        data={"digits": val_str},
                    )
                    #endregion
                    self._assign_value(val_str)

                i = j
            else:
                i += 1

    # -----------------------------------------------------
    # Decoding logic
    # -----------------------------------------------------
    def _safe_set(self, key: str, value: float):
        with self.lock:
            self.cache[key] = round(value, 3)

    def _assign_value(self, digits: str):
        if len(digits) < 3:
            return

        first = digits[0]

        #region agent log
        self._agent_debug_log(
            run_id="initial",
            hypothesis_id="H3",
            location="meter_rish_multi.py:_assign_value",
            message="assign_value called",
            data={"digits": digits, "first": first},
        )
        #endregion

        try:
            # Frequency
            if first == "5" and len(digits) >= 4:
                core = digits[-4:]
                rev = core[::-1]
                freq = float(rev) / 100.0
                self._safe_set("freq", freq)
                return

            # All others
            core = digits[1:-2]
            rev = core[::-1]
            value = float(rev) / 100.0

            if first == "1":
                self._safe_set("vin", value)

            elif first == "2":
                self._safe_set("iin", value)

            elif first == "3":
                self._safe_set("vout", value)

            elif first == "4":
                self._safe_set("iout", value)

        except ValueError:
            pass


# -----------------------------------------------------
# Test block
# -----------------------------------------------------
if __name__ == "__main__":
    meter = RishMultiMeter({"meter": {"rish": {"port": "COM3"}}})
    meter.connect()

    print("Reading... (Ctrl+C to stop)")
    try:
        while True:
            print(meter.read_all())
            time.sleep(1)
    except KeyboardInterrupt:
        meter.disconnect()
