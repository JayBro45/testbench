"""
worker.py
=========

Background polling worker for AVR Test Bench Software.

This module defines a QThread-based worker responsible for:
- Managing the meter connection lifecycle
- Polling measurements at fixed intervals
- Emitting structured measurement data to the UI layer
- Detecting transient vs fatal communication failures
- Shutting down safely on request

Design Principles
-----------------
- The worker owns the meter while running
- No UI elements are accessed directly
- No acceptance logic or reporting logic exists here
- All meter-specific behavior is delegated to the meter driver
- Thread termination is cooperative and mutex-protected

This worker is designed to be **robust under real hardware conditions**,
including intermittent communication failures.
"""

from PySide6.QtCore import QThread, Signal, QMutex
import time
from logger import get_logger  # <--- NEW IMPORT


class MeterPollingWorker(QThread):
    """
    Background thread that continuously polls the connected meter.

    Responsibilities
    ----------------
    - Establish meter connection (if not already connected)
    - Periodically read all required measurements
    - Emit data, warnings, and fatal errors via Qt signals
    - Detect loss of communication and terminate safely
    - Ensure the meter is always disconnected on exit

    Signals
    -------
    data_ready : Signal(dict)
        Emitted when a complete measurement snapshot is available.

    warning : Signal(str)
        Emitted when a recoverable read error occurs.

    error : Signal(str)
        Emitted when a fatal error occurs and polling must stop.

    status : Signal(str)
        Emitted for high-level lifecycle status updates.

    finished_polling : Signal()
        Emitted exactly once when the worker exits (guaranteed).
    """

    # ------------------------------------------------------------------
    # Qt Signals
    # ------------------------------------------------------------------
    data_ready = Signal(dict)
    warning = Signal(str)
    error = Signal(str)
    status = Signal(str)
    finished_polling = Signal()   # GUARANTEED lifecycle signal

    # ------------------------------------------------------------------
    # Initialization
    # ------------------------------------------------------------------
    def __init__(self, meter, interval_sec: float = 1.0):
        """
        Initialize the polling worker.

        Parameters
        ----------
        meter : object
            Meter abstraction instance (e.g., Hioki meter driver).
            Must implement:
            - connect()
            - disconnect()
            - is_connected()
            - read_all()

        interval_sec : float, optional
            Polling interval in seconds (default: 1.0).
        """
        super().__init__()
        self.meter = meter
        self.interval = interval_sec
        self._running = True
        self._mutex = QMutex()
        self.logger = get_logger("MeterWorker")  # <--- NEW LOGGER

    # ------------------------------------------------------------------
    # Thread Entry Point
    # ------------------------------------------------------------------
    def run(self):
        """
        Main execution loop of the worker thread.

        Execution Phases
        ----------------
        1. Connection Phase
           - Establish meter connection if required

        2. Polling Phase
           - Read measurements at fixed intervals
           - Track consecutive failures
           - Escalate to fatal error on repeated failure

        3. Shutdown Phase
           - Disconnect meter
           - Emit finished_polling signal
        """

        # ---------- PHASE 1: CONNECT ----------
        try:
            self.status.emit("Connecting to meter...")
            self.logger.info("Worker starting - attempting connection")  # <--- LOG

            if not self.meter.is_connected():
                self.meter.connect()

            self.status.emit("Connected - Monitoring")
            self.logger.info("Meter connected successfully")  # <--- LOG

        except Exception as e:
            msg = f"Connection Failed: {e}"
            self.logger.error(msg, exc_info=True)  # <--- LOG TRACEBACK
            self.error.emit(msg)
            self.finished_polling.emit()
            return   # CLEAN EXIT

        # ---------- PHASE 2: POLLING ----------
        consecutive_errors = 0
        MAX_RETRIES = 5

        try:
            while self.is_running():
                try:
                    data = self.meter.read_all()
                    if not data:
                        raise ValueError("Empty data")

                    self.data_ready.emit(data)
                    
                    # Reset counter on success
                    if consecutive_errors > 0:
                        self.logger.info("Communication recovered")
                        consecutive_errors = 0

                except Exception as e:
                    consecutive_errors += 1
                    msg = f"Read missed ({consecutive_errors}/{MAX_RETRIES})"
                    self.warning.emit(msg)
                    self.logger.warning(f"{msg} | Error: {e}")  # <--- LOG WARNING

                    if consecutive_errors >= MAX_RETRIES:
                        raise ConnectionError("Lost connection to meter")

                self.smart_sleep(self.interval)

        except Exception as e:
            self.logger.error(f"Fatal Polling Error: {e}", exc_info=True)  # <--- LOG FATAL
            self.error.emit(f"Fatal Error: {e}")

        finally:
            self.safe_disconnect()
            self.finished_polling.emit()

    # ------------------------------------------------------------------
    # Control & State Helpers
    # ------------------------------------------------------------------
    def stop(self):
        """
        Request the worker thread to stop.

        This method is thread-safe and non-blocking.
        """
        self.logger.info("Worker stop requested")
        self._mutex.lock()
        self._running = False
        self._mutex.unlock()

    def is_running(self) -> bool:
        """
        Check whether the worker should continue running.

        Returns
        -------
        bool
            True if polling should continue, False otherwise.
        """
        self._mutex.lock()
        state = self._running
        self._mutex.unlock()
        return state

    # ------------------------------------------------------------------
    # Utility Helpers
    # ------------------------------------------------------------------
    def smart_sleep(self, duration: float):
        """
        Sleep in small increments so stop requests are handled quickly.

        Parameters
        ----------
        duration : float
            Total sleep duration in seconds.
        """
        steps = int(duration / 0.1)
        for _ in range(steps):
            if not self.is_running():
                return
            time.sleep(0.1)

    def safe_disconnect(self):
        """
        Safely disconnect the meter, ignoring any errors.

        This method is guaranteed to be called on thread exit.
        """
        self.logger.info("Cleaning up connection...")
        try:
            self.meter.disconnect()
        except Exception as e:
            self.logger.warning(f"Disconnect error (ignored): {e}")
        self.status.emit("Stopped")