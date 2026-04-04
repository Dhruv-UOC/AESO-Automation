"""
core/psse_interface.py
----------------------
Central interface for initializing and communicating with PSS/E via psspy.
All study modules import from here — never import psspy directly elsewhere.

Usage:
    from core.psse_interface import PSSEInterface
    psse = PSSEInterface(psse_path=r"C:\\Program Files\\PTI\\PSSE35\\PSSBIN")
    psse.initialize()
    psse.load_case("path/to/case.sav")
"""

import os
import sys
import logging

logger = logging.getLogger(__name__)


class PSSEInitError(Exception):
    """Raised when PSSE cannot be initialized."""


class PSSEInterface:
    """
    Thin wrapper around psspy that handles initialization, version detection,
    and graceful fallback to a mock mode when PSSE is not installed.

    Parameters
    ----------
    psse_path : str
        Path to the PSSE PSSBIN directory, e.g.
        r"C:\\Program Files\\PTI\\PSSE35\\PSSBIN"
    psse_version : int
        Major PSSE version number (33, 34, 35 …). Used to pick the
        correct psspy import path.  Set to None for auto-detect.
    mock : bool
        If True, skip real PSSE and use a stub.  Useful for unit tests
        on machines without a PSSE licence.
    """

    SUPPORTED_VERSIONS = [33, 34, 35]

    def __init__(self, psse_path: str = None, psse_version: int = 35, mock: bool = False):
        self.psse_path = psse_path
        self.psse_version = psse_version
        self.mock = mock
        self._psspy = None
        self._initialized = False

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def initialize(self, buses: int = 50000, output_level: int = 6) -> None:
        """
        Start a PSSE session.

        Parameters
        ----------
        buses : int
            Maximum bus count for the session (default 50 000).
        output_level : int
            PSSE output verbosity (6 = suppress most messages).
        """
        if self.mock:
            logger.warning("PSSEInterface running in MOCK mode — no real PSSE calls.")
            self._initialized = True
            return

        self._add_psse_to_path()
        try:
            import psspy  # type: ignore
            self._psspy = psspy
        except ImportError as exc:
            raise PSSEInitError(
                f"Could not import psspy. Verify PSSE is installed and "
                f"psse_path='{self.psse_path}' is correct."
            ) from exc

        ret = self._psspy.psseinit(buses)
        if ret != 0:
            raise PSSEInitError(f"psspy.psseinit() returned error code {ret}")

        # Suppress PSSE progress/output window
        self._psspy.report_output(6, "", [])
        self._psspy.progress_output(6, "", [])

        self._initialized = True
        logger.info("PSSE initialized (version %s, buses=%d)", self.psse_version, buses)

    def load_case(self, sav_path: str) -> None:
        """Load a .sav power flow case file."""
        self._check_initialized()
        sav_path = os.path.abspath(sav_path)
        
        if self.mock:
            logger.info("[MOCK] load_case('%s')", sav_path)
            return

        if not os.path.isfile(sav_path):
            raise FileNotFoundError(f"Case file not found: {sav_path}")

        ret = self._psspy.case(sav_path)
        self._check_return(ret, "case()")
        logger.info("Loaded case: %s", sav_path)

    def load_dynamics(self, dyr_path: str) -> None:
        """Load a .dyr dynamic model file (required for transient stability)."""
        self._check_initialized()
        if self.mock:
            logger.info("[MOCK] load_dynamics('%s')", dyr_path)
            return

        ret = self._psspy.dyre_new([1, 1, 1, 1], dyr_path, "", "", "")
        self._check_return(ret, "dyre_new()")
        logger.info("Loaded dynamics: %s", dyr_path)

    def save_case(self, sav_path: str) -> None:
        """Save the current working case to a .sav file."""
        self._check_initialized()
        if self.mock:
            logger.info("[MOCK] save_case('%s')", sav_path)
            return

        ret = self._psspy.save(sav_path)
        self._check_return(ret, "save()")

    @property
    def psspy(self):
        """Direct access to the psspy module for advanced calls."""
        if self.mock:
            raise RuntimeError("psspy not available in mock mode.")
        return self._psspy

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _add_psse_to_path(self) -> None:
        """Add PSSE binaries and Python bindings to system path."""
        # For Python 3.7, add PSSPY37 path
        psspy_path = os.path.join(self.psse_path, "..", "PSSPY37")
        psspy_path = os.path.abspath(psspy_path)
        
        if psspy_path and psspy_path not in sys.path:
            sys.path.insert(0, psspy_path)
        
        if self.psse_path and self.psse_path not in sys.path:
            sys.path.insert(0, self.psse_path)
            os.environ["PATH"] = self.psse_path + os.pathsep + os.environ.get("PATH", "")

    def _check_initialized(self) -> None:
        if not self._initialized:
            raise RuntimeError("Call PSSEInterface.initialize() before using PSSE.")

    @staticmethod
    def _check_return(ret: int, fn_name: str) -> None:
        if ret != 0:
            raise RuntimeError(f"psspy.{fn_name} returned error code {ret}")
