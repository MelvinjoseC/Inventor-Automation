import logging
import os
import subprocess
import time
from typing import Any, Optional

try:
    import win32com.client as win32
    from win32com.client import gencache
except ImportError:
    win32 = None
    gencache = None

logger = logging.getLogger("InventorAutomation.Connection")


class COMErrorHelper:
    """Provides formatting and diagnostics for Windows COM and PyWin32 exceptions."""

    @staticmethod
    def format_error(exception: Exception) -> str:
        """Extracts structured diagnostic details from COM errors."""
        # Check if exception has COM error details (hresult, details tuple)
        args = getattr(exception, "args", ())
        if len(args) >= 3 and isinstance(args[2], tuple):
            hresult = args[0]
            msg = args[1]
            details = args[2]
            # pywin32 format: (wcode, source, description, helpfile, helpcontext, scode)
            desc = details[2] if len(details) > 2 else "No description"
            source = details[1] if len(details) > 1 else "Unknown source"
            return f"COM Error [HRESULT 0x{hresult & 0xFFFFFFFF:08X}] from '{source}': {desc} (Message: {msg})"
        return str(exception)


class TransactionContext:
    """Manages an atomic Inventor transaction.

    If an error occurs within the context block, the transaction is automatically aborted,
    rolling back the document state to prevent corruption.
    """

    def __init__(self, inventor_app: Any, doc: Any, transaction_name: str = "AutomatedOperation"):
        self.app = inventor_app
        self.doc = doc
        self.name = transaction_name
        self.transaction: Optional[Any] = None

    def __enter__(self) -> "TransactionContext":
        try:
            tm = getattr(self.app, "TransactionManager", None)
            if tm and hasattr(tm, "StartTransaction"):
                self.transaction = tm.StartTransaction(self.doc, self.name)
                logger.debug(f"Started transaction: '{self.name}'")
        except Exception as e:
            logger.warning(f"Could not start transaction '{self.name}': {e}")
            self.transaction = None
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if self.transaction:
            try:
                if exc_type is not None:
                    logger.warning(f"Aborting transaction '{self.name}' due to error: {exc_val}")
                    self.transaction.Abort()
                else:
                    self.transaction.End()
                    logger.debug(f"Committed transaction: '{self.name}'")
            except Exception as e:
                logger.error(f"Error finalizing transaction '{self.name}': {e}")


class PerformanceMode:
    """Context manager to suppress UI updates, animations, and prompt dialogs
    during heavy batch CAD generation to achieve maximum throughput.
    """

    def __init__(self, inventor_app: Any, enable_turbo: bool = True):
        self.app = inventor_app
        self.enable_turbo = enable_turbo
        self._prev_screen_updating = True
        self._prev_silent = False

    def __enter__(self) -> "PerformanceMode":
        if self.enable_turbo and self.app:
            try:
                if hasattr(self.app, "ScreenUpdating"):
                    self._prev_screen_updating = self.app.ScreenUpdating
                    self.app.ScreenUpdating = False
                if hasattr(self.app, "SilentOnSelectSet"):
                    self._prev_silent = self.app.SilentOnSelectSet
                    self.app.SilentOnSelectSet = True
                logger.debug("Enabled PerformanceMode: Screen updating disabled.")
            except Exception as e:
                logger.warning(f"Failed to enable performance mode: {e}")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if self.enable_turbo and self.app:
            try:
                if hasattr(self.app, "ScreenUpdating"):
                    self.app.ScreenUpdating = self._prev_screen_updating
                if hasattr(self.app, "SilentOnSelectSet"):
                    self.app.SilentOnSelectSet = self._prev_silent
                logger.debug("Restored previous UI and screen updating settings.")
            except Exception as e:
                logger.warning(f"Failed to restore screen updating settings: {e}")


class InventorConnectionManager:
    """Manages connection lifecycle with Autodesk Inventor COM API."""

    PROGID = "Inventor.Application"

    def __init__(
        self,
        visible: bool = True,
        silent_mode: bool = False,
        retries: int = 3,
        delay_seconds: float = 2.0,
    ):
        self.visible = visible
        self.silent_mode = silent_mode
        self.retries = retries
        self.delay_seconds = delay_seconds
        self.app: Optional[Any] = None

    def connect(self) -> Any:
        """Connects to a running Inventor instance or starts a new session with retries."""
        if win32 is None:
            raise RuntimeError("pywin32 is not installed or available on this platform.")

        for attempt in range(1, self.retries + 1):
            try:
                logger.info(
                    f"Attempting connection to Autodesk Inventor (Attempt {attempt}/{self.retries})..."
                )

                # First, try to attach to a running Inventor instance
                try:
                    self.app = win32.GetObject(Class=self.PROGID)
                    logger.info("Successfully attached to a running Autodesk Inventor instance.")
                except Exception:
                    # If attachment fails, try dispatching a new instance
                    logger.info(
                        "No running instance detected. Dispatched new Inventor application."
                    )
                    try:
                        self.app = gencache.EnsureDispatch(self.PROGID)
                    except Exception as e_dispatch:
                        logger.warning(
                            f"EnsureDispatch failed: {e_dispatch}. Retrying standard dynamic Dispatch."
                        )
                        self.app = win32.Dispatch(self.PROGID)

                if self.app:
                    # Apply configurations
                    self.app.Visible = self.visible
                    self.app.SilentOnSelectSet = self.silent_mode

                    # Log software version
                    version = getattr(
                        getattr(self.app, "SoftwareVersion", None), "DisplayVersion", "Unknown"
                    )
                    logger.info(f"Connected to Autodesk Inventor. Version: {version}")
                    return self.app

            except Exception as e:
                err_msg = COMErrorHelper.format_error(e)
                logger.warning(f"Connection attempt {attempt} failed: {err_msg}")
                if attempt < self.retries:
                    logger.info(f"Waiting {self.delay_seconds} seconds before next attempt...")
                    time.sleep(self.delay_seconds)
                else:
                    logger.error(
                        "Exceeded maximum connection retry attempts. Autodesk Inventor could not be loaded."
                    )
                    raise ConnectionError(
                        f"Failed to connect to Autodesk Inventor COM API: {err_msg}"
                    ) from e

    def is_connected(self) -> bool:
        """Verifies if the connection to Inventor is alive by calling a lightweight property."""
        if not self.app:
            return False
        try:
            _ = self.app.SoftwareVersion.DisplayVersion
            return True
        except Exception:
            return False

    def disconnect(self, close_app: bool = False) -> None:
        """Releases the COM reference. Optionally closes the application."""
        if self.app:
            try:
                if close_app:
                    logger.info("Closing Autodesk Inventor application...")
                    self.app.Quit()
                self.app = None
                logger.info("Successfully disconnected and cleaned COM references.")
            except Exception as e:
                logger.error(f"Error during disconnect: {e}")
                self.app = None

    @staticmethod
    def kill_orphaned_processes() -> int:
        """Forcefully kills any orphaned Inventor.exe processes. Useful for CI cleanup."""
        killed = 0
        try:
            output = subprocess.check_output(
                ["taskkill", "/F", "/IM", "Inventor.exe"],
                stderr=subprocess.DEVNULL,
                universal_newlines=True,
            )
            killed += 1
            logger.info("Killed orphaned Inventor.exe process(es).")
        except Exception:
            pass
        return killed

    def __enter__(self) -> Any:
        return self.connect()

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.disconnect(close_app=False)
