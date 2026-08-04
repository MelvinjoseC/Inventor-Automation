import time
import logging
from typing import Any, Optional
import win32com.client as win32
from win32com.client import gencache

logger = logging.getLogger("InventorAutomation.Connection")

class InventorConnectionManager:
    """Manages connection lifecycle with Autodesk Inventor COM API."""

    PROGID = "Inventor.Application"

    def __init__(
        self, 
        visible: bool = True, 
        silent_mode: bool = False, 
        retries: int = 3, 
        delay_seconds: float = 2.0
    ):
        self.visible = visible
        self.silent_mode = silent_mode
        self.retries = retries
        self.delay_seconds = delay_seconds
        self.app: Optional[Any] = None

    def connect(self) -> Any:
        """Connects to a running Inventor instance or starts a new session with retries."""
        for attempt in range(1, self.retries + 1):
            try:
                logger.info(f"Attempting connection to Autodesk Inventor (Attempt {attempt}/{self.retries})...")
                
                # First, try to attach to a running Inventor instance
                try:
                    self.app = win32.GetObject(Class=self.PROGID)
                    logger.info("Successfully attached to a running Autodesk Inventor instance.")
                except Exception:
                    # If attachment fails, try dispatching a new instance
                    logger.info("No running instance detected. Dispatched new Inventor application.")
                    try:
                        self.app = gencache.EnsureDispatch(self.PROGID)
                    except Exception as e_dispatch:
                        logger.warning(f"EnsureDispatch failed: {e_dispatch}. Retrying standard dynamic Dispatch.")
                        self.app = win32.Dispatch(self.PROGID)

                if self.app:
                    # Apply configurations
                    self.app.Visible = self.visible
                    self.app.SilentOnSelectSet = self.silent_mode
                    
                    # Log software version
                    version = self.app.SoftwareVersion.DisplayVersion
                    logger.info(f"Connected to Autodesk Inventor. Version: {version}")
                    return self.app

            except Exception as e:
                logger.warning(f"Connection attempt {attempt} failed: {e}")
                if attempt < self.retries:
                    logger.info(f"Waiting {self.delay_seconds} seconds before next attempt...")
                    time.sleep(self.delay_seconds)
                else:
                    logger.error("Exceeded maximum connection retry attempts. Autodesk Inventor could not be loaded.")
                    raise ConnectionError("Failed to connect to Autodesk Inventor COM API.") from e

    def is_connected(self) -> bool:
        """Verifies if the connection to Inventor is alive by calling a lightweight property."""
        if not self.app:
            return False
        try:
            # Touch a basic property to verify responsiveness
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

    def __enter__(self) -> Any:
        return self.connect()

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.disconnect(close_app=False)
