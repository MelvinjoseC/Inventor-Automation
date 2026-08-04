import os
import sys
import platform
import logging

# Ensure setup path to import local modules
from pathlib import Path
sys.path.append(str(Path(__file__).resolve().parent))

# Set basic log configuration if executed standalone
logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger("InventorAutomation.Diagnostics")

def check_platform() -> bool:
    """Verifies that the operating system is Windows."""
    os_name = platform.system()
    logger.info(f"Operating System detected: {os_name} {platform.release()}")
    if os_name != "Windows":
        logger.error("❌ CRITICAL: Autodesk Inventor API is only available on Microsoft Windows platforms.")
        return False
    logger.info("✅ OS Platform Compatible.")
    return True

def check_win32com() -> bool:
    """Verifies pywin32 bindings are installed and accessible."""
    try:
        import win32com.client as win32
        import win32api
        logger.info(f"✅ win32com python extensions found. win32com version/path: {win32.__file__}")
        return True
    except ImportError as e:
        logger.error(
            f"❌ CRITICAL: pywin32 library is not installed. Run 'pip install pywin32' to resolve this. Detail: {e}"
        )
        return False

def check_registry() -> bool:
    """Checks the Windows Registry for Inventor COM Server registration."""
    if platform.system() != "Windows":
        logger.warning("Skipping Registry verification: OS is not Windows.")
        return False
        
    import winreg
    try:
        # Check standard HKEY_CLASSES_ROOT for ProgID registry entries
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Inventor.Application") as key:
            # Successfully opened ProgID key
            logger.info("✅ Registry: 'Inventor.Application' ProgID entry is registered.")
            
        # Try to read the CLSID associated with Inventor
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Inventor.Application\\CLSID") as clsid_key:
            clsid, _ = winreg.QueryValueEx(clsid_key, None)
            logger.info(f"✅ Registry: Inventor Application Class ID resolved to: {clsid}")
        return True
    except WindowsError:
        logger.error("❌ CRITICAL: Inventor COM registry key is missing. Is Autodesk Inventor installed?")
        return False

def check_com_responsiveness() -> bool:
    """Probes Inventor COM to see if it responds or is active."""
    try:
        import win32com.client as win32
        # Non-blocking check for active process
        try:
            app = win32.GetObject(Class="Inventor.Application")
            version = app.SoftwareVersion.DisplayVersion
            logger.info(f"✅ COM Probe: Connected to active Inventor session (Version: {version}).")
            return True
        except Exception:
            logger.warning("COM Probe: No active/running Autodesk Inventor instance detected.")
            logger.info("Note: The automation engine will try to start a new session automatically when run.")
            return True
    except Exception as e:
        logger.warning(f"COM Probe: Failed while attempting connection: {e}")
        return False

def run_diagnostics() -> bool:
    """Executes the full suite of diagnostic checks."""
    logger.info("=========================================")
    logger.info("AUTODESK INVENTOR AUTOMATION DIAGNOSTICS")
    logger.info("=========================================")
    
    logger.info(f"Python Version: {sys.version}")
    
    p = check_platform()
    w = check_win32com()
    r = check_registry()
    c = check_com_responsiveness()
    
    logger.info("-----------------------------------------")
    if p and w and r and c:
        logger.info("Result: Environment is fully configured and ready for AutoCAD Inventor Automation!")
        logger.info("=========================================")
        return True
    else:
        logger.warning("Result: Issues detected. Please verify errors/warnings reported above.")
        logger.info("=========================================")
        return False

if __name__ == "__main__":
    success = run_diagnostics()
    sys.exit(0 if success else 1)
