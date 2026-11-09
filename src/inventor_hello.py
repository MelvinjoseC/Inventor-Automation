import logging
import sys
from pathlib import Path

# Add root directory to path to enable importing core modules
sys.path.append(str(Path(__file__).resolve().parents[1]))

from src.core.connection import COMErrorHelper, InventorConnectionManager
from src.core.logger import setup_logger

logger = logging.getLogger("InventorAutomation.Hello")


def main() -> None:
    setup_logger()
    logger.info("Initializing Autodesk Inventor Session...")

    conn = InventorConnectionManager(visible=True, silent_mode=False)

    try:
        app = conn.connect()
        version = getattr(getattr(app, "SoftwareVersion", None), "DisplayVersion", "Unknown")
        user_name = getattr(getattr(app, "GeneralOptions", None), "UserName", "Local User")

        logger.info("==========================================")
        logger.info("✅ SUCCESS: Connected to Autodesk Inventor")
        logger.info(f"   Software Version : {version}")
        logger.info(f"   Inventor User    : {user_name}")
        logger.info(f"   Session Alive    : {conn.is_connected()}")
        logger.info("==========================================")
    except Exception as e:
        logger.error(f"❌ Connection failed: {COMErrorHelper.format_error(e)}")
        sys.exit(1)
    finally:
        conn.disconnect()


if __name__ == "__main__":
    main()
