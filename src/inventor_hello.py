import os
import sys
from pathlib import Path

# Add root directory to path to enable importing core modules
sys.path.append(str(Path(__file__).resolve().parents[1]))

from src.core.connection import InventorConnectionManager

def main():
    print("Initializing environment connection...")
    
    # Instance connection manager
    conn = InventorConnectionManager(visible=True, silent_mode=False)
    
    try:
        # Connect to Inventor application
        app = conn.connect()
        print("✅ Success!")
        print("Inventor Version Display:", app.SoftwareVersion.DisplayVersion)
    except Exception as e:
        print(f"❌ Failed to attach/start Inventor instance: {e}")
        sys.exit(1)
    finally:
        conn.disconnect()

if __name__ == "__main__":
    main()
