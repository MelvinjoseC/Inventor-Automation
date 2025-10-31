import win32com.client as win32

PROGID = "Inventor.Application"

def get_inventor():
    try:
        inv = win32.GetObject(Class=PROGID)
        print("✅ Attached to running Inventor")
    except Exception:
        inv = win32.Dispatch(PROGID)
        inv.Visible = True
        print("🚀 Started a new Inventor session")
    return inv

def main():
    inv = get_inventor()
    inv.Visible = True
    print("Inventor Version:", inv.SoftwareVersion.DisplayVersion)

if __name__ == "__main__":
    main()
