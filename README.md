# Inventor Automation Engine

[![Inventor Automation CI](https://github.com/MelvinjoseC/Inventor-Automation/actions/workflows/ci.yml/badge.svg)](https://github.com/MelvinjoseC/Inventor-Automation/actions/workflows/ci.yml)
[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-windows-lightgrey.svg)](https://www.microsoft.com/windows)

A modular, production-ready Autodesk Inventor automation engine built in Python. Designed to run on DevOps agents, local workstations, or scheduled CI/CD runners to programmatically generate 3D parts, construct drawing sheets, and compile exports (PDF, DWG, DXF) from structured Bill of Materials (BOM) specifications.

---

## Key Features

- **Robust COM Connection Manager**: Features automatic attachment to active Inventor instances, fallback instantiation, and retry logic.
- **BOM Automation**: Decoupled CSV and JSON parser with shape parameter validators.
- **Auto-Generated Technical Drawings**: Programmatic sheet sizing (e.g. A3) layout placement of Front, Top, and Shaded Isometric drawing views.
- **Automatic CAD Translation**: Native DWG, DXF, and PDF exports driven by Inventor Translator Add-ins.
- **DevOps Integrations**: Configured with automated unit testing, mock COM validation, coding standards, and Windows-based CI/CD templates.

---

## Directory Structure

```text
Inventor-Automation/
├── .github/
│   └── workflows/
│       └── ci.yml             # GitHub Actions CI workflow
├── config/
│   ├── config.json            # Central config parameters
│   └── sample_bom.json        # Test BOM configuration
├── src/
│   ├── core/
│   │   ├── bom_processor.py   # BOM parser & validator
│   │   ├── config_loader.py   # Absolute path & config parser
│   │   ├── connection.py      # Inventor COM Session Manager
│   │   ├── drawing_generator.py# Drawing sheets & views automator
│   │   ├── exporter.py        # PDF/DWG/DXF translations
│   │   └── logger.py          # Unified logging setup
│   ├── automation.py          # Central modeling engine
│   ├── cli.py                 # Command line runner tool
│   ├── diagnose.py            # Registry & COM checks script
│   ├── inventor_hello.py      # Hello world demo script
│   └── inventor_make_block.py # Block generation demo script
├── tests/
│   ├── test_bom_processor.py  # BOM loading tests
│   ├── test_cli.py            # CLI argument routing tests
│   ├── test_com_mock.py       # Mocked COM interfaces tests
│   └── test_config_loader.py  # Path expansion & loading tests
├── pyproject.toml             # Project package & tool configuration
├── requirements.txt           # Package dependencies list
└── README.md                  # Project Documentation
```

---

## Installation & Setup

1. **System Requirements**: 
   - **OS**: Windows 10/11 or Windows Server (required for Autodesk Inventor COM Interop).
   - **Software**: Autodesk Inventor (2020 or newer).
   - **Python**: Python 3.8 to 3.12.

2. **Clone the repository**:
   ```bash
   git clone https://github.com/MelvinjoseC/Inventor-Automation.git
   cd Inventor-Automation
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Verify Environment**:
   Run the diagnostics tool to inspect system compatibility and Inventor COM registration:
   ```bash
   python src/diagnose.py
   ```

---

## CLI Usage

The project includes an extensible Command Line Interface (`src/cli.py`) for orchestrating pipelines and batch processes.

### 1. Help
```bash
python src/cli.py --help
```

### 2. Environment Diagnostics
Verify registry bindings and system configuration:
```bash
python src/cli.py diagnose
```

### 3. Single Part Generation
Create a 3D part directly from the terminal, lay out drawing views, and translate to exports:
```bash
python src/cli.py generate-part --width 3.5 --height 2.0 --thickness 0.8 --part-number BM-101 --name "Custom Bracket" --export pdf,dwg
```
*Creates `Block.ipt`, `Block.idw`, `Block.pdf`, and `Block.dwg` output files.*

### 4. Batch BOM Processing
Automatically generate geometry, annotations, and drawings for an entire bill of materials:
```bash
python src/cli.py process-bom --bom-file config/sample_bom.json --export pdf,dwg,dxf
```

---

## Configuration Settings

Configure default dimensions, export target directories, logger rules, and drawing sheet templates in `config/config.json`:

```json
{
  "inventor": {
    "visible": true,
    "silent_mode": false
  },
  "paths": {
    "output_dir": "~/Desktop/InventorAutomation/output",
    "drawings_dir": "~/Desktop/InventorAutomation/drawings",
    "exports_dir": "~/Desktop/InventorAutomation/exports"
  },
  "drawings": {
    "sheet_size": "A3",
    "view_scale": 1.0
  }
}
```

---

## Testing & Quality Control

To run automated checks and tests without requiring a local installation of Autodesk Inventor (e.g. inside virtual runners or docker containers), use:

```bash
# Run the complete test suite
pytest -v

# Run Python code style checking
black --check .

# Run lint checks
flake8 .
```

---

## DevOps Pipeline (CI/CD)

The project includes a GitHub Actions configuration under `.github/workflows/ci.yml`. It runs code quality reviews (`black`, `flake8`) and executes mock-based unit tests (`pytest`) automatically on every code push or pull request to the `main` branch.
