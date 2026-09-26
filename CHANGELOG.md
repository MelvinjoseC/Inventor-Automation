# Changelog

All notable changes to the **Autodesk Inventor Automation Engine** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.0.0] - 2026-09-26

### Added
- **Multi-Primitive Modeling**:
  - `create_cylinder`: Parametric circular extrusions.
  - `create_plate_with_holes`: Mounting plates with automatic corner and center bolt patterns.
  - `create_tube`: Concentric hollow cylindrical tubes and pipes.
- **Feature Modifications**:
  - `add_fillet`: Edge-break constant radius fillet rounding.
  - `add_chamfer`: Edge-break distance beveled chamfering.
  - `create_cut_pocket`: Subtractive extrude cuts and rectangular pocketing.
- **Parametric Equations & iProperties**:
  - `ParameterManager`: Programmatic creation, evaluation, and batch updating of User and Model parameters with automatic document regeneration.
  - `PropertyManager`: Query, set, and synchronize standard and custom iProperties metadata across Part, Assembly, and Drawing documents.
- **Assembly Automation (`AssemblyAutomationEngine`)**:
  - Insertion and 3D translation matrix positioning of part components.
  - Grounding management for base components.
  - Assembly constraint automation: Mate, Flush, and Insert constraints.
  - Extraction of structured assembly BOM summaries and component counts.
- **Automated Quality Control & Design Rule Checking (`ModelInspector`)**:
  - Physical properties analysis: mass (kg), volume (cm³), surface area (cm²), center of gravity (x, y, z), and 3D bounding box dimensions.
  - Automated DRC checks: solid geometry presence, aspect ratio limits, and minimum wall thickness rules.
- **Universal Multi-Format Exporter**:
  - 3D STEP (`.stp`, `.step`) translation with AP214 and AP203 protocol selection.
  - Additive manufacturing STL (`.stl`) export with configurable resolution.
  - 2D high-resolution raster image snapshot capture (`.png`, `.jpg`).
  - Batch translation pipeline (`export_batch`).
- **Technical Drawing Enhancements**:
  - Intelligent multi-view positioning for Front, Top, and Isometric projections.
  - Automatic center mark placement for circular geometry and holes.
  - Assembly drawing generation with automatic structured Parts List (BOM Table) insertion.
- **Visual Reporting Engine (`RunReporter`)**:
  - Generation of executive HTML dashboard reports with responsive metric cards and status badges.
  - GitHub-flavored Markdown manufacturing run summaries.
- **CLI Commands & Subcommands**:
  - Added `create-assembly`, `inspect`, and `export-batch` subcommands.
  - Added `--type`, `--unit`, `--inspect`, and extended parameter flags to `generate-part` and `process-bom`.
- **DevOps CI/CD**:
  - Matrix testing strategy across Python 3.9, 3.10, 3.11, and 3.12.
  - Automated code coverage measurement (`pytest-cov`) and artifact archiving.
  - Added pre-commit hook configuration (`.pre-commit-config.yaml`).
- **Comprehensive Datasets**:
  - Added `config/sample_assembly_bom.json` and `config/sample_machined_parts.csv`.
  - Added `docs/ARCHITECTURE.md`.

### Changed
- Refactored `ConfigLoader` to support environment variable overrides (`INVENTOR_*`), schema validation, and unit normalization (`to_cm`).
- Upgraded `InventorConnectionManager` with atomic transactions (`TransactionContext`), performance mode (`PerformanceMode`), and detailed COM diagnostics (`COMErrorHelper`).
- Modernized legacy demo scripts (`inventor_hello.py`, `inventor_make_block.py`).
