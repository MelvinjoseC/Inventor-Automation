# Changelog

All notable changes to the **Autodesk Inventor Automation Engine** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.2.0] - 2026-11-09

### Added
- **Advanced 3D Modeling Primitives**:
  - `create_hex_fastener`: Programmatic generation of ISO/ANSI hexagonal bolts and fasteners with polygonal head geometry and cylindrical shank extrusion.
  - `create_stepped_shaft`: Multi-diameter stepped shaft, spindle, and bushing synthesizer using revolved rotational sketch profiles (`RevolveFeatures`).
  - `create_l_bracket`: Structural angle mounting bracket generator with customizable leg lengths and uniform wall thickness.
  - `create_circular_flange`: Circular pipe and coupling flange synthesizer with Pitch Circle Diameter (PCD) bolt hole arrays and center bore passage.
  - `create_hole_feature`: Parametric engineering hole generator supporting clearance drills, counterbore pockets for socket head cap screws, and tapped threads.
- **Assembly Occurrence Patterning**:
  - `create_linear_occurrence_pattern`: 1D and 2D rectangular array patterning for component occurrences (`OccurrencePatterns.AddRectangularPattern`).
  - `create_circular_occurrence_pattern`: Polar circular array patterning around rotational axes (`OccurrencePatterns.AddCircularPattern`).
- **3D Collision & Interference Detection**:
  - `check_interference`: Automated volumetric collision analysis across assembly occurrences using Inventor's `AnalyzeInterference` API, reporting colliding components, overlap volume (cm³), and 3D centroid coordinates.
- **Intelligent Technical Drawing Enhancements**:
  - `_add_balloons`: Automated item callout balloon annotations on assembly drawings linked directly to the Parts List.
  - `_add_revision_table`: Programmatic revision tracking table placement with automated revision tag, description, and approval metadata.
  - `calculate_auto_scale`: Dynamic drawing view scale calculation (1:1, 1:2, 1:5, 2:1, etc.) fitting part bounding extents cleanly within sheet margins.
  - `_add_overall_dimensions`: Automated horizontal bounding dimension placement on base views.
- **Manufacturing Cost & Raw Stock Volume Estimator (`CostEstimator`)**:
  - Billet stock bounding envelope calculation with machining allowance margins.
  - Stock mass and chip scrap removal percentage calculation.
  - Material removal rate (MRR) based machining cycle time calculation.
  - Material cost, machining cost, and total fabrication cost estimation across mild steel, stainless steel, aluminum 6061, brass, and titanium.
- **YAML Bill of Materials Specification**:
  - Native YAML parser (`_parse_yaml`) in `BomProcessor` supporting `.yaml` and `.yml` specification files.
  - Added sample dataset: `config/sample_machined_parts.yaml`.
  - Expanded `BomItem` schema to validate hex fasteners, stepped shafts, angle brackets, and circular flanges.
- **Manufacturing Release Bundler & SHA-256 Verification**:
  - `create_manufacturing_bundle`: Versioned release ZIP packaging uniting 3D models (`.ipt`, `.iam`), 2D sheets (`.idw`), manufacturing exchange formats (`.step`, `.stl`, `.pdf`), and audit reports.
  - Automated `MANIFEST.json` generation containing cryptographic SHA-256 hashes and file size audits.
- **Enterprise CLI Expansion**:
  - Added `check-interference` subcommand for assembly collision checking.
  - Added `estimate-cost` subcommand for fabrication cost breakdown.
  - Added `bundle-release` subcommand for package archiving with verification manifest.
- **Comprehensive Test Suite**:
  - Expanded mock COM test coverage to 49 unit and integration tests across Linux and Windows runners.
- **Official MIT License**:
  - Added standard `LICENSE` file.

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
