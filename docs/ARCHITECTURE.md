# Autodesk Inventor Automation Engine: System Architecture

## 1. Overview & Architectural Goals

The **Autodesk Inventor Automation Engine** is an enterprise-grade automation framework engineered to bridge modern DevOps workflows (CI/CD runners, scheduled pipelines, REST workers) with Autodesk Inventor's legacy Windows COM API (`Inventor.Application`).

The engine addresses key challenges in programmatic CAD generation:
- **Resilience**: Safe transaction rollbacks (`TransactionContext`) and robust COM error unwrapping (`COMErrorHelper`).
- **Performance**: High-throughput batch processing via `PerformanceMode` (UI suspension and prompt suppression).
- **Flexibility**: Decoupled multi-primitive geometric modeling (`Block`, `Cylinder`, `PlateWithHoles`, `Tube`, `HexFastener`, `SteppedShaft`, `LBracket`, `CircularFlange`), unit-aware BOM ingestion (JSON, CSV, YAML), and multi-format CAD translation (`STEP`, `STL`, `DWG`, `DXF`, `PDF`, `PNG`).
- **Assembly Intelligence**: 3D constraint synthesis, occurrence patterning (linear and circular polar), and volumetric collision/interference detection (`check_interference`).
- **Quality Assurance**: Automated Design Rule Checking (DRC) for geometry presence, aspect ratios, wall thickness, hole-to-edge clearances, build envelopes, and mass budgets (`ModelInspector`).
- **Manufacturing Cost Estimation**: Volumetric raw billet sizing, chip scrap percentage, machining cycle time, and material/machining cost modeling (`CostEstimator`).
- **Release Verification**: Cryptographic SHA-256 manifest generation and package bundling (`create_manufacturing_bundle`).

---

## 2. High-Level Component Topology

```mermaid
graph TD
    subgraph Input_Layer ["Input & Configuration Layer"]
        CFG["ConfigLoader (JSON + Env Overrides)"]
        BOM["BomProcessor (JSON / CSV / YAML / MM-to-CM)"]
        CLI["Argparse CLI (Subcommand Dispatcher)"]
    end

    subgraph Core_Engine ["Automation Core Layer"]
        CONN["InventorConnectionManager (COM Lifecycle & Retry)"]
        PERF["PerformanceMode (ScreenUpdating Toggle)"]
        TRANS["TransactionContext (Atomic Rollback)"]
        ENG["InventorAutomationEngine (3D Parametric Features & Holes)"]
        ASM["AssemblyAutomationEngine (Occurrences, Patterns, Constraints, Collision)"]
        DRAW["DrawingGenerator (Views, Scales, Auto-Balloons & Parts Lists)"]
        PROP["PropertyManager & ParameterManager"]
        DRC["ModelInspector (Physical Properties & Advanced DRC)"]
        COST["CostEstimator (Billet Sizing & Machining Cost)"]
    end

    subgraph Output_Layer ["Output & Export Translation"]
        EXP["InventorExporter (STEP, STL, DWG, DXF, PDF, Image)"]
        BUNDLE["Release Bundler (ZIP + SHA-256 Manifest)"]
        REP["RunReporter (HTML Dashboard & Markdown)"]
        FS["File System (ipt, iam, idw, exports, releases, logs)"]
    end

    CLI --> CFG
    CLI --> BOM
    CLI --> CONN
    CONN --> PERF
    PERF --> TRANS
    TRANS --> ENG
    TRANS --> ASM
    ENG --> PROP
    ENG --> DRC
    ENG --> DRAW
    DRC --> COST
    ASM --> DRAW
    DRAW --> EXP
    ENG --> EXP
    EXP --> BUNDLE
    DRC --> REP
    COST --> REP
    EXP --> FS
    BUNDLE --> FS
    REP --> FS
```

---

## 3. Key Design Patterns & Engineering Decisions

### 3.1 COM Lifecycle & Connection Management
Autodesk Inventor is exposed as an out-of-process COM server (`Inventor.Application`). 
- **Attachment First, Creation Second**: `InventorConnectionManager` first invokes `win32com.client.GetObject(Class="Inventor.Application")` to attach to any running instance, preventing unnecessary process overhead.
- **Retry Mechanism**: If attachment fails, it initiates `EnsureDispatch` with exponential backoff retries.
- **Context Manager**: Implements Python's context management protocol (`__enter__` / `__exit__`), ensuring clean reference release and optional termination upon completion.

### 3.2 Atomic Transactions (`TransactionContext`)
CAD operations can fail mid-sketch due to non-manifold geometry, self-intersecting profiles, or invalid parameter inputs. The engine uses Inventor's native `TransactionManager`:
- `StartTransaction(doc, name)` creates a discrete undo state.
- If an unhandled exception occurs, `transaction.Abort()` is executed, restoring the document to its pre-operation state.
- On success, `transaction.End()` commits the feature history cleanly.

### 3.3 Unit Independence & Normalization
Autodesk Inventor's internal database standardizes all length dimensions in **centimeters (cm)** regardless of user display preferences.
- Mechanical engineering BOMs typically specify dimensions in **millimeters (mm)** or **inches (in)**.
- `ConfigLoader.to_cm()` automatically normalizes all user dimensions to internal database units, while preserving user units for drawings and notes.

### 3.4 Multi-Primitive Modeling Engine
The modeling engine (`InventorAutomationEngine`) separates 2D sketch generation, profile creation, and feature extrusions:
- **Primitives**: `create_block`, `create_cylinder`, `create_plate_with_holes`, `create_tube`, `create_hex_fastener`, `create_stepped_shaft`, `create_l_bracket`, `create_circular_flange`.
- **Hole Engineering**: `create_hole_feature` supports clearance drills, counterbore pockets for socket head screws, and tapped threads.
- **Modifications**: `add_fillet`, `add_chamfer`, `create_cut_pocket`.
- **Parametric Equations**: `ParameterManager` binds dimensions to named user parameters for dynamic parametric reconfiguration.

### 3.5 Assembly Automation & Collision Detection
The assembly engine (`AssemblyAutomationEngine`) manages occurrence placement and 3D relationships:
- **Occurrence Patterning**: Linear grid arrays (`AddRectangularPattern`) and polar circular arrays (`AddCircularPattern`).
- **Constraints**: Automated Mate, Flush, and Insert constraints.
- **Interference Detection**: `check_interference` calls Inventor's `AnalyzeInterference` API, calculating exact collision volumes and centroid locations to prevent manufacturing tool crashes.

### 3.6 Automated Design Rule Checking (DRC) & Cost Modeling
`ModelInspector` validates generated models before export:
1. **Geometry Presence**: Verifies `Volume > 0` and non-zero mass.
2. **Aspect Ratio Limit**: Evaluates maximum-to-minimum bounding box dimension ratio to identify unmanufacturable or warped profiles.
3. **Minimum Wall Thickness**: Validates thin-walled features against configured fabrication constraints.
4. **Hole-to-Edge Clearance**: Ensures bolt and pilot holes maintain safe margin distances from workpiece boundaries.
5. **Machine Build Envelope**: Confirms that overall bounding extents fit within CNC or 3D printing equipment work areas.
6. **Cost & Stock Estimator**: `CostEstimator` derives raw billet envelope dimensions, chip removal percentage, and machining cost using material removal rates (MRR).

### 3.7 Universal Export & Cryptographic Release Bundling
Inventor utilizes specialized internal translator add-ins identified by GUID/CLSID for high-fidelity vector and 3D conversions:
- **PDF**: `{0AC6DF95-25A1-11D5-8E73-0010B541CD80}`
- **DWG / DXF**: `{C24E3AC2-122E-11D5-8E91-0010B541CD80}`
- **STEP (AP214 / AP203)**: `{90AF203A-3FC1-4883-85FA-630C2918072E}`
- **STL (Additive Manufacturing)**: `{53324084-226B-4895-A446-24E6E342E701}`
- **Manufacturing Release Bundle**: Packages models, drawings, exports, and HTML reports into a versioned ZIP containing a cryptographically verified `MANIFEST.json` with SHA-256 hashes.
