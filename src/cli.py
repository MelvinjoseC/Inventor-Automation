import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Any, List, Optional

# Setup path to import local modules
sys.path.append(str(Path(__file__).resolve().parent))

from automation import InventorAutomationEngine
from core.assembly_engine import AssemblyAutomationEngine
from core.bom_processor import BomItem, BomProcessor
from core.config_loader import ConfigLoader
from core.cost_estimator import CostEstimator
from core.connection import InventorConnectionManager, PerformanceMode, TransactionContext
from core.drawing_generator import DrawingGenerator
from core.exporter import InventorExporter
from core.inspector import ModelInspector
from core.logger import setup_logger
from core.properties import PropertyManager

logger = logging.getLogger("InventorAutomation.CLI")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="AutoCAD Inventor Automation Enterprise CLI for Modeling, Assemblies, Drawings, and CI/CD.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    parser.add_argument(
        "--config",
        type=str,
        default="config/config.json",
        help="Path to the JSON configuration file.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable debug-level logging.",
    )

    subparsers = parser.add_subparsers(dest="command", help="Available subcommands")

    # Command: Diagnose
    subparsers.add_parser(
        "diagnose",
        help="Perform diagnostic checks on Windows registry, COM interop, and hardware status.",
    )

    # Command: Generate Part
    part_parser = subparsers.add_parser(
        "generate-part",
        help="Generate a parametric 3D part (block, cylinder, plate_with_holes, tube) with drawings and exports.",
    )
    part_parser.add_argument("--type", type=str, default="block", choices=["block", "cylinder", "plate_with_holes", "tube", "hex_fastener", "l_bracket", "circular_flange"], help="Part primitive geometry type")
    part_parser.add_argument("--unit", type=str, default="cm", choices=["mm", "cm", "in"], help="Dimensional unit")
    part_parser.add_argument("--width", type=float, help="Part width")
    part_parser.add_argument("--height", type=float, help="Part height")
    part_parser.add_argument("--thickness", type=float, help="Part thickness")
    part_parser.add_argument("--diameter", type=float, help="Cylinder diameter")
    part_parser.add_argument("--length", type=float, help="Cylinder or tube length")
    part_parser.add_argument("--outer-diameter", type=float, help="Tube outer diameter")
    part_parser.add_argument("--inner-diameter", type=float, help="Tube inner diameter")
    part_parser.add_argument("--hole-diameter", type=float, help="Plate hole diameter")
    part_parser.add_argument("--fillet-radius", type=float, help="Fillet edge radius")
    part_parser.add_argument("--chamfer-distance", type=float, help="Chamfer edge distance")
    part_parser.add_argument("--name", type=str, default="Part", help="Part description/name")
    part_parser.add_argument("--part-number", type=str, default="PART-001", help="Unique part identifier")
    part_parser.add_argument("--revision", type=str, default="A", help="Engineering revision tag")
    part_parser.add_argument("--material", type=str, help="Material name")
    part_parser.add_argument("--color", type=str, help="Appearance / color style name")
    part_parser.add_argument("--skip-drawing", action="store_true", help="Do not generate drawing sheet")
    part_parser.add_argument("--inspect", action="store_true", help="Perform DRC and physical properties inspection")
    part_parser.add_argument(
        "--export",
        type=str,
        default="pdf,dwg",
        help="Comma-separated export formats (e.g. pdf,dwg,dxf,step,stl,png)",
    )

    # Command: Process BOM
    bom_parser = subparsers.add_parser(
        "process-bom",
        help="Batch process a CSV/JSON Bill of Materials to generate parts, drawings, and exports.",
    )
    bom_parser.add_argument(
        "--bom-file",
        type=str,
        help="Path to BOM file. If omitted, default in config is used.",
    )
    bom_parser.add_argument(
        "--export",
        type=str,
        default="pdf,dwg",
        help="Comma-separated export formats (e.g. pdf,dwg,dxf,step,stl,png)",
    )
    bom_parser.add_argument(
        "--inspect",
        action="store_true",
        help="Run DRC quality check across all items",
    )

    # Command: Create Assembly
    asm_parser = subparsers.add_parser(
        "create-assembly",
        help="Assemble components into an IAM assembly document and generate drawing with parts list.",
    )
    asm_parser.add_argument("--assembly-number", type=str, default="ASM-001", help="Assembly part number")
    asm_parser.add_argument("--name", type=str, default="Main Assembly", help="Assembly name")
    asm_parser.add_argument("--components", type=str, required=True, help="Comma-separated list of component .ipt paths")
    asm_parser.add_argument("--export", type=str, default="pdf,dwg,step", help="Export formats")

    # Command: Inspect
    inspect_parser = subparsers.add_parser(
        "inspect",
        help="Inspect physical properties (mass, volume, CoG) and run DRC rules on an existing part.",
    )
    inspect_parser.add_argument("--file", type=str, required=True, help="Path to .ipt part file to inspect")

    # Command: Export Batch
    export_parser = subparsers.add_parser(
        "export-batch",
        help="Batch export an existing Inventor document (.ipt, .iam, .idw) into STEP, STL, DWG, PDF, or PNG.",
    )
    export_parser.add_argument("--file", type=str, required=True, help="Path to Inventor source file")
    export_parser.add_argument("--formats", type=str, default="step,stl,pdf,dwg", help="Target formats")

    # Command: Check Interference
    interf_parser = subparsers.add_parser(
        "check-interference",
        help="Run 3D collision and interference analysis on an assembly document (.iam).",
    )
    interf_parser.add_argument("--assembly", type=str, required=True, help="Path to .iam assembly file")

    # Command: Estimate Cost
    cost_parser = subparsers.add_parser(
        "estimate-cost",
        help="Estimate raw stock billet requirements, scrap chip volume, machining time, and fabrication cost.",
    )
    cost_parser.add_argument("--file", type=str, required=True, help="Path to .ipt part file")
    cost_parser.add_argument("--material", type=str, default="Steel, Mild", help="Material name")
    cost_parser.add_argument("--hourly-rate", type=float, default=75.0, help="Machine shop rate ($/hr)")

    # Command: Bundle Release
    bundle_parser = subparsers.add_parser(
        "bundle-release",
        help="Assemble CAD deliverables and reports into a versioned release ZIP with SHA-256 verification manifest.",
    )
    bundle_parser.add_argument("--name", type=str, required=True, help="Release package base name")
    bundle_parser.add_argument("--files", type=str, required=True, help="Comma-separated list of deliverable file paths")
    bundle_parser.add_argument("--output-dir", type=str, default="releases", help="Target directory for release bundle")

    return parser.parse_args()


def handle_generate_part(args: argparse.Namespace, config: ConfigLoader, app: Any) -> None:
    defaults = config.default_part_settings
    unit = args.unit or defaults.get("unit", "cm")

    engine = InventorAutomationEngine(app)
    part_doc = engine.new_part_document()

    # Update metadata
    prop_mgr = PropertyManager(part_doc)
    prop_mgr.set_standard_properties(
        part_number=args.part_number,
        description=args.name,
        author="DevOps Automation",
        revision=args.revision,
    )

    material = args.material or defaults.get("material", "Steel, Mild")
    color = args.color or defaults.get("color", "Default")

    # Generate Geometry according to type
    ptype = args.type.lower()
    if ptype == "cylinder":
        dia = ConfigLoader.to_cm(args.diameter if args.diameter else 2.0, unit)
        length = ConfigLoader.to_cm(args.length if args.length else 5.0, unit)
        engine.create_cylinder(part_doc, dia, length, material, color)
    elif ptype == "plate_with_holes":
        w = ConfigLoader.to_cm(args.width if args.width else 10.0, unit)
        h = ConfigLoader.to_cm(args.height if args.height else 8.0, unit)
        t = ConfigLoader.to_cm(args.thickness if args.thickness else 1.0, unit)
        hd = ConfigLoader.to_cm(args.hole_diameter if args.hole_diameter else 1.0, unit)
        engine.create_plate_with_holes(part_doc, w, h, t, hd, material=material, color=color)
    elif ptype == "tube":
        od = ConfigLoader.to_cm(args.outer_diameter if args.outer_diameter else 4.0, unit)
        id_val = ConfigLoader.to_cm(args.inner_diameter if args.inner_diameter else 2.0, unit)
        length = ConfigLoader.to_cm(args.length if args.length else 6.0, unit)
        engine.create_tube(part_doc, od, id_val, length, material, color)
    else:  # block
        w = ConfigLoader.to_cm(args.width if args.width is not None else float(defaults.get("width", 2.0)), unit)
        h = ConfigLoader.to_cm(args.height if args.height is not None else float(defaults.get("height", 1.0)), unit)
        t = ConfigLoader.to_cm(args.thickness if args.thickness is not None else float(defaults.get("thickness", 0.5)), unit)
        engine.create_block(part_doc, w, h, t, material, color)

    # Optional feature modifications
    if args.fillet_radius:
        engine.add_fillet(part_doc, ConfigLoader.to_cm(args.fillet_radius, unit))
    if args.chamfer_distance:
        engine.add_chamfer(part_doc, ConfigLoader.to_cm(args.chamfer_distance, unit))

    # Optional Inspection
    if args.inspect:
        inspector = ModelInspector(app, config.inspection_settings)
        drc_result = inspector.run_drc(part_doc)
        print("\n--- Model Inspection & DRC Summary ---")
        print(f"Overall Status: {'PASSED' if drc_result['overall_passed'] else 'FAILED'}")
        print(f"Mass: {drc_result['properties']['mass_kg']} kg | Volume: {drc_result['properties']['volume_cm3']} cm³")
        for c in drc_result["checks"]:
            status_icon = "✅" if c["passed"] else "❌"
            print(f"  {status_icon} {c['rule']}: {c['details']}")

    # Save Part
    output_dir = config.get_path("paths.output_dir")
    part_path = output_dir / f"{args.part_number}.ipt"
    engine.save_document(part_doc, part_path)

    # Drawing
    drawing_doc = None
    if not args.skip_drawing:
        generator = DrawingGenerator(app, config.drawing_settings)
        drawing_doc = generator.create_drawing(part_doc, args.part_number, args.name, revision=args.revision)
        drawings_dir = config.get_path("paths.drawings_dir")
        drawing_path = drawings_dir / f"{args.part_number}.idw"
        engine.save_document(drawing_doc, drawing_path)

    # Exports
    if args.export:
        formats = [f.strip().lower() for f in args.export.split(",") if f.strip()]
        exporter = InventorExporter(app)
        exports_dir = config.get_path("paths.exports_dir")
        exporter.export_batch(part_doc, args.part_number, exports_dir, formats, drawing_doc=drawing_doc)

    # Cleanup
    if drawing_doc:
        engine.close_document(drawing_doc)
    engine.close_document(part_doc)


def handle_process_bom(args: argparse.Namespace, config: ConfigLoader, app: Any) -> None:
    bom_file = args.bom_file if args.bom_file else config.bom_settings.get("default_bom_file", "config/sample_bom.json")
    bom_path = Path(bom_file)

    processor = BomProcessor(config.default_part_settings)
    bom_items = processor.parse_file(bom_path)

    engine = InventorAutomationEngine(app)
    generator = DrawingGenerator(app, config.drawing_settings)
    exporter = InventorExporter(app)
    inspector = ModelInspector(app, config.inspection_settings) if args.inspect else None

    output_dir = config.get_path("paths.output_dir")
    drawings_dir = config.get_path("paths.drawings_dir")
    exports_dir = config.get_path("paths.exports_dir")
    formats = [f.strip().lower() for f in args.export.split(",") if f.strip()] if args.export else []

    logger.info(f"Processing BOM: {len(bom_items)} parts to generate...")

    with PerformanceMode(app, enable_turbo=True):
        for item in bom_items:
            logger.info(f"--- Processing {item.part_number} ({item.name}) [{item.part_type}] ---")
            try:
                part_doc = engine.new_part_document()

                with TransactionContext(app, part_doc, f"Build_{item.part_number}"):
                    # Dispatch to specific primitive method
                    if item.part_type == "cylinder":
                        engine.create_cylinder(part_doc, item.diameter_cm, item.length_cm, item.material, item.color)
                    elif item.part_type == "tube":
                        engine.create_tube(part_doc, item.outer_diameter_cm, item.inner_diameter_cm, item.length_cm, item.material, item.color)
                    elif item.part_type == "plate_with_holes":
                        engine.create_plate_with_holes(
                            part_doc, item.width_cm, item.height_cm, item.thickness_cm,
                            item.hole_diameter_cm, item.hole_margin_x_cm, item.hole_margin_y_cm,
                            item.corner_holes, item.center_hole, item.material, item.color,
                        )
                    elif item.part_type == "hex_fastener":
                        engine.create_hex_fastener(
                            part_doc, item.across_flats_cm, item.head_height_cm,
                            item.shank_diameter_cm, item.shank_length_cm, item.material, item.color,
                        )
                    elif item.part_type == "l_bracket":
                        engine.create_l_bracket(
                            part_doc, item.leg1_length_cm, item.leg2_length_cm,
                            item.width_cm, item.thickness_cm, item.material, item.color,
                        )
                    elif item.part_type == "circular_flange":
                        engine.create_circular_flange(
                            part_doc, item.outer_diameter_cm, item.inner_diameter_cm,
                            item.thickness_cm, item.pcd_cm, item.bolt_count,
                            item.bolt_hole_diameter_cm, item.material, item.color,
                        )
                    elif item.part_type == "stepped_shaft":
                        engine.create_stepped_shaft(part_doc, item.steps, item.material, item.color)
                    else:
                        engine.create_block(part_doc, item.width_cm, item.height_cm, item.thickness_cm, item.material, item.color)

                    # Feature modifications
                    if item.fillet_radius_cm:
                        engine.add_fillet(part_doc, item.fillet_radius_cm)
                    if item.chamfer_distance_cm:
                        engine.add_chamfer(part_doc, item.chamfer_distance_cm)

                # Optional DRC
                if inspector:
                    inspector.run_drc(part_doc, item)

                part_path = output_dir / f"{item.part_number}.ipt"
                engine.save_document(part_doc, part_path)

                drawing_doc = generator.create_drawing(part_doc, item.part_number, item.name, revision=item.revision)
                drawing_path = drawings_dir / f"{item.part_number}.idw"
                engine.save_document(drawing_doc, drawing_path)

                if formats:
                    exporter.export_batch(part_doc, item.part_number, exports_dir, formats, drawing_doc=drawing_doc)

                engine.close_document(drawing_doc)
                engine.close_document(part_doc)
                logger.info(f"Successfully processed part {item.part_number}")
            except Exception as ex:
                logger.error(f"Error processing BOM item {item.part_number}: {ex}", exc_info=True)


def handle_create_assembly(args: argparse.Namespace, config: ConfigLoader, app: Any) -> None:
    comp_paths = [Path(p.strip()) for p in args.components.split(",") if p.strip()]
    asm_engine = AssemblyAutomationEngine(app)
    drawing_gen = DrawingGenerator(app, config.drawing_settings)
    exporter = InventorExporter(app)

    asm_doc = asm_engine.new_assembly_document()
    spacing = 10.0  # cm spacing between placed components

    for idx, comp_path in enumerate(comp_paths):
        is_base = (idx == 0)
        asm_engine.place_component(asm_doc, comp_path, x_cm=idx * spacing, grounded=is_base)

    output_dir = config.get_path("paths.output_dir")
    asm_path = output_dir / f"{args.assembly_number}.iam"
    output_dir.mkdir(parents=True, exist_ok=True)
    asm_doc.SaveAs(str(asm_path), True)
    logger.info(f"Saved Assembly to: {asm_path}")

    # Generate assembly drawing with parts list
    drawing_doc = drawing_gen.create_assembly_drawing(
        asm_doc,
        assembly_number=args.assembly_number,
        assembly_name=args.name,
        include_parts_list=True,
    )
    drawings_dir = config.get_path("paths.drawings_dir")
    drawing_path = drawings_dir / f"{args.assembly_number}.idw"
    drawing_doc.SaveAs(str(drawing_path), True)
    logger.info(f"Saved Assembly Drawing to: {drawing_path}")

    if args.export:
        formats = [f.strip().lower() for f in args.export.split(",") if f.strip()]
        exports_dir = config.get_path("paths.exports_dir")
        exporter.export_batch(asm_doc, args.assembly_number, exports_dir, formats, drawing_doc=drawing_doc)

    drawing_doc.Close(True)
    asm_doc.Close(True)
    logger.info(f"Assembly {args.assembly_number} creation complete.")


def handle_inspect(args: argparse.Namespace, config: ConfigLoader, app: Any) -> None:
    file_path = Path(args.file).expanduser().resolve()
    if not file_path.exists():
        logger.error(f"Part file not found: {file_path}")
        sys.exit(1)

    doc = app.Documents.Open(str(file_path), False)
    try:
        inspector = ModelInspector(app, config.inspection_settings)
        drc = inspector.run_drc(doc)
        print("\n==========================================")
        print(f"INSPECTION REPORT: {file_path.name}")
        print("==========================================")
        print(f"Overall Status: {'PASSED' if drc['overall_passed'] else 'FAILED'}")
        p = drc["properties"]
        print(f"Mass: {p['mass_kg']} kg")
        print(f"Volume: {p['volume_cm3']} cm³")
        print(f"Surface Area: {p['surface_area_cm2']} cm²")
        print(f"Bounding Box: {p['bounding_box_cm']['dx']} x {p['bounding_box_cm']['dy']} x {p['bounding_box_cm']['dz']} cm")
        print(f"Center of Mass: ({p['center_of_gravity']['x_cm']}, {p['center_of_gravity']['y_cm']}, {p['center_of_gravity']['z_cm']}) cm")
        print("\nDesign Rule Checks:")
        for c in drc["checks"]:
            icon = "✅ PASS" if c["passed"] else "❌ FAIL"
            print(f"  [{icon}] {c['rule']}: {c['details']}")
        print("==========================================\n")
    finally:
        doc.Close(True)


def handle_export_batch(args: argparse.Namespace, config: ConfigLoader, app: Any) -> None:
    file_path = Path(args.file).expanduser().resolve()
    if not file_path.exists():
        logger.error(f"File not found: {file_path}")
        sys.exit(1)

    formats = [f.strip().lower() for f in args.formats.split(",") if f.strip()]
    doc = app.Documents.Open(str(file_path), False)
    try:
        exporter = InventorExporter(app)
        exports_dir = config.get_path("paths.exports_dir")
        results = exporter.export_batch(doc, file_path.stem, exports_dir, formats)
        print(f"\nGenerated {len(results)} exports in {exports_dir}:")
        for fmt, path in results.items():
            print(f"  - {fmt.upper()}: {path}")
    finally:
        doc.Close(True)


def handle_check_interference(args: argparse.Namespace, config: ConfigLoader, app: Any) -> None:
    path = Path(args.assembly).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Assembly file not found: {path}")

    doc = app.Documents.Open(str(path))
    try:
        engine = AssemblyAutomationEngine(app)
        report = engine.check_interference(doc)

        print("\n=== 3D Assembly Interference Report ===")
        print(f"File: {path.name}")
        print(f"Collision Detected: {'YES (COLLISION WARNING)' if report['has_interference'] else 'NO (CLEAR)'}")
        print(f"Total Interferences: {report['interference_count']}")
        print(f"Total Interference Volume: {report['total_interference_volume_cm3']} cm³")

        for idx, inter in enumerate(report["interferences"], 1):
            print(f"  [{idx}] {inter['occurrence_1']} <--> {inter['occurrence_2']}: {inter['volume_cm3']} cm³ at {inter['centroid']}")
    finally:
        doc.Close(True)


def handle_estimate_cost(args: argparse.Namespace, config: ConfigLoader, app: Any) -> None:
    path = Path(args.file).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Part file not found: {path}")

    doc = app.Documents.Open(str(path))
    try:
        inspector = ModelInspector(app, config.inspection_settings)
        props = inspector.get_physical_properties(doc)

        estimator = CostEstimator(machining_hourly_rate=args.hourly_rate)
        cost_report = estimator.estimate_part_cost(
            bounding_box_cm=props["bounding_box_cm"],
            finished_volume_cm3=props["volume_cm3"],
            material_name=args.material,
        )

        print("\n=== Manufacturing Cost & Stock Estimation ===")
        print(f"Part: {path.name} | Material: {cost_report['material']}")
        dims = cost_report['stock_dimensions_cm']
        print(f"Required Stock Billet: {dims['dx']} x {dims['dy']} x {dims['dz']} cm ({cost_report['stock_volume_cm3']} cm³, {cost_report['stock_mass_kg']} kg)")
        print(f"Finished Part Mass: {cost_report['finished_mass_kg']} kg | Scrap Volume: {cost_report['scrap_volume_cm3']} cm³ ({cost_report['material_removal_percentage']}%)")
        print(f"Estimated Machining Time: {cost_report['machining_time_minutes']} min")
        print(f"Raw Material Cost: ${cost_report['material_cost_usd']:.2f}")
        print(f"Machining Cost: ${cost_report['machining_cost_usd']:.2f}")
        print(f"Total Estimated Fabrication Cost: ${cost_report['total_estimated_cost_usd']:.2f}")
    finally:
        doc.Close(True)


def handle_bundle_release(args: argparse.Namespace) -> None:
    file_paths = [Path(f.strip()) for f in args.files.split(",") if f.strip()]
    exporter = InventorExporter(app=None)
    out_dir = Path(args.output_dir)
    zip_path, manifest = exporter.create_manufacturing_bundle(
        bundle_name=args.name,
        files_to_pack=file_paths,
        output_dir=out_dir,
    )
    print(f"\n=== Manufacturing Release Bundle Created ===")
    print(f"Archive: {zip_path}")
    print(f"Verified Files: {manifest['file_count']}")
    for f in manifest["files"]:
        print(f"  - {f['name']} ({f['size_bytes']} bytes) SHA-256: {f['sha256'][:16]}...")


def main() -> None:
    args = parse_args()
    config = ConfigLoader(args.config)

    log_level = logging.DEBUG if args.verbose else logging.INFO
    setup_logger(log_dir=str(config.get_path("paths.logs_dir")), log_level=log_level)

    logger.info("Initializing Inventor Automation Command Line Interface.")

    if args.command == "diagnose":
        logger.info("Running Environment Diagnostics...")
        try:
            import diagnose
            diagnose.run_diagnostics()
        except ImportError:
            import winreg
            try:
                key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Inventor.Application")
                logger.info("✅ SUCCESS: Inventor.Application ProgID registered in Windows Registry.")
                winreg.CloseKey(key)
            except WindowsError:
                logger.error("❌ FAILURE: Inventor.Application is NOT registered in this system.")
                sys.exit(1)
        sys.exit(0)

    if args.command == "bundle-release":
        handle_bundle_release(args)
        sys.exit(0)

    if not args.command:
        logger.error("No sub-command supplied. Use -h or --help for instructions.")
        sys.exit(1)

    inv_settings = config.inventor_settings
    conn = InventorConnectionManager(
        visible=inv_settings.get("visible", True),
        silent_mode=inv_settings.get("silent_mode", False),
    )

    try:
        with conn as app:
            if args.command == "generate-part":
                handle_generate_part(args, config, app)
            elif args.command == "process-bom":
                handle_process_bom(args, config, app)
            elif args.command == "create-assembly":
                handle_create_assembly(args, config, app)
            elif args.command == "inspect":
                handle_inspect(args, config, app)
            elif args.command == "export-batch":
                handle_export_batch(args, config, app)
            elif args.command == "check-interference":
                handle_check_interference(args, config, app)
            elif args.command == "estimate-cost":
                handle_estimate_cost(args, config, app)
    except Exception as e:
        logger.critical(f"Execution failed due to connection or COM errors: {e}", exc_info=True)
        sys.exit(1)

    logger.info("Automation task completed successfully.")


if __name__ == "__main__":
    main()
