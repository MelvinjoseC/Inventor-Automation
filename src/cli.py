import argparse
import sys
import logging
from pathlib import Path
from typing import Any, List

# Setup path to import local modules
sys.path.append(str(Path(__file__).resolve().parent))

from core.logger import setup_logger
from core.config_loader import ConfigLoader
from core.connection import InventorConnectionManager
from core.bom_processor import BomProcessor, BomItem
from core.drawing_generator import DrawingGenerator
from core.exporter import InventorExporter
from automation import InventorAutomationEngine

logger = logging.getLogger("InventorAutomation.CLI")

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="AutoCAD Inventor Automation CLI Tool for QA/QC and DevOps Pipelines.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    parser.add_argument(
        "--config", 
        type=str, 
        default="config/config.json", 
        help="Path to the JSON configuration file."
    )
    parser.add_argument(
        "--verbose", 
        action="store_true", 
        help="Enable debug-level logging."
    )

    subparsers = parser.add_subparsers(dest="command", help="Available subcommands")

    # Command: Diagnose
    subparsers.add_parser(
        "diagnose", 
        help="Perform diagnostic checks on Windows registry and Inventor COM status."
    )

    # Command: Generate Part
    part_parser = subparsers.add_parser(
        "generate-part", 
        help="Generate a single block part with optional drawings."
    )
    part_parser.add_argument("--width", type=float, help="Part width in cm")
    part_parser.add_argument("--height", type=float, help="Part height in cm")
    part_parser.add_argument("--thickness", type=float, help="Part thickness in cm")
    part_parser.add_argument("--name", type=str, default="BlockPart", help="Part description/name")
    part_parser.add_argument("--part-number", type=str, default="PART-001", help="Unique part identifier")
    part_parser.add_argument("--material", type=str, help="Material name")
    part_parser.add_argument("--color", type=str, help="Appearance / color style name")
    part_parser.add_argument("--skip-drawing", action="store_true", help="Do not generate drawing sheet")
    part_parser.add_argument(
        "--export", 
        type=str, 
        default="pdf,dwg", 
        help="Comma-separated export formats (e.g. pdf,dwg,dxf)"
    )

    # Command: Process BOM
    bom_parser = subparsers.add_parser(
        "process-bom", 
        help="Batch process a CSV/JSON Bill of Materials to generate parts/drawings."
    )
    bom_parser.add_argument(
        "--bom-file", 
        type=str, 
        help="Path to BOM file. If omitted, default in config is used."
    )
    bom_parser.add_argument(
        "--export", 
        type=str, 
        default="pdf,dwg", 
        help="Comma-separated export formats (e.g. pdf,dwg,dxf)"
    )

    return parser.parse_args()

def handle_generate_part(args: argparse.Namespace, config: ConfigLoader, app: Any) -> None:
    # Resolve parameters prioritizing CLI arguments, then config defaults
    defaults = config.default_part_settings
    width = args.width if args.width is not None else float(defaults.get("width", 2.0))
    height = args.height if args.height is not None else float(defaults.get("height", 1.0))
    thickness = args.thickness if args.thickness is not None else float(defaults.get("thickness", 0.5))
    material = args.material if args.material else defaults.get("material", "Steel, Mild")
    color = args.color if args.color else defaults.get("color", "Default")

    engine = InventorAutomationEngine(app)
    part_doc = engine.new_part_document()
    
    # 1. Generate Part Geometry
    engine.create_block(part_doc, width, height, thickness, material, color)
    
    # 2. Save Part
    output_dir = config.get_path("paths.output_dir")
    part_path = output_dir / f"{args.part_number}.ipt"
    engine.save_document(part_doc, part_path)

    # 3. Create Drawing View if not skipped
    drawing_doc = None
    if not args.skip_drawing:
        generator = DrawingGenerator(app, config.drawing_settings)
        drawing_doc = generator.create_drawing(part_doc, args.part_number, args.name)
        
        drawings_dir = config.get_path("paths.drawings_dir")
        drawing_path = drawings_dir / f"{args.part_number}.idw"
        engine.save_document(drawing_doc, drawing_path)

    # 4. Handle Exports
    if args.export:
        formats = [f.strip().lower() for f in args.export.split(",") if f.strip()]
        exporter = InventorExporter(app)
        exports_dir = config.get_path("paths.exports_dir")
        
        for fmt in formats:
            if fmt == "pdf" and drawing_doc:
                exporter.export_pdf(drawing_doc, exports_dir / f"{args.part_number}.pdf")
            elif fmt == "dwg" and drawing_doc:
                exporter.export_dwg(drawing_doc, exports_dir / f"{args.part_number}.dwg")
            elif fmt == "dxf":
                # Export dxf from part flat pattern or drawing sheet
                doc_to_export = drawing_doc if drawing_doc else part_doc
                exporter.export_dxf(doc_to_export, exports_dir / f"{args.part_number}.dxf")
            elif fmt == "pdf" and not drawing_doc:
                exporter.export_pdf(part_doc, exports_dir / f"{args.part_number}_part.pdf")

    # Clean up docs
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
    
    output_dir = config.get_path("paths.output_dir")
    drawings_dir = config.get_path("paths.drawings_dir")
    exports_dir = config.get_path("paths.exports_dir")
    
    formats = [f.strip().lower() for f in args.export.split(",") if f.strip()] if args.export else []

    logger.info(f"Processing BOM: {len(bom_items)} parts to generate...")
    
    for item in bom_items:
        logger.info(f"--- Processing {item.part_number} ({item.name}) ---")
        try:
            part_doc = engine.new_part_document()
            engine.create_block(part_doc, item.width, item.height, item.thickness, item.material, item.color)
            
            part_path = output_dir / f"{item.part_number}.ipt"
            engine.save_document(part_doc, part_path)

            # Generate drawing
            drawing_doc = generator.create_drawing(part_doc, item.part_number, item.name)
            drawing_path = drawings_dir / f"{item.part_number}.idw"
            engine.save_document(drawing_doc, drawing_path)

            # Handle exports
            for fmt in formats:
                if fmt == "pdf":
                    exporter.export_pdf(drawing_doc, exports_dir / f"{item.part_number}.pdf")
                elif fmt == "dwg":
                    exporter.export_dwg(drawing_doc, exports_dir / f"{item.part_number}.dwg")
                elif fmt == "dxf":
                    exporter.export_dxf(drawing_doc, exports_dir / f"{item.part_number}.dxf")

            # Close documents
            engine.close_document(drawing_doc)
            engine.close_document(part_doc)
            logger.info(f"Successfully processed part {item.part_number}")
        except Exception as ex:
            logger.error(f"Error processing BOM item {item.part_number}: {ex}", exc_info=True)

def main() -> None:
    args = parse_args()
    
    # Load config
    config = ConfigLoader(args.config)
    
    # Setup Logger
    log_level = logging.DEBUG if args.verbose else logging.INFO
    setup_logger(
        log_dir=str(config.get_path("paths.logs_dir")), 
        log_level=log_level
    )
    
    logger.info("Initializing Inventor Automation Command Line Interface.")

    # Route command to appropriate logic
    if args.command == "diagnose":
        logger.info("Running Environment Diagnostics...")
        try:
            import diagnose
            diagnose.run_diagnostics()
        except ImportError:
            # Fallback inline registry verification
            import winreg
            try:
                key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Inventor.Application")
                logger.info("✅ SUCCESS: Inventor.Application ProgID registered in Windows Registry.")
                winreg.CloseKey(key)
            except WindowsError:
                logger.error("❌ FAILURE: Inventor.Application is NOT registered in this system.")
                sys.exit(1)
        sys.exit(0)

    if not args.command:
        logger.error("No sub-command supplied. Use -h or --help for instructions.")
        sys.exit(1)

    # Establish COM connection and execute
    inv_settings = config.inventor_settings
    conn = InventorConnectionManager(
        visible=inv_settings.get("visible", True),
        silent_mode=inv_settings.get("silent_mode", False)
    )
    
    try:
        with conn as app:
            if args.command == "generate-part":
                handle_generate_part(args, config, app)
            elif args.command == "process-bom":
                handle_process_bom(args, config, app)
    except Exception as e:
        logger.critical(f"Execution failed due to connection or COM errors: {e}")
        sys.exit(1)
        
    logger.info("Automation task completed successfully.")

if __name__ == "__main__":
    main()
