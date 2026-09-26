import logging
import sys
from pathlib import Path

# Add root directory to path to enable importing core modules
sys.path.append(str(Path(__file__).resolve().parents[1]))

from src.automation import InventorAutomationEngine
from src.core.config_loader import ConfigLoader
from src.core.connection import InventorConnectionManager, PerformanceMode, TransactionContext
from src.core.drawing_generator import DrawingGenerator
from src.core.exporter import InventorExporter
from src.core.inspector import ModelInspector
from src.core.logger import setup_logger
from src.core.properties import PropertyManager

logger = logging.getLogger("InventorAutomation.DemoBlock")


def main() -> None:
    config = ConfigLoader("config/config.json")
    setup_logger(log_dir=str(config.get_path("paths.logs_dir")))
    logger.info("Starting parametric block generator pipeline...")

    conn = InventorConnectionManager(
        visible=config.inventor_settings.get("visible", True),
        silent_mode=config.inventor_settings.get("silent_mode", False),
    )

    try:
        with conn as app:
            with PerformanceMode(app, enable_turbo=True):
                engine = InventorAutomationEngine(app)
                drawing_gen = DrawingGenerator(app, config.drawing_settings)
                exporter = InventorExporter(app)
                inspector = ModelInspector(app, config.inspection_settings)

                # 1. Create a new part document
                part_doc = engine.new_part_document()

                with TransactionContext(app, part_doc, "CreateDemoBlock"):
                    # 2. Model 20x10x55 mm (2.0x1.0x5.5 cm) block
                    engine.create_block(
                        part_doc=part_doc,
                        width_cm=2.0,
                        height_cm=1.0,
                        thickness_cm=5.5,
                        material="Steel, Mild",
                        color="Polished",
                    )

                    # Add edge fillet
                    engine.add_fillet(part_doc, radius_cm=0.2)

                    # Set metadata
                    prop_mgr = PropertyManager(part_doc)
                    prop_mgr.set_standard_properties(
                        part_number="BM-BLOCK",
                        description="Auto-Generated Parametric Block",
                        author="DevOps Automation",
                        revision="B",
                    )

                # 3. Quality & DRC Inspection
                drc = inspector.run_drc(part_doc)
                logger.info(f"DRC Status: {'PASSED' if drc['overall_passed'] else 'FAILED'}")
                logger.info(f"Calculated Mass: {drc['properties']['mass_kg']} kg")

                # 4. Save Part File (.ipt)
                output_dir = config.get_path("paths.output_dir")
                part_out = output_dir / "Block.ipt"
                engine.save_document(part_doc, part_out)

                # 5. Generate Technical Drawing (.idw)
                drawing_doc = drawing_gen.create_drawing(
                    part_doc,
                    part_number="BM-BLOCK",
                    part_name="Auto-Generated Block Part",
                    revision="B",
                )
                drawings_dir = config.get_path("paths.drawings_dir")
                drawing_out = drawings_dir / "Block.idw"
                engine.save_document(drawing_doc, drawing_out)

                # 6. Multi-Format Exports (PDF, DWG, STEP, STL)
                exports_dir = config.get_path("paths.exports_dir")
                results = exporter.export_batch(
                    part_doc,
                    base_name="Block",
                    output_dir=exports_dir,
                    formats=["pdf", "dwg", "step", "stl"],
                    drawing_doc=drawing_doc,
                )

                # 7. Clean up
                engine.close_document(drawing_doc)
                engine.close_document(part_doc)

                logger.info("==========================================")
                logger.info("✅ SUCCESS: Parametric Part Pipeline Completed!")
                logger.info(f"   Part Model      : {part_out}")
                logger.info(f"   Drawing Sheet   : {drawing_out}")
                for fmt, path in results.items():
                    logger.info(f"   Export [{fmt.upper()}] : {path}")
                logger.info("==========================================")

    except Exception as e:
        logger.critical(f"❌ Error occurred during automation script execution: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
