import json
import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from src.automation import InventorAutomationEngine
from src.core.assembly_engine import AssemblyAutomationEngine
from src.core.bom_processor import BomProcessor
from src.core.config_loader import ConfigLoader
from src.core.connection import PerformanceMode, TransactionContext
from src.core.drawing_generator import DrawingGenerator
from src.core.exporter import InventorExporter
from src.core.inspector import ModelInspector
from src.core.parameters import ParameterManager
from src.core.properties import PropertyManager
from src.core.reporter import RunReporter


def test_full_e2e_automation_pipeline(tmp_path):
    """End-to-End test validating the complete CAD engineering lifecycle."""
    # 1. Config Loader
    config = ConfigLoader("config/config.json")
    assert config.inventor_settings["application_name"] == "Inventor.Application"

    # 2. Mock COM App Setup
    app = MagicMock()
    app.FileManager.GetTemplateFile.return_value = "C:/Templates/standard.ipt"
    app.ScreenUpdating = True
    app.SilentOnSelectSet = False

    # Mock Transient Objects & Geometry
    tg = MagicMock()
    tg.CreatePoint2d.side_effect = lambda x, y: (x, y)
    app.TransientGeometry = tg
    app.TransientObjects.CreateTranslationContext.return_value = MagicMock()
    app.TransientObjects.CreateNameValueMap.return_value = MagicMock()
    app.TransientObjects.CreateDataMedium.return_value = MagicMock()
    app.TransientObjects.CreateEdgeCollection.return_value = MagicMock()

    # Mock Document
    doc = MagicMock()
    doc.DisplayName = "TestPart.ipt"
    app.Documents.Add.return_value = doc

    # Mock Addins for Exporter
    translator = MagicMock()
    translator.Activated = True
    translator.HasSaveCopyAsOptions.return_value = True
    app.ApplicationAddIns.ItemById.return_value = translator

    # 3. Create Multi-Primitive Sample BOM
    sample_bom_data = [
        {
            "part_number": "ENG-BRACKET",
            "name": "Engine Mounting Bracket",
            "part_type": "plate_with_holes",
            "width": 120.0,
            "height": 90.0,
            "thickness": 12.0,
            "hole_diameter": 10.0,
            "fillet_radius": 4.0,
            "unit": "mm",
            "material": "Steel, Mild",
        },
        {
            "part_number": "MAIN-SHAFT",
            "name": "Transmission Shaft",
            "part_type": "cylinder",
            "diameter": 40.0,
            "length": 250.0,
            "unit": "mm",
            "material": "Steel, Alloy",
        },
    ]

    bom_file = tmp_path / "e2e_bom.json"
    bom_file.write_text(json.dumps(sample_bom_data), encoding="utf-8")

    # 4. Parse BOM
    processor = BomProcessor(config.default_part_settings)
    bom_items = processor.parse_file(bom_file)
    assert len(bom_items) == 2

    # 5. Initialize Engines
    engine = InventorAutomationEngine(app)
    drawing_gen = DrawingGenerator(app, config.drawing_settings)
    exporter = InventorExporter(app)
    inspector = ModelInspector(app, config.inspection_settings)
    reporter = RunReporter("E2E Automated Pipeline Verification")

    # 6. Execute Batch Pipeline with Performance Mode & Transactions
    with PerformanceMode(app, enable_turbo=True):
        for item in bom_items:
            part_doc = MagicMock()
            cd = part_doc.ComponentDefinition
            sketch = MagicMock()
            cd.Sketches.Add.return_value = sketch
            profile = MagicMock()
            sketch.Profiles.AddForSolid.return_value = profile
            ext_feature = MagicMock()
            cd.Features.ExtrudeFeatures.Add.return_value = ext_feature

            # Mock Mass Properties for Inspector
            mass_props = MagicMock()
            mass_props.Mass = 2.4
            mass_props.Volume = 310.0
            mass_props.Area = 450.0
            cd.MassProperties = mass_props
            rb = MagicMock()
            rb.MinPoint.X, rb.MinPoint.Y, rb.MinPoint.Z = 0.0, 0.0, 0.0
            rb.MaxPoint.X, rb.MaxPoint.Y, rb.MaxPoint.Z = (
                item.width_cm,
                item.height_cm,
                item.thickness_cm,
            )
            cd.RangeBox = rb

            # Model feature
            if item.part_type == "cylinder":
                engine.create_cylinder(part_doc, item.diameter_cm, item.length_cm, item.material)
            elif item.part_type == "plate_with_holes":
                engine.create_plate_with_holes(
                    part_doc,
                    item.width_cm,
                    item.height_cm,
                    item.thickness_cm,
                    item.hole_diameter_cm,
                    material=item.material,
                )

            # Quality Check
            drc = inspector.run_drc(part_doc, item)
            assert drc["overall_passed"] is True

            # Export batch
            export_formats = ["pdf", "step", "stl"]
            exports = exporter.export_batch(
                part_doc, item.part_number, tmp_path / "exports", export_formats
            )
            assert len(exports) == 3

            # Record in Reporter
            reporter.add_result(
                part_number=item.part_number,
                name=item.name,
                part_type=item.part_type,
                mass_kg=drc["properties"]["mass_kg"],
                volume_cm3=drc["properties"]["volume_cm3"],
                drc_passed=drc["overall_passed"],
                exports=list(exports.keys()),
            )

    # 7. Generate Reports
    md_report = reporter.generate_markdown(tmp_path / "report.md")
    html_report = reporter.generate_html(tmp_path / "report.html")

    assert reporter.total_count == 2
    assert reporter.passed_count == 2
    assert (tmp_path / "report.md").exists()
    assert (tmp_path / "report.html").exists()
