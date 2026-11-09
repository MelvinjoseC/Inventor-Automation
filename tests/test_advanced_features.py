import json
import zipfile
import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from src.core.assembly_engine import AssemblyAutomationEngine
from src.core.drawing_generator import DrawingGenerator
from src.core.cost_estimator import CostEstimator
from src.core.bom_processor import BomProcessor
from src.core.exporter import InventorExporter


@pytest.fixture
def mock_inventor_app():
    app = MagicMock()
    app.SoftwareVersion.DisplayVersion = "2024"
    tg = MagicMock()
    tg.CreatePoint2d.side_effect = lambda x, y: MagicMock(X=x, Y=y)
    app.TransientGeometry = tg
    app.TransientObjects.CreateObjectCollection.return_value = MagicMock()
    return app


def test_assembly_patterning(mock_inventor_app):
    assembly_doc = MagicMock()
    cd = assembly_doc.ComponentDefinition
    occ = MagicMock()

    engine = AssemblyAutomationEngine(mock_inventor_app)

    # Linear pattern
    lin_pattern = engine.create_linear_occurrence_pattern(
        assembly_doc, occ, count_x=3, spacing_x_cm=5.0
    )
    assert cd.OccurrencePatterns.AddRectangularPattern.called

    # Circular pattern
    circ_pattern = engine.create_circular_occurrence_pattern(
        assembly_doc, occ, count=6, angle_deg=360.0
    )
    assert cd.OccurrencePatterns.AddCircularPattern.called


def test_assembly_interference_detection(mock_inventor_app):
    assembly_doc = MagicMock()
    cd = assembly_doc.ComponentDefinition
    cd.Occurrences.Count = 2

    # Mock Interference results
    res1 = MagicMock()
    res1.Volume = 0.45
    res1.OccurrenceOne.Name = "PartA:1"
    res1.OccurrenceTwo.Name = "PartB:1"
    res1.Centroid.X = 10.0
    res1.Centroid.Y = 5.0
    res1.Centroid.Z = 0.0

    interference_collection = MagicMock()
    interference_collection.Count = 1
    interference_collection.Item.return_value = res1

    cd.AnalyzeInterference.return_value = interference_collection

    engine = AssemblyAutomationEngine(mock_inventor_app)
    report = engine.check_interference(assembly_doc)

    assert report["has_interference"] is True
    assert report["interference_count"] == 1
    assert report["total_interference_volume_cm3"] == 0.45
    assert report["interferences"][0]["occurrence_1"] == "PartA:1"


def test_drawing_balloons_and_revision_table(mock_inventor_app):
    sheet = MagicMock()
    sheet.Width = 42.0
    sheet.Height = 29.7
    view = MagicMock()
    view.Center.X = 20.0
    view.Center.Y = 15.0

    curve = MagicMock()
    view.DrawingCurves.Count = 3
    view.DrawingCurves.Item.return_value = curve

    gen = DrawingGenerator(mock_inventor_app)
    # Test balloon placement
    count = gen._add_balloons(sheet, view)
    assert sheet.Balloons.AutoBalloon.called or sheet.Balloons.Add.called

    # Test revision table
    rev_tbl = gen._add_revision_table(sheet, revision="B", description="Design Revision")
    assert sheet.RevisionTables.Add.called


def test_drawing_auto_scale(mock_inventor_app):
    sheet = MagicMock()
    sheet.Width = 42.0
    sheet.Height = 29.7

    part_doc = MagicMock()
    rb = MagicMock()
    rb.MinPoint.X = 0.0
    rb.MinPoint.Y = 0.0
    rb.MinPoint.Z = 0.0
    rb.MaxPoint.X = 100.0  # Large 100cm part
    rb.MaxPoint.Y = 50.0
    rb.MaxPoint.Z = 20.0
    part_doc.ComponentDefinition.RangeBox = rb

    gen = DrawingGenerator(mock_inventor_app)
    scale = gen.calculate_auto_scale(sheet, part_doc)
    assert scale <= 0.25  # Scale down large model


def test_cost_estimator():
    estimator = CostEstimator(
        machining_hourly_rate=80.0, stock_allowance_cm=0.5, setup_minutes=15.0
    )

    # 10x10x10 cm block (1000 cm3) in Aluminum 6061
    bbox = {"dx": 10.0, "dy": 10.0, "dz": 10.0}
    finished_vol = 700.0  # 300 cm3 removed as chips
    cost_info = estimator.estimate_part_cost(bbox, finished_vol, material_name="Aluminum 6061")

    assert cost_info["stock_dimensions_cm"]["dx"] == 10.5
    assert cost_info["stock_volume_cm3"] > finished_vol
    assert cost_info["scrap_volume_cm3"] > 0
    assert cost_info["material_cost_usd"] > 0
    assert cost_info["machining_cost_usd"] > 0
    assert cost_info["total_estimated_cost_usd"] == round(
        cost_info["material_cost_usd"] + cost_info["machining_cost_usd"], 2
    )


def test_yaml_bom_parsing(tmp_path):
    yaml_content = """
parts:
  - part_number: "TST-01"
    name: "Test Plate"
    part_type: "plate_with_holes"
    width: 100
    height: 80
    thickness: 10
    hole_diameter: 8
    unit: "mm"
    material: "Steel, Mild"
"""
    yaml_file = tmp_path / "test_bom.yaml"
    yaml_file.write_text(yaml_content, encoding="utf-8")

    processor = BomProcessor()
    items = processor.parse_file(yaml_file)

    assert len(items) == 1
    assert items[0].part_number == "TST-01"
    assert items[0].part_type == "plate_with_holes"
    assert items[0].width_cm == 10.0


def test_manufacturing_package_bundler(tmp_path):
    exporter = InventorExporter(inventor_app=MagicMock())

    file1 = tmp_path / "Part1.ipt"
    file1.write_text("ipt-dummy-data", encoding="utf-8")
    file2 = tmp_path / "Part1.step"
    file2.write_text("step-dummy-data", encoding="utf-8")

    out_dir = tmp_path / "releases"
    zip_path, manifest = exporter.create_manufacturing_bundle(
        bundle_name="RELEASE-V1",
        files_to_pack=[file1, file2],
        output_dir=out_dir,
        release_metadata={"version": "1.0.0"},
    )

    assert zip_path.exists()
    assert manifest["file_count"] == 2
    assert len(manifest["files"]) == 2
    assert manifest["files"][0]["sha256"] != ""

    # Verify zip contains files and MANIFEST.json
    with zipfile.ZipFile(zip_path, "r") as zf:
        names = zf.namelist()
        assert "Part1.ipt" in names
        assert "Part1.step" in names
        assert "MANIFEST.json" in names
        manifest_data = json.loads(zf.read("MANIFEST.json").decode("utf-8"))
        assert manifest_data["bundle_name"] == "RELEASE-V1"
