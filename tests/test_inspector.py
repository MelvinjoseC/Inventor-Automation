import pytest
from unittest.mock import MagicMock
from src.core.inspector import ModelInspector
from src.core.bom_processor import BomItem


def test_inspector_physical_properties():
    app = MagicMock()
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    
    # Mock MassProperties
    mass_props = MagicMock()
    mass_props.Mass = 1.25
    mass_props.Volume = 159.2
    mass_props.Area = 240.5
    cg = MagicMock()
    cg.X = 1.5
    cg.Y = 2.0
    cg.Z = 0.5
    mass_props.CenterOfMass = cg
    cd.MassProperties = mass_props
    
    # Mock RangeBox
    rb = MagicMock()
    rb.MinPoint.X = 0.0
    rb.MinPoint.Y = 0.0
    rb.MinPoint.Z = 0.0
    rb.MaxPoint.X = 10.0
    rb.MaxPoint.Y = 5.0
    rb.MaxPoint.Z = 2.0
    cd.RangeBox = rb
    
    inspector = ModelInspector(app)
    props = inspector.get_physical_properties(part_doc)
    
    assert props["mass_kg"] == 1.25
    assert props["volume_cm3"] == 159.2
    assert props["bounding_box_cm"]["dx"] == 10.0
    assert props["bounding_box_cm"]["dy"] == 5.0
    assert props["center_of_gravity"]["x_cm"] == 1.5


def test_inspector_drc():
    app = MagicMock()
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    
    # Valid model properties
    mass_props = MagicMock()
    mass_props.Mass = 0.5
    mass_props.Volume = 50.0
    mass_props.Area = 100.0
    cd.MassProperties = mass_props
    
    rb = MagicMock()
    rb.MinPoint.X = 0.0
    rb.MinPoint.Y = 0.0
    rb.MinPoint.Z = 0.0
    rb.MaxPoint.X = 5.0
    rb.MaxPoint.Y = 2.0
    rb.MaxPoint.Z = 1.0
    cd.RangeBox = rb
    
    bom_item = BomItem({
        "part_number": "P-100",
        "width": 5.0,
        "height": 2.0,
        "thickness": 1.0,  # 1.0 cm = 10.0 mm
    }, defaults={})
    
    inspector = ModelInspector(app, inspection_rules={"min_wall_thickness_mm": 1.0, "max_aspect_ratio": 20.0})
    report = inspector.run_drc(part_doc, bom_item)
    
    assert report["overall_passed"] is True
    assert report["passed_count"] >= 3
    
    # Test failure case: wall thickness too thin
    thin_item = BomItem({
        "part_number": "P-THIN",
        "thickness": 0.05,  # 0.05 cm = 0.5 mm (< 1.0 mm limit)
    }, defaults={})
    
    fail_report = inspector.run_drc(part_doc, thin_item)
    assert fail_report["overall_passed"] is False


def test_inspector_envelope_and_mass_budget():
    app = MagicMock()
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition

    mass_props = MagicMock()
    mass_props.Mass = 12.0  # Over budget
    mass_props.Volume = 50.0
    mass_props.Area = 100.0
    cd.MassProperties = mass_props

    rb = MagicMock()
    rb.MinPoint.X = 0.0
    rb.MinPoint.Y = 0.0
    rb.MinPoint.Z = 0.0
    rb.MaxPoint.X = 150.0  # Exceeds 100cm envelope
    rb.MaxPoint.Y = 2.0
    rb.MaxPoint.Z = 1.0
    cd.RangeBox = rb

    rules = {
        "max_envelope_cm": [100.0, 100.0, 100.0],
        "max_mass_kg": 10.0,
    }
    inspector = ModelInspector(app, inspection_rules=rules)
    report = inspector.run_drc(part_doc)

    assert report["overall_passed"] is False
    failed_rules = [c["rule"] for c in report["checks"] if not c["passed"]]
    assert "BuildEnvelopeLimit" in failed_rules
    assert "MassBudgetLimit" in failed_rules
