import pytest
from unittest.mock import MagicMock, patch, call
from src.automation import InventorAutomationEngine


@pytest.fixture
def mock_inventor_app():
    app = MagicMock()
    app.SoftwareVersion.DisplayVersion = "2024 (Build 280145000)"

    tg = MagicMock()
    tg.CreatePoint2d.side_effect = lambda x, y: (x, y)
    app.TransientGeometry = tg

    file_mgr = MagicMock()
    file_mgr.GetTemplateFile.return_value = "C:/Templates/standard.ipt"
    app.FileManager = file_mgr

    docs = MagicMock()
    doc = MagicMock()
    doc.DisplayName = "MockPart.ipt"
    docs.Add.return_value = doc
    app.Documents = docs
    return app


def test_create_hex_fastener(mock_inventor_app):
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    sketch = MagicMock()
    cd.Sketches.Add.return_value = sketch

    extrusion_feature = MagicMock()
    cd.Features.ExtrudeFeatures.Add.return_value = extrusion_feature

    engine = InventorAutomationEngine(mock_inventor_app)
    feat = engine.create_hex_fastener(
        part_doc,
        across_flats_cm=1.8,
        head_height_cm=0.8,
        shank_diameter_cm=1.0,
        shank_length_cm=4.0,
        material="Steel, High Strength",
    )

    assert feat == extrusion_feature
    assert cd.WorkPlanes.Item.called
    assert sketch.SketchLines.AddByTwoPoints.call_count == 6
    assert cd.Features.ExtrudeFeatures.Add.call_count == 2


def test_create_stepped_shaft(mock_inventor_app):
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    sketch = MagicMock()
    cd.Sketches.Add.return_value = sketch

    revolve_feature = MagicMock()
    cd.Features.RevolveFeatures.Add.return_value = revolve_feature

    engine = InventorAutomationEngine(mock_inventor_app)
    steps = [(3.0, 5.0), (2.0, 4.0), (1.5, 2.0)]
    feat = engine.create_stepped_shaft(part_doc, steps=steps, material="Stainless Steel")

    assert feat == revolve_feature
    assert cd.Features.RevolveFeatures.CreateRevolveDefinition.called
    assert cd.Features.RevolveFeatures.Add.called


def test_create_l_bracket(mock_inventor_app):
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    sketch = MagicMock()
    cd.Sketches.Add.return_value = sketch

    extrusion_feature = MagicMock()
    cd.Features.ExtrudeFeatures.Add.return_value = extrusion_feature

    engine = InventorAutomationEngine(mock_inventor_app)
    feat = engine.create_l_bracket(
        part_doc,
        leg1_length_cm=6.0,
        leg2_length_cm=6.0,
        width_cm=4.0,
        thickness_cm=0.8,
        material="Aluminum 6061",
    )

    assert feat == extrusion_feature
    assert sketch.SketchLines.AddByTwoPoints.call_count == 6
    assert cd.Features.ExtrudeFeatures.Add.called


def test_create_circular_flange(mock_inventor_app):
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    sketch = MagicMock()
    cd.Sketches.Add.return_value = sketch

    extrusion_feature = MagicMock()
    cd.Features.ExtrudeFeatures.Add.return_value = extrusion_feature

    engine = InventorAutomationEngine(mock_inventor_app)
    feat = engine.create_circular_flange(
        part_doc,
        outer_diameter_cm=12.0,
        inner_diameter_cm=4.0,
        thickness_cm=1.5,
        pcd_cm=8.5,
        bolt_count=6,
        bolt_hole_diameter_cm=0.8,
    )

    assert feat == extrusion_feature
    # 1 outer circle + 1 inner bore + 6 bolt holes = 8 circles
    assert sketch.SketchCircles.AddByCenterRadius.call_count == 8
    assert cd.Features.ExtrudeFeatures.Add.called


def test_create_hole_feature(mock_inventor_app):
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    sketch = MagicMock()
    cd.Sketches.Add.return_value = sketch

    hole_feature = MagicMock()
    cd.Features.HoleFeatures.AddDrilledThroughAll.return_value = hole_feature

    engine = InventorAutomationEngine(mock_inventor_app)
    feat = engine.create_hole_feature(
        part_doc,
        center_points=[(2.0, 2.0), (8.0, 2.0)],
        diameter_cm=0.8,
        hole_type="simple",
    )

    assert feat == hole_feature
    assert cd.Features.HoleFeatures.AddDrilledThroughAll.called


def test_primitive_validation_errors(mock_inventor_app):
    part_doc = MagicMock()
    engine = InventorAutomationEngine(mock_inventor_app)

    # L-bracket thickness >= leg
    with pytest.raises(ValueError):
        engine.create_l_bracket(part_doc, leg1_length_cm=2.0, leg2_length_cm=2.0, width_cm=1.0, thickness_cm=2.5)

    # Flange inner >= outer
    with pytest.raises(ValueError):
        engine.create_circular_flange(part_doc, outer_diameter_cm=5.0, inner_diameter_cm=6.0, thickness_cm=1.0, pcd_cm=5.5)

    # Stepped shaft empty
    with pytest.raises(ValueError):
        engine.create_stepped_shaft(part_doc, steps=[])
