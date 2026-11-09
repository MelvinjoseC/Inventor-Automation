import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path
from src.core.assembly_engine import AssemblyAutomationEngine


def test_assembly_engine_creation_and_placement(tmp_path):
    # Setup mock component file
    comp_file = tmp_path / "Bracket.ipt"
    comp_file.write_text("dummy ipt content")

    app = MagicMock()
    app.FileManager.GetTemplateFile.return_value = "C:/Templates/standard.iam"

    assembly_doc = MagicMock()
    assembly_doc.DisplayName = "TestAssembly.iam"
    app.Documents.Add.return_value = assembly_doc

    cd = assembly_doc.ComponentDefinition
    occurrences = cd.Occurrences
    mock_occ = MagicMock()
    mock_occ.Grounded = False
    occurrences.Add.return_value = mock_occ

    engine = AssemblyAutomationEngine(app)

    # 1. Create assembly document
    with patch("src.core.assembly_engine.CastTo", lambda doc, iface: doc):
        doc = engine.new_assembly_document()
        assert doc == assembly_doc
        app.Documents.Add.assert_called_with(12291, "C:/Templates/standard.iam", True)

    # 2. Place component
    occ = engine.place_component(
        assembly_doc, comp_file, x_cm=5.0, y_cm=2.0, z_cm=0.0, grounded=True
    )
    assert occ == mock_occ
    assert mock_occ.Grounded is True
    occurrences.Add.assert_called_once()


def test_assembly_constraints():
    app = MagicMock()
    assembly_doc = MagicMock()
    cd = assembly_doc.ComponentDefinition
    constraints = cd.Constraints

    e1 = MagicMock()
    e2 = MagicMock()

    engine = AssemblyAutomationEngine(app)

    # Mate constraint
    engine.add_mate_constraint(assembly_doc, e1, e2, offset_cm=1.5)
    constraints.AddMateConstraint.assert_called_with(e1, e2, 1.5)

    # Flush constraint
    engine.add_flush_constraint(assembly_doc, e1, e2, offset_cm=0.0)
    constraints.AddFlushConstraint.assert_called_with(e1, e2, 0.0)

    # Insert constraint
    engine.add_insert_constraint(assembly_doc, e1, e2, axes_opposed=True, distance_cm=0.5)
    constraints.AddInsertConstraint.assert_called_with(e1, e2, True, 0.5)


def test_assembly_bom_summary():
    app = MagicMock()
    assembly_doc = MagicMock()
    cd = assembly_doc.ComponentDefinition
    occurrences = cd.Occurrences
    occurrences.Count = 3

    # Setup mock occurrences (2 of Part A, 1 of Part B)
    occ1 = MagicMock()
    occ1.Name = "Bracket:1"
    occ1.Definition.Document.PropertySets.Item().Item().Value = "PART-A"

    occ2 = MagicMock()
    occ2.Name = "Bracket:2"
    occ2.Definition.Document.PropertySets.Item().Item().Value = "PART-A"

    occ3 = MagicMock()
    occ3.Name = "Pin:1"
    occ3.Definition.Document.PropertySets.Item().Item().Value = "PART-B"

    occurrences.Item.side_effect = lambda idx: [occ1, occ2, occ3][idx - 1]

    engine = AssemblyAutomationEngine(app)
    summary = engine.get_assembly_bom_summary(assembly_doc)

    assert len(summary) == 2
    part_a = next(item for item in summary if item["part_number"] == "PART-A")
    part_b = next(item for item in summary if item["part_number"] == "PART-B")
    assert part_a["quantity"] == 2
    assert part_b["quantity"] == 1


def test_assembly_drawing_generation():
    from src.core.drawing_generator import DrawingGenerator

    app = MagicMock()
    assembly_doc = MagicMock()
    drawing_doc = MagicMock()
    sheet = MagicMock()
    sheet.Width = 42.0
    sheet.Height = 29.7
    drawing_doc.Sheets.Item.return_value = sheet
    app.Documents.Add.return_value = drawing_doc

    gen = DrawingGenerator(app, {"sheet_size": "A3", "view_scale": 1.0})
    with patch("src.core.drawing_generator.CastTo", lambda doc, iface: doc):
        doc = gen.create_assembly_drawing(
            assembly_doc,
            assembly_number="ASM-100",
            assembly_name="Pivoting Clamp Assembly",
            include_parts_list=True,
        )
        assert doc == drawing_doc
        sheet.DrawingViews.AddBaseView.assert_called()
        sheet.PartsLists.Add.assert_called_once()
