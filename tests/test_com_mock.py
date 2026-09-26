import pytest
from unittest.mock import MagicMock, patch, call
from pathlib import Path
from src.automation import InventorAutomationEngine
from src.core.drawing_generator import DrawingGenerator
from src.core.exporter import InventorExporter
from src.core.connection import InventorConnectionManager

@pytest.fixture
def mock_inventor_app():
    # Setup nested mock structures mirroring the Autodesk Inventor COM object model
    app = MagicMock()
    
    # Mock SoftwareVersion
    app.SoftwareVersion.DisplayVersion = "2024 (Build 280145000)"
    
    # Mock TransientObjects & TransientGeometry
    app.TransientObjects.CreateTranslationContext.return_value = MagicMock()
    app.TransientObjects.CreateNameValueMap.return_value = MagicMock()
    app.TransientObjects.CreateDataMedium.return_value = MagicMock()
    
    tg = MagicMock()
    tg.CreatePoint2d.side_effect = lambda x, y: (x, y)
    app.TransientGeometry = tg
    
    # Mock FileManager
    file_mgr = MagicMock()
    file_mgr.GetTemplateFile.return_value = "C:/Templates/standard.ipt"
    app.FileManager = file_mgr
    
    # Mock Documents collection
    docs = MagicMock()
    doc = MagicMock()
    doc.DisplayName = "MockPart.ipt"
    docs.Add.return_value = doc
    app.Documents = docs
    
    # Mock AddIns
    translator = MagicMock()
    translator.Activated = False
    app.ApplicationAddIns.ItemById.return_value = translator
    
    return app

def test_connection_manager_success(mock_inventor_app):
    # Patch win32com and gencache to return mock app
    with patch("win32com.client.GetObject", return_value=mock_inventor_app):
        conn = InventorConnectionManager(retries=1)
        app = conn.connect()
        assert app == mock_inventor_app
        assert conn.is_connected() is True

def test_transaction_context_commit_and_abort(mock_inventor_app):
    from src.core.connection import TransactionContext, PerformanceMode, COMErrorHelper
    
    doc = MagicMock()
    mock_trans = MagicMock()
    mock_inventor_app.TransactionManager.StartTransaction.return_value = mock_trans
    
    # Test commit on success
    with TransactionContext(mock_inventor_app, doc, "TestOp"):
        pass
    mock_trans.End.assert_called_once()
    mock_trans.Abort.assert_not_called()
    
    # Test abort on exception
    mock_trans.reset_mock()
    with pytest.raises(RuntimeError):
        with TransactionContext(mock_inventor_app, doc, "FailingOp"):
            raise RuntimeError("Failure inside transaction")
    mock_trans.Abort.assert_called_once()
    mock_trans.End.assert_not_called()

def test_performance_mode(mock_inventor_app):
    from src.core.connection import PerformanceMode
    mock_inventor_app.ScreenUpdating = True
    mock_inventor_app.SilentOnSelectSet = False
    
    with PerformanceMode(mock_inventor_app, enable_turbo=True):
        assert mock_inventor_app.ScreenUpdating is False
        assert mock_inventor_app.SilentOnSelectSet is True
        
    assert mock_inventor_app.ScreenUpdating is True
    assert mock_inventor_app.SilentOnSelectSet is False

def test_com_error_helper():
    from src.core.connection import COMErrorHelper
    # Test standard exception
    err = ValueError("Standard error")
    assert "Standard error" in COMErrorHelper.format_error(err)
    
    # Test COM-like tuple exception
    class MockCOMException(Exception):
        def __init__(self):
            super().__init__(-2147467259, "Unspecified error", (0, "Inventor.Server", "Geometry self-intersects", None, 0, -2147467259), None)
            
    com_err = MockCOMException()
    formatted = COMErrorHelper.format_error(com_err)
    assert "Geometry self-intersects" in formatted
    assert "Inventor.Server" in formatted

def test_automation_engine_part_creation(mock_inventor_app):
    # Patch CastTo to bypass win32com casting mechanics during mock testing
    with patch("src.automation.CastTo", lambda doc, interface: doc):
        engine = InventorAutomationEngine(mock_inventor_app)
        part = engine.new_part_document()
        
        # Verify template query and Add document methods are invoked
        mock_inventor_app.FileManager.GetTemplateFile.assert_called_with(12290)
        mock_inventor_app.Documents.Add.assert_called_with(12290, "C:/Templates/standard.ipt", True)
        assert part.DisplayName == "MockPart.ipt"

def test_automation_engine_block_geometry(mock_inventor_app):
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    
    # Mock Sketches, Lines, Profile, Extrusion interfaces
    sketch = MagicMock()
    cd.Sketches.Add.return_value = sketch
    
    profile = MagicMock()
    sketch.Profiles.AddForSolid.return_value = profile
    
    extrusion_feature = MagicMock()
    cd.Features.ExtrudeFeatures.Add.return_value = extrusion_feature
    
    engine = InventorAutomationEngine(mock_inventor_app)
    
    # Generate 3D block
    ext = engine.create_block(part_doc, 2.0, 1.0, 0.5)
    
    # Verify sketch added on XY plane (index 3)
    cd.WorkPlanes.Item.assert_called_with(3)
    cd.Sketches.Add.assert_called_once()
    
    # Verify rectangle creation
    mock_inventor_app.TransientGeometry.CreatePoint2d.assert_has_calls([
        call(0, 0), call(2.0, 1.0)
    ])
    sketch.SketchLines.AddAsTwoPointRectangle.assert_called_once()
    
    # Verify Extrusion creation
    cd.Features.ExtrudeFeatures.CreateExtrudeDefinition.assert_called_once_with(profile, 20481)
    cd.Features.ExtrudeFeatures.Add.assert_called_once()
    assert ext == extrusion_feature

def test_drawing_generator(mock_inventor_app):
    part_doc = MagicMock()
    drawing_doc = MagicMock()
    sheet = MagicMock()
    sheet.Width = 42.0
    sheet.Height = 29.7
    drawing_doc.Sheets.Item.return_value = sheet
    mock_inventor_app.Documents.Add.return_value = drawing_doc
    
    settings = {"sheet_size": "A3", "view_scale": 1.5}
    
    with patch("src.core.drawing_generator.CastTo", lambda doc, interface: doc):
        generator = DrawingGenerator(mock_inventor_app, settings)
        
        # Invoke drawing creation
        doc = generator.create_drawing(part_doc, "BM-100", "Bracket Plate")
        
        # Verify sheet size setup
        assert sheet.Size == 9993  # kA3SheetSize
        
        # Verify Front view and Isometric view added
        sheet.DrawingViews.AddBaseView.assert_has_calls([
            # Front view: scale 1.5, Front orientation (10764), HiddenLine style (32258)
            call(part_doc, (42.0 * 0.3, 29.7 * 0.4), 1.5, 10764, 32258),
            # Isometric view: scale 1.5, IsoTopRight orientation (10774), Shaded style (32259)
            call(part_doc, (42.0 * 0.75, 29.7 * 0.7), 1.5, 10774, 32259)
        ])
        # Verify projected view (Top view) added
        sheet.DrawingViews.AddProjectedView.assert_called_once()

def test_exporter_pdf_dwg(mock_inventor_app, tmp_path):
    doc = MagicMock()
    exporter = InventorExporter(mock_inventor_app)
    
    # Create temp target export paths
    pdf_out = tmp_path / "part.pdf"
    dwg_out = tmp_path / "part.dwg"
    
    translator = mock_inventor_app.ApplicationAddIns.ItemById.return_value
    translator.HasSaveCopyAsOptions.return_value = True
    
    with patch("src.core.exporter.CastTo", lambda addin, interface: addin):
        # Test PDF export
        exporter.export_pdf(doc, pdf_out)
        mock_inventor_app.ApplicationAddIns.ItemById.assert_any_call("{0AC6DF95-25A1-11D5-8E73-0010B541CD80}")
        
        # Test DWG export
        exporter.export_dwg(doc, dwg_out)
        mock_inventor_app.ApplicationAddIns.ItemById.assert_any_call("{C24E3AC2-122E-11D5-8E91-0010B541CD80}")
        
        # Verify translator saved copies
        assert translator.SaveCopyAs.call_count == 2

def test_cylinder_primitive(mock_inventor_app):
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    sketch = MagicMock()
    cd.Sketches.Add.return_value = sketch
    profile = MagicMock()
    sketch.Profiles.AddForSolid.return_value = profile
    ext_feature = MagicMock()
    cd.Features.ExtrudeFeatures.Add.return_value = ext_feature
    
    engine = InventorAutomationEngine(mock_inventor_app)
    res = engine.create_cylinder(part_doc, diameter_cm=4.0, length_cm=10.0)
    
    sketch.SketchCircles.AddByCenterRadius.assert_called_once()
    cd.Features.ExtrudeFeatures.CreateExtrudeDefinition.assert_called_once_with(profile, 20481)
    assert res == ext_feature

def test_plate_with_holes_primitive(mock_inventor_app):
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    sketch = MagicMock()
    cd.Sketches.Add.return_value = sketch
    profile = MagicMock()
    sketch.Profiles.AddForSolid.return_value = profile
    ext_feature = MagicMock()
    cd.Features.ExtrudeFeatures.Add.return_value = ext_feature
    
    engine = InventorAutomationEngine(mock_inventor_app)
    res = engine.create_plate_with_holes(
        part_doc, width_cm=10.0, height_cm=8.0, thickness_cm=1.0,
        hole_diameter_cm=1.0, corner_holes=True, center_hole=True
    )
    
    sketch.SketchLines.AddAsTwoPointRectangle.assert_called_once()
    # 4 corner holes + 1 center hole = 5 circles
    assert sketch.SketchCircles.AddByCenterRadius.call_count == 5
    assert res == ext_feature

def test_tube_primitive(mock_inventor_app):
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    sketch = MagicMock()
    cd.Sketches.Add.return_value = sketch
    profile = MagicMock()
    sketch.Profiles.AddForSolid.return_value = profile
    ext_feature = MagicMock()
    cd.Features.ExtrudeFeatures.Add.return_value = ext_feature
    
    engine = InventorAutomationEngine(mock_inventor_app)
    res = engine.create_tube(part_doc, outer_diameter_cm=6.0, inner_diameter_cm=4.0, length_cm=12.0)
    
    # 2 concentric circles
    assert sketch.SketchCircles.AddByCenterRadius.call_count == 2
    assert res == ext_feature
    
    # Validation error if ID >= OD
    with pytest.raises(ValueError):
        engine.create_tube(part_doc, outer_diameter_cm=4.0, inner_diameter_cm=6.0, length_cm=10.0)

def test_feature_modifications_fillet_chamfer_pocket(mock_inventor_app):
    part_doc = MagicMock()
    cd = part_doc.ComponentDefinition
    
    # Mock Edge collection & bodies
    body = MagicMock()
    edge1 = MagicMock()
    body.Edges.Count = 4
    body.Edges.Item.return_value = edge1
    cd.SurfaceBodies.Item.return_value = body
    
    edge_col = MagicMock()
    mock_inventor_app.TransientObjects.CreateEdgeCollection.return_value = edge_col
    
    # Mock Fillets
    fillet_feature = MagicMock()
    cd.Features.FilletFeatures.AddSimple.return_value = fillet_feature
    
    # Mock Chamfers
    chamfer_feature = MagicMock()
    cd.Features.ChamferFeatures.AddUsingDistance.return_value = chamfer_feature
    
    # Mock Cut Extrude
    sketch = MagicMock()
    cd.Sketches.Add.return_value = sketch
    profile = MagicMock()
    sketch.Profiles.AddForSolid.return_value = profile
    pocket_feature = MagicMock()
    cd.Features.ExtrudeFeatures.Add.return_value = pocket_feature
    
    engine = InventorAutomationEngine(mock_inventor_app)
    
    # Test Fillet
    f = engine.add_fillet(part_doc, radius_cm=0.2)
    assert f == fillet_feature
    
    # Test Chamfer
    c = engine.add_chamfer(part_doc, distance_cm=0.1)
    assert c == chamfer_feature
    
    # Test Cut Pocket
    p = engine.create_cut_pocket(part_doc, 1.0, 1.0, 2.0, 2.0, 0.5)
    assert p == pocket_feature
    cd.Features.ExtrudeFeatures.CreateExtrudeDefinition.assert_called_with(profile, 20482)


