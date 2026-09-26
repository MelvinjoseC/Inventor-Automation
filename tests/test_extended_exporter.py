import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path
from src.core.exporter import InventorExporter


def test_exporter_step_stl_image(tmp_path):
    app = MagicMock()
    translator = MagicMock()
    translator.Activated = False
    translator.HasSaveCopyAsOptions.return_value = True
    app.ApplicationAddIns.ItemById.return_value = translator
    
    doc = MagicMock()
    exporter = InventorExporter(app)
    
    # 1. Test STEP Export
    step_out = tmp_path / "part.step"
    res_step = exporter.export_step(doc, step_out, protocol=214)
    assert res_step == step_out
    app.ApplicationAddIns.ItemById.assert_called_with("{90AF203A-3FC1-4883-85FA-630C2918072E}")
    
    # 2. Test STL Export
    stl_out = tmp_path / "part.stl"
    res_stl = exporter.export_stl(doc, stl_out, resolution="High")
    assert res_stl == stl_out
    app.ApplicationAddIns.ItemById.assert_called_with("{53324084-226B-4895-A446-24E6E342E701}")
    
    # 3. Test Image Export via ActiveView
    active_view = MagicMock()
    app.ActiveView = active_view
    img_out = tmp_path / "snapshot.png"
    res_img = exporter.export_image(doc, img_out, width=1920, height=1080)
    assert res_img == img_out
    active_view.SaveAsBitmap.assert_called_with(str(img_out.resolve()), 1920, 1080)


def test_exporter_batch(tmp_path):
    app = MagicMock()
    translator = MagicMock()
    translator.Activated = True
    translator.HasSaveCopyAsOptions.return_value = True
    app.ApplicationAddIns.ItemById.return_value = translator
    
    doc = MagicMock()
    drawing_doc = MagicMock()
    exporter = InventorExporter(app)
    
    formats = ["pdf", "dwg", "step", "stl"]
    results = exporter.export_batch(doc, "BRACKET-01", tmp_path, formats, drawing_doc=drawing_doc)
    
    assert "pdf" in results
    assert "dwg" in results
    assert "step" in results
    assert "stl" in results
    assert translator.SaveCopyAs.call_count == 4
