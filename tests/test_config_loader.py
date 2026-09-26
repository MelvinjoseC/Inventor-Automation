import json
import pytest
from pathlib import Path
from src.core.config_loader import ConfigLoader

def test_config_loader_default_fallback(tmp_path):
    # If file doesn't exist, check that defaults are loaded
    loader = ConfigLoader(config_path=str(tmp_path / "nonexistent.json"))
    assert loader.inventor_settings["visible"] is True
    assert loader.default_part_settings["material"] == "Steel, Mild"

def test_config_loader_loading(tmp_path):
    config_file = tmp_path / "test_config.json"
    custom_data = {
        "inventor": {
            "visible": False,
            "silent_mode": True,
            "application_name": "Test.Inventor"
        },
        "default_part": {
            "width": 5.0,
            "height": 2.5
        }
    }
    with open(config_file, "w", encoding="utf-8") as f:
        json.dump(custom_data, f)

    loader = ConfigLoader(config_path=str(config_file))
    assert loader.inventor_settings["visible"] is False
    assert loader.inventor_settings["silent_mode"] is True
    assert loader.inventor_settings["application_name"] == "Test.Inventor"
    
    # Check that non-specified settings fall back to default deep merge
    assert loader.default_part_settings["thickness"] == 0.5
    assert loader.default_part_settings["width"] == 5.0

def test_get_path_resolution(tmp_path):
    config_file = tmp_path / "test_config.json"
    custom_data = {
        "paths": {
            "output_dir": "~/Desktop/TestOutput",
            "logs_dir": "custom_logs"
        }
    }
    with open(config_file, "w", encoding="utf-8") as f:
        json.dump(custom_data, f)

    loader = ConfigLoader(config_path=str(config_file))
    
    # Path with ~ should expand to absolute path containing User's home / Desktop / TestOutput
    expanded_path = loader.get_path("paths.output_dir")
    assert expanded_path.is_absolute()
    assert "Desktop" in str(expanded_path)
    
    # Relative path should resolve to absolute relative to repository root
    logs_path = loader.get_path("paths.logs_dir")
    assert logs_path.is_absolute()
    assert logs_path.name == "custom_logs"

def test_unit_conversion():
    assert ConfigLoader.to_cm(100.0, "mm") == 10.0
    assert ConfigLoader.to_cm(1.0, "in") == 2.54
    assert ConfigLoader.to_cm(2.5, "cm") == 2.5
    assert ConfigLoader.to_cm(1.0, "m") == 100.0
    with pytest.raises(ValueError):
        ConfigLoader.to_cm(10.0, "invalid_unit")

def test_environment_variable_overrides(monkeypatch, tmp_path):
    monkeypatch.setenv("INVENTOR_VISIBLE", "false")
    monkeypatch.setenv("INVENTOR_SHEET_SIZE", "A4")
    loader = ConfigLoader(config_path=str(tmp_path / "nonexistent.json"))
    assert loader.inventor_settings["visible"] is False
    assert loader.drawing_settings["sheet_size"] == "A4"
    assert loader.get("inventor.visible") is False
