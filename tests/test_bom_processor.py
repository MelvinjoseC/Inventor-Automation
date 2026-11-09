import json
import pytest
from pathlib import Path
from src.core.bom_processor import BomProcessor, BomItem


def test_bom_item_validation():
    defaults = {
        "width": 2.0,
        "height": 1.0,
        "thickness": 0.5,
        "material": "Steel",
        "color": "Default",
    }

    # Valid item
    data = {"part_number": "BM-001", "name": "Test", "width": 4.5, "thickness": 1.2}
    item = BomItem(data, defaults)
    assert item.part_number == "BM-001"
    assert item.width == 4.5
    assert item.height == 1.0  # fell back to default
    assert item.thickness == 1.2

    # Invalid dimensions
    with pytest.raises(ValueError):
        BomItem({"part_number": "BM-002", "width": -1.0}, defaults)
    with pytest.raises(ValueError):
        BomItem({"part_number": "BM-002", "height": "not-a-number"}, defaults)


def test_bom_processor_json(tmp_path):
    bom_file = tmp_path / "bom.json"
    bom_data = [
        {"part_number": "P1", "name": "Plate 1", "width": 10.0, "height": 5.0, "thickness": 0.5},
        {
            "part_number": "P2",
            "name": "Plate 2",
            "width": 12.0,
            "height": 6.0,
            "thickness": 0.8,
            "material": "Alum",
        },
    ]
    with open(bom_file, "w", encoding="utf-8") as f:
        json.dump(bom_data, f)

    defaults = {
        "width": 2.0,
        "height": 1.0,
        "thickness": 0.5,
        "material": "Steel",
        "color": "Default",
    }
    processor = BomProcessor(defaults)
    items = processor.parse_file(bom_file)

    assert len(items) == 2
    assert items[0].part_number == "P1"
    assert items[1].material == "Alum"


def test_bom_processor_csv(tmp_path):
    bom_file = tmp_path / "bom.csv"
    csv_content = (
        "part_number, name, width, height, thickness, material\n"
        "C1, Bracket 1, 3.0, 4.0, 0.2, Brass\n"
        "C2, Bracket 2, 3.0, 4.0, 0.2, Brass\n"
    )
    with open(bom_file, "w", encoding="utf-8") as f:
        f.write(csv_content)

    defaults = {
        "width": 2.0,
        "height": 1.0,
        "thickness": 0.5,
        "material": "Steel",
        "color": "Default",
    }
    processor = BomProcessor(defaults)
    items = processor.parse_file(bom_file)

    assert len(items) == 2
    assert items[0].part_number == "C1"
    assert items[0].name == "Bracket 1"
    assert items[1].material == "Brass"


def test_bom_primitives_and_units():
    defaults = {"unit": "mm"}

    # 1. Cylinder with mm units
    cyl_data = {"part_number": "CYL-01", "diameter": 50.0, "length": 150.0, "unit": "mm"}
    cyl_item = BomItem(cyl_data, defaults)
    assert cyl_item.part_type == "cylinder"
    assert cyl_item.diameter_cm == 5.0
    assert cyl_item.length_cm == 15.0

    # 2. Tube with inch units
    tube_data = {
        "part_number": "TB-01",
        "part_type": "tube",
        "outer_diameter": 2.0,
        "inner_diameter": 1.5,
        "length": 10.0,
        "unit": "in",
    }
    tube_item = BomItem(tube_data, defaults)
    assert tube_item.part_type == "tube"
    assert tube_item.outer_diameter_cm == 2.0 * 2.54
    assert tube_item.inner_diameter_cm == 1.5 * 2.54

    # 3. Plate with holes and fillet
    plate_data = {
        "part_number": "PL-01",
        "part_type": "plate_with_holes",
        "width": 100.0,
        "height": 80.0,
        "thickness": 10.0,
        "hole_diameter": 8.0,
        "fillet_radius": 5.0,
        "unit": "mm",
    }
    plate_item = BomItem(plate_data, defaults)
    assert plate_item.part_type == "plate_with_holes"
    assert plate_item.width_cm == 10.0
    assert plate_item.hole_diameter_cm == 0.8
    assert plate_item.fillet_radius_cm == 0.5
