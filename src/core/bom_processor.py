import csv
import json
import logging
from pathlib import Path
from typing import Any, Dict, List

logger = logging.getLogger("InventorAutomation.BomProcessor")

class BomItem:
    """Represents a validated Bill of Materials (BOM) item representing a part to automate."""
    
    def __init__(self, data: Dict[str, Any], defaults: Dict[str, Any]):
        self.part_number = str(data.get("part_number", "")).strip()
        self.name = str(data.get("name", "")).strip()
        
        # Parse and validate dimensions (widths, heights, thickness are expected in cm for Inventor API)
        try:
            self.width = float(data.get("width", defaults.get("width", 2.0)))
            self.height = float(data.get("height", defaults.get("height", 1.0)))
            self.thickness = float(data.get("thickness", defaults.get("thickness", 0.5)))
        except ValueError as e:
            raise ValueError(f"Invalid dimension values for BOM part {self.part_number}: {e}")

        if self.width <= 0 or self.height <= 0 or self.thickness <= 0:
            raise ValueError(f"Dimensions must be positive values. Part: {self.part_number}")

        self.material = str(data.get("material", defaults.get("material", "Steel, Mild"))).strip()
        self.color = str(data.get("color", defaults.get("color", "Default"))).strip()

    def to_dict(self) -> Dict[str, Any]:
        return {
            "part_number": self.part_number,
            "name": self.name,
            "width": self.width,
            "height": self.height,
            "thickness": self.thickness,
            "material": self.material,
            "color": self.color
        }

    def __repr__(self) -> str:
        return f"<BomItem {self.part_number} ({self.name}): {self.width}x{self.height}x{self.thickness}>"


class BomProcessor:
    """Parses and validates BOM source files (JSON or CSV)."""

    def __init__(self, default_part_settings: Dict[str, Any]):
        self.defaults = default_part_settings

    def parse_file(self, file_path: Path) -> List[BomItem]:
        """Entrypoint to parse and validate JSON or CSV BOM list."""
        path = Path(file_path).expanduser().resolve()
        if not path.exists():
            logger.error(f"BOM input file not found: {path}")
            raise FileNotFoundError(f"BOM file not found: {path}")

        logger.info(f"Parsing BOM configuration from {path}...")
        
        suffix = path.suffix.lower()
        if suffix == ".json":
            raw_items = self._parse_json(path)
        elif suffix in [".csv", ".txt"]:
            raw_items = self._parse_csv(path)
        else:
            raise ValueError(f"Unsupported BOM file format '{suffix}'. Use .json or .csv.")

        validated_items: List[BomItem] = []
        for index, item in enumerate(raw_items):
            part_no = item.get("part_number") or f"PART-{index:03d}"
            item["part_number"] = part_no
            try:
                bom_item = BomItem(item, self.defaults)
                validated_items.append(bom_item)
            except Exception as e:
                logger.warning(f"Skipping invalid BOM item at index {index} ({part_no}): {e}")
                
        logger.info(f"Successfully validated {len(validated_items)} of {len(raw_items)} items from BOM.")
        return validated_items

    def _parse_json(self, path: Path) -> List[Dict[str, Any]]:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, list):
                return data
            elif isinstance(data, dict) and "parts" in data:
                return data["parts"]
            else:
                logger.warning("JSON format not standard. Expecting a list or object with 'parts' array.")
                return [data]

    def _parse_csv(self, path: Path) -> List[Dict[str, Any]]:
        items: List[Dict[str, Any]] = []
        with open(path, mode="r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                # Strip keys and values
                cleaned_row = {k.strip(): v.strip() for k, v in row.items() if k}
                items.append(cleaned_row)
        return items
