import csv
import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

try:
    from core.config_loader import ConfigLoader
except ImportError:
    from src.core.config_loader import ConfigLoader

logger = logging.getLogger("InventorAutomation.BomProcessor")


class BomItem:
    """Represents a validated Bill of Materials (BOM) item representing a part to automate."""

    VALID_PRIMITIVES = {"block", "cylinder", "plate_with_holes", "tube"}

    def __init__(self, data: Dict[str, Any], defaults: Dict[str, Any]):
        self.raw_data = data
        self.part_number = str(data.get("part_number", "")).strip()
        self.name = str(data.get("name", self.part_number)).strip()
        self.quantity = int(data.get("quantity", 1))

        # Determine unit
        self.unit = str(data.get("unit", defaults.get("unit", "cm"))).strip().lower()

        # Infer or validate part type
        self.part_type = self._determine_part_type(data)

        # Parse geometric dimensions according to part type
        self._parse_geometry(data, defaults)

        # Optional feature modifications
        self.fillet_radius_cm: Optional[float] = None
        if "fillet_radius" in data and data["fillet_radius"] is not None:
            self.fillet_radius_cm = ConfigLoader.to_cm(float(data["fillet_radius"]), self.unit)

        self.chamfer_distance_cm: Optional[float] = None
        if "chamfer_distance" in data and data["chamfer_distance"] is not None:
            self.chamfer_distance_cm = ConfigLoader.to_cm(float(data["chamfer_distance"]), self.unit)

        # Metadata
        self.material = str(data.get("material", defaults.get("material", "Steel, Mild"))).strip()
        self.color = str(data.get("color", defaults.get("color", "Default"))).strip()
        self.revision = str(data.get("revision", "A")).strip()
        self.vendor = str(data.get("vendor", "")).strip()

    def _determine_part_type(self, data: Dict[str, Any]) -> str:
        pt = str(data.get("part_type", "")).strip().lower()
        if pt in self.VALID_PRIMITIVES:
            return pt

        # Auto-infer
        if "inner_diameter" in data or "id" in data:
            return "tube"
        if "diameter" in data and "hole_diameter" not in data:
            return "cylinder"
        if "hole_diameter" in data or "holes" in data:
            return "plate_with_holes"
        return "block"

    def _parse_geometry(self, data: Dict[str, Any], defaults: Dict[str, Any]) -> None:
        try:
            if self.part_type == "cylinder":
                dia = float(data.get("diameter", defaults.get("diameter", 2.0)))
                length = float(data.get("length", defaults.get("length", 5.0)))
                if dia <= 0 or length <= 0:
                    raise ValueError(f"Cylinder dimensions must be positive: {dia}x{length}")
                self.diameter_cm = ConfigLoader.to_cm(dia, self.unit)
                self.length_cm = ConfigLoader.to_cm(length, self.unit)
                # Fallback attributes for generic handlers
                self.width_cm = self.diameter_cm
                self.height_cm = self.diameter_cm
                self.thickness_cm = self.length_cm

            elif self.part_type == "tube":
                od = float(data.get("outer_diameter", data.get("od", defaults.get("outer_diameter", 4.0))))
                id_val = float(data.get("inner_diameter", data.get("id", defaults.get("inner_diameter", 2.0))))
                length = float(data.get("length", defaults.get("length", 6.0)))
                if od <= 0 or id_val <= 0 or length <= 0:
                    raise ValueError("Tube dimensions must be positive")
                if id_val >= od:
                    raise ValueError(f"Inner diameter ({id_val}) must be strictly less than outer ({od})")
                self.outer_diameter_cm = ConfigLoader.to_cm(od, self.unit)
                self.inner_diameter_cm = ConfigLoader.to_cm(id_val, self.unit)
                self.length_cm = ConfigLoader.to_cm(length, self.unit)
                self.width_cm = self.outer_diameter_cm
                self.height_cm = self.outer_diameter_cm
                self.thickness_cm = self.length_cm

            elif self.part_type == "plate_with_holes":
                w = float(data.get("width", defaults.get("width", 5.0)))
                h = float(data.get("height", defaults.get("height", 4.0)))
                t = float(data.get("thickness", defaults.get("thickness", 0.5)))
                hd = float(data.get("hole_diameter", defaults.get("hole_diameter", 0.5)))
                if w <= 0 or h <= 0 or t <= 0 or hd <= 0:
                    raise ValueError("Plate and hole dimensions must be positive")
                self.width_cm = ConfigLoader.to_cm(w, self.unit)
                self.height_cm = ConfigLoader.to_cm(h, self.unit)
                self.thickness_cm = ConfigLoader.to_cm(t, self.unit)
                self.hole_diameter_cm = ConfigLoader.to_cm(hd, self.unit)
                
                margin_x = data.get("hole_margin_x")
                margin_y = data.get("hole_margin_y")
                self.hole_margin_x_cm = ConfigLoader.to_cm(float(margin_x), self.unit) if margin_x is not None else None
                self.hole_margin_y_cm = ConfigLoader.to_cm(float(margin_y), self.unit) if margin_y is not None else None
                self.corner_holes = bool(data.get("corner_holes", True))
                self.center_hole = bool(data.get("center_hole", False))

            else:  # "block"
                w = float(data.get("width", defaults.get("width", 2.0)))
                h = float(data.get("height", defaults.get("height", 1.0)))
                t = float(data.get("thickness", defaults.get("thickness", 0.5)))
                if w <= 0 or h <= 0 or t <= 0:
                    raise ValueError(f"Dimensions must be positive values. Part: {self.part_number}")
                self.width_cm = ConfigLoader.to_cm(w, self.unit)
                self.height_cm = ConfigLoader.to_cm(h, self.unit)
                self.thickness_cm = ConfigLoader.to_cm(t, self.unit)

        except (ValueError, TypeError) as e:
            raise ValueError(f"Invalid dimension values for BOM part {self.part_number}: {e}")

    # Aliases for backward compatibility
    @property
    def width(self) -> float:
        return self.width_cm

    @property
    def height(self) -> float:
        return self.height_cm

    @property
    def thickness(self) -> float:
        return self.thickness_cm

    def to_dict(self) -> Dict[str, Any]:
        return {
            "part_number": self.part_number,
            "name": self.name,
            "part_type": self.part_type,
            "unit": self.unit,
            "width_cm": self.width_cm,
            "height_cm": self.height_cm,
            "thickness_cm": self.thickness_cm,
            "material": self.material,
            "color": self.color,
            "quantity": self.quantity,
            "revision": self.revision,
            "vendor": self.vendor,
        }

    def __repr__(self) -> str:
        return f"<BomItem {self.part_number} ({self.name}) [{self.part_type}]: {self.width_cm}x{self.height_cm}x{self.thickness_cm}cm>"


class BomProcessor:
    """Parses, validates, and normalizes BOM source files (JSON, CSV, or dictionaries)."""

    def __init__(self, default_part_settings: Optional[Dict[str, Any]] = None):
        self.defaults = default_part_settings or {}

    def parse_file(self, file_path: Union[str, Path]) -> List[BomItem]:
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

        return self.parse_data(raw_items)

    def parse_data(self, raw_items: List[Dict[str, Any]]) -> List[BomItem]:
        """Validates a list of dictionary part specifications into strongly typed BomItem instances."""
        validated_items: List[BomItem] = []
        for index, item in enumerate(raw_items):
            part_no = item.get("part_number") or f"PART-{index + 1:03d}"
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
                cleaned_row = {k.strip(): v.strip() for k, v in row.items() if k and v is not None}
                if cleaned_row:
                    items.append(cleaned_row)
        return items
