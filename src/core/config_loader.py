import json
import logging
import os
from pathlib import Path
from typing import Any, Dict, Optional, Union

logger = logging.getLogger("InventorAutomation.ConfigLoader")

# Standard unit multipliers to convert to centimeters (Inventor internal standard unit)
UNIT_CONVERSION_TO_CM = {
    "mm": 0.1,
    "millimeter": 0.1,
    "millimeters": 0.1,
    "cm": 1.0,
    "centimeter": 1.0,
    "centimeters": 1.0,
    "m": 100.0,
    "meter": 100.0,
    "meters": 100.0,
    "in": 2.54,
    "inch": 2.54,
    "inches": 2.54,
    "ft": 30.48,
    "foot": 30.48,
    "feet": 30.48,
}


class ConfigLoader:
    """Manages parsing, validation, environment variable overrides, and retrieval of automation configurations."""

    DEFAULT_CONFIG: Dict[str, Any] = {
        "inventor": {
            "visible": True,
            "silent_mode": False,
            "application_name": "Inventor.Application",
            "screen_updating": True,
            "default_unit": "mm",
        },
        "paths": {
            "output_dir": "~/Desktop/InventorAutomation/output",
            "drawings_dir": "~/Desktop/InventorAutomation/drawings",
            "exports_dir": "~/Desktop/InventorAutomation/exports",
            "logs_dir": "logs",
            "reports_dir": "~/Desktop/InventorAutomation/reports",
        },
        "drawings": {
            "sheet_size": "A3",
            "title_block": "ISO",
            "standard_views": ["Front", "Top", "Isometric"],
            "view_scale": 1.0,
            "auto_dimensions": True,
            "auto_centermarks": True,
        },
        "default_part": {
            "width": 2.0,
            "height": 1.0,
            "thickness": 0.5,
            "unit": "cm",
            "material": "Steel, Mild",
            "color": "Default",
        },
        "bom": {
            "default_bom_file": "config/sample_bom.json",
            "default_unit": "mm",
        },
        "inspection": {
            "min_wall_thickness_mm": 1.0,
            "min_hole_edge_dist_mm": 2.0,
            "generate_report": True,
        },
    }

    def __init__(self, config_path: str = "config/config.json"):
        self.config_path = Path(config_path)
        self.config_data = self._load_config()
        self._apply_env_overrides()

    def _load_config(self) -> Dict[str, Any]:
        """Loads configuration from JSON file. Falls back to safe defaults on failure."""
        import copy

        merged_config = copy.deepcopy(self.DEFAULT_CONFIG)

        if not self.config_path.exists():
            logger.warning(
                f"Config file not found at {self.config_path}. Using default configuration."
            )
            return merged_config

        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                if not isinstance(loaded, dict):
                    logger.error(
                        f"Config file {self.config_path} does not contain a JSON object. Using defaults."
                    )
                    return merged_config

                # Deep merge loaded configuration with defaults
                for section, values in merged_config.items():
                    if (
                        section in loaded
                        and isinstance(loaded[section], dict)
                        and isinstance(values, dict)
                    ):
                        for k, v in loaded[section].items():
                            values[k] = v
                    elif section in loaded:
                        merged_config[section] = loaded[section]

                # Also capture any extra custom sections present in loaded
                for section, values in loaded.items():
                    if section not in merged_config:
                        merged_config[section] = values

                logger.info(f"Successfully loaded configuration from {self.config_path}")
                return merged_config
        except Exception as e:
            logger.error(
                f"Failed to parse config file: {e}. Falling back to default configuration."
            )
            return merged_config

    def _apply_env_overrides(self) -> None:
        """Applies environment variable overrides to configuration."""
        env_mappings = {
            "INVENTOR_VISIBLE": ("inventor.visible", lambda v: v.lower() in ("true", "1", "yes")),
            "INVENTOR_SILENT_MODE": (
                "inventor.silent_mode",
                lambda v: v.lower() in ("true", "1", "yes"),
            ),
            "INVENTOR_SCREEN_UPDATING": (
                "inventor.screen_updating",
                lambda v: v.lower() in ("true", "1", "yes"),
            ),
            "INVENTOR_DEFAULT_UNIT": ("inventor.default_unit", str),
            "INVENTOR_OUTPUT_DIR": ("paths.output_dir", str),
            "INVENTOR_DRAWINGS_DIR": ("paths.drawings_dir", str),
            "INVENTOR_EXPORTS_DIR": ("paths.exports_dir", str),
            "INVENTOR_LOGS_DIR": ("paths.logs_dir", str),
            "INVENTOR_REPORTS_DIR": ("paths.reports_dir", str),
            "INVENTOR_SHEET_SIZE": ("drawings.sheet_size", str),
        }

        for env_var, (config_key, transform_fn) in env_mappings.items():
            val = os.getenv(env_var)
            if val is not None:
                try:
                    transformed = transform_fn(val)
                    self._set_nested(config_key, transformed)
                    logger.info(
                        f"Applied environment override: {env_var} -> {config_key} = {transformed}"
                    )
                except Exception as e:
                    logger.warning(f"Failed to parse environment variable {env_var}='{val}': {e}")

    def _set_nested(self, key_path: str, value: Any) -> None:
        """Sets a value in the nested configuration dictionary."""
        parts = key_path.split(".")
        current = self.config_data
        for part in parts[:-1]:
            if part not in current or not isinstance(current[part], dict):
                current[part] = {}
            current = current[part]
        current[parts[-1]] = value

    def get(self, key: str, default: Any = None) -> Any:
        """Retrieves a configuration key (supports dot notation like 'inventor.visible')."""
        if "." in key:
            parts = key.split(".")
            val = self.config_data
            for p in parts:
                if isinstance(val, dict):
                    val = val.get(p)
                else:
                    return default
            return val if val is not None else default
        return self.config_data.get(key, default)

    def get_path(self, key_path: str) -> Path:
        """Retrieves a path configuration, expands user directories, and returns absolute Path."""
        parts = key_path.split(".")
        val = self.config_data
        for p in parts:
            if isinstance(val, dict):
                val = val.get(p)
            else:
                val = None
                break

        if not val:
            raise KeyError(f"Configuration path key '{key_path}' not found.")

        resolved_path = Path(str(val)).expanduser()
        # If relative, make it absolute relative to project root
        if not resolved_path.is_absolute():
            resolved_path = Path(__file__).resolve().parents[2] / resolved_path
        return resolved_path

    @staticmethod
    def to_cm(value: Union[float, int], unit: Optional[str] = "mm") -> float:
        """Converts dimension value from specified unit to centimeters (Inventor COM standard)."""
        if unit is None:
            unit = "mm"
        norm_unit = unit.strip().lower()
        multiplier = UNIT_CONVERSION_TO_CM.get(norm_unit)
        if multiplier is None:
            raise ValueError(
                f"Unsupported unit '{unit}'. Supported units: {list(UNIT_CONVERSION_TO_CM.keys())}"
            )
        return float(value) * multiplier

    @property
    def inventor_settings(self) -> Dict[str, Any]:
        return self.config_data.get("inventor", {})

    @property
    def drawing_settings(self) -> Dict[str, Any]:
        return self.config_data.get("drawings", {})

    @property
    def default_part_settings(self) -> Dict[str, Any]:
        return self.config_data.get("default_part", {})

    @property
    def bom_settings(self) -> Dict[str, Any]:
        return self.config_data.get("bom", {})

    @property
    def inspection_settings(self) -> Dict[str, Any]:
        return self.config_data.get("inspection", {})
