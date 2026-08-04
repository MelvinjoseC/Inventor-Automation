import json
import logging
from pathlib import Path
from typing import Any, Dict

logger = logging.getLogger("InventorAutomation.ConfigLoader")

class ConfigLoader:
    """Manages parsing, validation, and retrieval of automation configurations."""
    
    def __init__(self, config_path: str = "config/config.json"):
        self.config_path = Path(config_path)
        self.config_data = self._load_config()

    def _load_config(self) -> Dict[str, Any]:
        """Loads configuration from JSON file. Falls back to safe defaults on failure."""
        default_config = {
            "inventor": {
                "visible": True,
                "silent_mode": False,
                "application_name": "Inventor.Application"
            },
            "paths": {
                "output_dir": "~/Desktop/InventorAutomation/output",
                "drawings_dir": "~/Desktop/InventorAutomation/drawings",
                "exports_dir": "~/Desktop/InventorAutomation/exports",
                "logs_dir": "logs"
            },
            "drawings": {
                "sheet_size": "A3",
                "title_block": "ISO",
                "standard_views": ["Front", "Top", "Isometric"],
                "view_scale": 1.0
            },
            "default_part": {
                "width": 2.0,
                "height": 1.0,
                "thickness": 0.5,
                "material": "Steel, Mild",
                "color": "Default"
            },
            "bom": {
                "default_bom_file": "config/sample_bom.json"
            }
        }

        if not self.config_path.exists():
            logger.warning(f"Config file not found at {self.config_path}. Using default configuration.")
            return default_config

        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                # Deep merge loaded configuration with defaults
                for section, values in default_config.items():
                    if section not in loaded:
                        loaded[section] = values
                    elif isinstance(values, dict):
                        for k, v in values.items():
                            if k not in loaded[section]:
                                loaded[section][k] = v
                logger.info(f"Successfully loaded configuration from {self.config_path}")
                return loaded
        except Exception as e:
            logger.error(f"Failed to parse config file: {e}. Falling back to default configuration.")
            return default_config

    def get(self, key: str, default: Any = None) -> Any:
        """Retrieves a configuration key."""
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
