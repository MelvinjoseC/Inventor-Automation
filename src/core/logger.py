import logging
import os
from pathlib import Path
from typing import Optional


def setup_logger(
    log_dir: str = "logs", log_level: int = logging.INFO, log_file_name: str = "automation.log"
) -> logging.Logger:
    """Configures and returns the central logger for Inventor Automation."""
    logger = logging.getLogger("InventorAutomation")

    # If handlers already configured, do not add them again
    if logger.handlers:
        return logger

    logger.setLevel(log_level)

    # Create console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)-7s] [%(name)s] - %(message)s", datefmt="%H:%M:%S"
    )
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    # Create file handler
    try:
        resolved_log_dir = Path(log_dir).expanduser()
        if not resolved_log_dir.is_absolute():
            # Relative to project root (2 levels up from core)
            resolved_log_dir = Path(__file__).resolve().parents[2] / resolved_log_dir

        resolved_log_dir.mkdir(parents=True, exist_ok=True)
        log_file_path = resolved_log_dir / log_file_name

        file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
        file_handler.setLevel(log_level)
        file_formatter = logging.Formatter(
            "%(asctime)s - %(levelname)s - [%(name)s:%(filename)s:%(lineno)d] - %(message)s"
        )
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)

        logger.info(f"Logging initialized. Outputting to console and file: {log_file_path}")
    except Exception as e:
        logger.error(f"Failed to initialize file logging: {e}. Console logging will continue.")

    return logger


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """Retrieves a named logger parented under the root InventorAutomation namespace."""
    if name:
        return logging.getLogger(f"InventorAutomation.{name}")
    return logging.getLogger("InventorAutomation")
