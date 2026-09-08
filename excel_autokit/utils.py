"""
Utility functions and shared helpers for excel_autokit.

Provides custom exceptions, logging configuration, and common helper
functions used across all modules.
"""

import logging
import sys
from pathlib import Path
from typing import Any, List, Optional, Union

import pandas as pd


# ---------------------------------------------------------------------------
# Custom Exceptions
# ---------------------------------------------------------------------------

class ExcelAutoKitError(Exception):
    """Base exception for all excel_autokit errors."""


class ConfigurationError(ExcelAutoKitError):
    """Raised when configuration (YAML rules, templates) is invalid."""


class DataValidationError(ExcelAutoKitError):
    """Raised when input data fails validation."""


class TemplateError(ExcelAutoKitError):
    """Raised when a Jinja2 template cannot be rendered."""


class ReconciliationError(ExcelAutoKitError):
    """Raised when reconciliation logic encounters an error."""


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

_DEFAULT_FORMAT = "%(asctime)s [%(name)s] %(levelname)s: %(message)s"
_DEFAULT_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


def get_logger(
    name: str,
    level: int = logging.INFO,
    fmt: Optional[str] = None,
    date_fmt: Optional[str] = None,
) -> logging.Logger:
    """Create and return a configured logger.

    Args:
        name: Logger name, typically __name__ of the calling module.
        level: Logging level (default INFO).
        fmt: Log message format string.
        date_fmt: Date format for log timestamps.

    Returns:
        Configured logging.Logger instance.
    """
    logger = logging.getLogger(name)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(
            logging.Formatter(
                fmt or _DEFAULT_FORMAT,
                datefmt=date_fmt or _DEFAULT_DATE_FORMAT,
            )
        )
        logger.addHandler(handler)
    logger.setLevel(level)
    return logger


# ---------------------------------------------------------------------------
# File helpers
# ---------------------------------------------------------------------------

def ensure_path(path: Union[str, Path], create_parent: bool = True) -> Path:
    """Resolve a path and optionally create parent directories.

    Args:
        path: File or directory path.
        create_parent: If True, create parent directories if they don't exist.

    Returns:
        Resolved pathlib.Path object.
    """
    p = Path(path).resolve()
    if create_parent:
        p.parent.mkdir(parents=True, exist_ok=True)
    return p


def validate_columns(df: pd.DataFrame, required: List[str], context: str = "") -> None:
    """Validate that a DataFrame contains all required columns.

    Args:
        df: Input DataFrame.
        required: List of required column names.
        context: Optional context string for error messages.

    Raises:
        DataValidationError: If any required column is missing.
    """
    missing = [col for col in required if col not in df.columns]
    if missing:
        ctx = f" ({context})" if context else ""
        raise DataValidationError(
            f"Missing required columns{ctx}: {missing}. "
            f"Available columns: {list(df.columns)}"
        )


def read_excel_safe(
    path: Union[str, Path],
    sheet_name: Union[str, int] = 0,
    **kwargs: Any,
) -> pd.DataFrame:
    """Read an Excel file with consistent error handling.

    Args:
        path: Path to the Excel file.
        sheet_name: Sheet to read (name or index).
        **kwargs: Additional arguments passed to pd.read_excel.

    Returns:
        DataFrame with the sheet data.

    Raises:
        DataValidationError: If the file cannot be read.
    """
    p = ensure_path(path, create_parent=False)
    if not p.exists():
        raise DataValidationError(f"File not found: {p}")
    try:
        return pd.read_excel(str(p), sheet_name=sheet_name, **kwargs)
    except Exception as exc:
        raise DataValidationError(f"Failed to read Excel file {p}: {exc}") from exc

# Created: 2026-09-08
