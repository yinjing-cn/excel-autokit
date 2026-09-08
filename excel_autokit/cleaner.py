"""
Data cleaning engine for excel_autokit.

Provides configurable data cleaning operations including format
standardization, deduplication, missing value handling, and
multi-table merging with chainable pipeline support.
"""

import logging
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Union

import pandas as pd

from excel_autokit.models import CleanOperation, CleanPipeline
from excel_autokit.utils import (
    DataValidationError,
    ensure_path,
    get_logger,
    read_excel_safe,
    validate_columns,
)

logger = get_logger(__name__)


class DataCleaner:
    """Configurable data cleaning engine.

    Supports a chain of operations applied sequentially to a DataFrame.
    Operations include date/number format normalization, deduplication,
    missing value handling, and table merging.

    Usage::

        cleaner = DataCleaner()
        result = (
            cleaner
            .load("raw_data.xlsx")
            .normalize_dates(["order_date", "ship_date"], fmt="%Y-%m-%d")
            .normalize_numbers(["amount", "quantity"])
            .deduplicate(subset=["order_id"])
            .fill_na(strategy="zero", columns=["amount"])
            .to_dataframe()
        )
    """

    def __init__(self) -> None:
        """Initialize the DataCleaner with an empty pipeline."""
        self._df: Optional[pd.DataFrame] = None
        self._operations_log: List[str] = []

    def load(
        self,
        source: Union[str, Path, pd.DataFrame],
        sheet_name: Union[str, int] = 0,
    ) -> "DataCleaner":
        """Load data from an Excel file or DataFrame.

        Args:
            source: File path or DataFrame to load.
            sheet_name: Sheet name/index if loading from file.

        Returns:
            Self for method chaining.
        """
        if isinstance(source, pd.DataFrame):
            self._df = source.copy()
        else:
            self._df = read_excel_safe(source, sheet_name=sheet_name)
        self._operations_log.append(f"load: {len(self._df)} rows loaded")
        logger.info("DataCleaner loaded %d rows", len(self._df))
        return self

    # ------------------------------------------------------------------
    # Format normalization
    # ------------------------------------------------------------------

    def normalize_dates(
        self,
        columns: List[str],
        fmt: str = "%Y-%m-%d",
        errors: str = "coerce",
    ) -> "DataCleaner":
        """Normalize date columns to a consistent format.

        Args:
            columns: Date column names to normalize.
            fmt: Target date format string.
            errors: How to handle parsing errors ('coerce', 'raise', 'ignore').

        Returns:
            Self for method chaining.
        """
        self._ensure_loaded()
        validate_columns(self._df, columns, "normalize_dates")
        for col in columns:
            self._df[col] = pd.to_datetime(self._df[col], errors=errors)
            if errors == "coerce":
                invalid_count = self._df[col].isna().sum()
                if invalid_count > 0:
                    logger.warning("Column '%s': %d dates could not be parsed", col, invalid_count)
        self._operations_log.append(f"normalize_dates: columns={columns}, fmt={fmt}")
        return self

    def normalize_numbers(self, columns: List[str]) -> "DataCleaner":
        """Normalize numeric columns, coercing non-numeric values to NaN.

        Args:
            columns: Column names to convert to numeric.

        Returns:
            Self for method chaining.
        """
        self._ensure_loaded()
        validate_columns(self._df, columns, "normalize_numbers")
        for col in columns:
            original = self._df[col].copy()
            self._df[col] = pd.to_numeric(self._df[col], errors="coerce")
            coerced = (original.notna() & self._df[col].isna()).sum()
            if coerced > 0:
                logger.warning("Column '%s': %d values coerced to NaN", col, coerced)
        self._operations_log.append(f"normalize_numbers: columns={columns}")
        return self

    def normalize_strings(self, columns: List[str], strip: bool = True, lower: bool = False) -> "DataCleaner":
        """Normalize string columns (strip whitespace, case conversion).

        Args:
            columns: String column names to normalize.
            strip: Remove leading/trailing whitespace.
            lower: Convert to lowercase.

        Returns:
            Self for method chaining.
        """
        self._ensure_loaded()
        validate_columns(self._df, columns, "normalize_strings")
        for col in columns:
            self._df[col] = self._df[col].astype(str)
            if strip:
                self._df[col] = self._df[col].str.strip()
            if lower:
                self._df[col] = self._df[col].str.lower()
        self._operations_log.append(f"normalize_strings: columns={columns}")
        return self

    def unify_encoding(self, columns: List[str], encoding: str = "utf-8") -> "DataCleaner":
        """Unify text encoding for specified columns.

        In practice this ensures all string columns are proper Python
        unicode strings (utf-8 internally).

        Args:
            columns: Columns to unify.
            encoding: Target encoding (informational; Python 3 strings are unicode).

        Returns:
            Self for method chaining.
        """
        self._ensure_loaded()
        validate_columns(self._df, columns, "unify_encoding")
        for col in columns:
            self._df[col] = self._df[col].astype(str).apply(
                lambda x: x.encode(encoding, errors="replace").decode(encoding)
            )
        self._operations_log.append(f"unify_encoding: columns={columns}, encoding={encoding}")
        return self

    # ------------------------------------------------------------------
    # Deduplication
    # ------------------------------------------------------------------

    def deduplicate(
        self,
        subset: Optional[List[str]] = None,
        keep: str = "first",
    ) -> "DataCleaner":
        """Remove duplicate rows.

        Args:
            subset: Columns to consider for identifying duplicates.
                    If None, all columns are used.
            keep: Which duplicates to keep ('first', 'last', False).

        Returns:
            Self for method chaining.
        """
        self._ensure_loaded()
        before = len(self._df)
        if subset:
            validate_columns(self._df, subset, "deduplicate")
        self._df = self._df.drop_duplicates(subset=subset, keep=keep).reset_index(drop=True)
        removed = before - len(self._df)
        self._operations_log.append(f"deduplicate: removed={removed}, subset={subset}")
        logger.info("Deduplication removed %d rows", removed)
        return self

    def detect_duplicates(self, subset: Optional[List[str]] = None) -> pd.DataFrame:
        """Return a DataFrame of duplicate rows (without removing them).

        Args:
            subset: Columns to consider for duplicate detection.

        Returns:
            DataFrame containing only the duplicate rows.
        """
        self._ensure_loaded()
        mask = self._df.duplicated(subset=subset, keep=False)
        return self._df[mask].copy()

    # ------------------------------------------------------------------
    # Missing value handling
    # ------------------------------------------------------------------

    def fill_na(
        self,
        strategy: str = "zero",
        columns: Optional[List[str]] = None,
        value: Any = None,
    ) -> "DataCleaner":
        """Handle missing values using the specified strategy.

        Args:
            strategy: Fill strategy - 'zero', 'mean', 'median', 'ffill', 'bfill', 'value'.
            columns: Target columns. If None, applies to all columns.
            value: Fill value when strategy='value'.

        Returns:
            Self for method chaining.
        """
        self._ensure_loaded()
        target_cols = columns or list(self._df.columns)
        if columns:
            validate_columns(self._df, columns, "fill_na")

        for col in target_cols:
            if strategy == "zero":
                if self._df[col].dtype in ("float64", "int64", "float32", "int32"):
                    self._df[col] = self._df[col].fillna(0)
                else:
                    self._df[col] = self._df[col].fillna("")
            elif strategy == "mean":
                mean_val = self._df[col].mean()
                self._df[col] = self._df[col].fillna(mean_val)
            elif strategy == "median":
                median_val = self._df[col].median()
                self._df[col] = self._df[col].fillna(median_val)
            elif strategy == "ffill":
                self._df[col] = self._df[col].ffill()
            elif strategy == "bfill":
                self._df[col] = self._df[col].bfill()
            elif strategy == "value":
                self._df[col] = self._df[col].fillna(value)
            elif strategy == "drop":
                self._df = self._df.dropna(subset=[col]).reset_index(drop=True)

        self._operations_log.append(f"fill_na: strategy={strategy}, columns={target_cols}")
        return self

    def mark_missing(self, columns: List[str], indicator_col: str = "_missing_flags") -> "DataCleaner":
        """Add boolean indicator columns for missing values.

        Args:
            columns: Columns to check for missing values.
            indicator_col: Name of the column to store missing flags.

        Returns:
            Self for method chaining.
        """
        self._ensure_loaded()
        validate_columns(self._df, columns, "mark_missing")
        flags = {}
        for col in columns:
            flags[f"{col}_was_missing"] = self._df[col].isna()
        flags_df = pd.DataFrame(flags, index=self._df.index)
        self._df = pd.concat([self._df, flags_df], axis=1)
        self._operations_log.append(f"mark_missing: columns={columns}")
        return self

    # ------------------------------------------------------------------
    # Table merging
    # ------------------------------------------------------------------

    def merge(
        self,
        other: Union[str, Path, pd.DataFrame],
        on: Optional[List[str]] = None,
        how: str = "left",
        suffixes: tuple = ("_x", "_y"),
    ) -> "DataCleaner":
        """Merge with another DataFrame or Excel file.

        Args:
            other: Right-side data source.
            on: Columns to join on.
            how: Join type ('left', 'right', 'inner', 'outer').
            suffixes: Suffixes for overlapping column names.

        Returns:
            Self for method chaining.
        """
        self._ensure_loaded()
        if isinstance(other, pd.DataFrame):
            right = other
        else:
            right = read_excel_safe(other)
        self._df = self._df.merge(right, on=on, how=how, suffixes=suffixes).reset_index(drop=True)
        self._operations_log.append(f"merge: how={how}, on={on}, result_rows={len(self._df)}")
        logger.info("Merge complete: %d rows", len(self._df))
        return self

    # ------------------------------------------------------------------
    # Pipeline support
    # ------------------------------------------------------------------

    def apply_pipeline(self, pipeline: CleanPipeline) -> "DataCleaner":
        """Apply a configured cleaning pipeline.

        Args:
            pipeline: CleanPipeline with ordered operations.

        Returns:
            Self for method chaining.
        """
        dispatch: Dict[str, Callable] = {
            "date_format": lambda op: self.normalize_dates(op.columns or [], **op.params),
            "number_format": lambda op: self.normalize_numbers(op.columns or []),
            "string_format": lambda op: self.normalize_strings(
                op.columns or [],
                strip=op.params.get("strip", True),
                lower=op.params.get("lower", False),
            ),
            "dedup": lambda op: self.deduplicate(
                subset=op.columns or op.params.get("subset"),
                keep=op.params.get("keep", "first"),
            ),
            "fill_na": lambda op: self.fill_na(
                strategy=op.params.get("strategy", "zero"),
                columns=op.columns,
                value=op.params.get("value"),
            ),
            "mark_missing": lambda op: self.mark_missing(
                op.columns or [],
                indicator_col=op.params.get("indicator_col", "_missing_flags"),
            ),
        }

        for operation in pipeline.operations:
            handler = dispatch.get(operation.operation)
            if handler is None:
                logger.warning("Unknown operation '%s', skipping", operation.operation)
                continue
            handler(operation)

        return self

    # ------------------------------------------------------------------
    # Output
    # ------------------------------------------------------------------

    def to_dataframe(self) -> pd.DataFrame:
        """Return the cleaned DataFrame.

        Returns:
            Cleaned copy of the internal DataFrame.

        Raises:
            DataValidationError: If no data has been loaded.
        """
        self._ensure_loaded()
        return self._df.copy()

    def to_excel(
        self,
        output_path: Union[str, Path],
        sheet_name: str = "Cleaned",
    ) -> Path:
        """Write the cleaned data to an Excel file.

        Args:
            output_path: Output file path.
            sheet_name: Name of the output sheet.

        Returns:
            Resolved output path.
        """
        self._ensure_loaded()
        path = ensure_path(output_path)
        self._df.to_excel(str(path), sheet_name=sheet_name, index=False)
        logger.info("Cleaned data written to %s", path)
        return path

    @property
    def operations_log(self) -> List[str]:
        """Return the log of applied operations."""
        return list(self._operations_log)

    def _ensure_loaded(self) -> None:
        """Ensure data has been loaded before operations."""
        if self._df is None:
            raise DataValidationError("No data loaded. Call load() first.")

# Created: 2026-09-08
