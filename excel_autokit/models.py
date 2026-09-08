"""
Data models for excel_autokit.

Defines dataclass-based models for reconciliation results, cleaning rules,
and report configuration used across the toolkit.
"""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple


@dataclass
class MatchKey:
    """Represents a composite match key for reconciliation.

    Attributes:
        columns: List of column names forming the composite key.
        ignore_case: Whether string comparisons should be case-insensitive.
    """

    columns: List[str]
    ignore_case: bool = False

    def extract(self, row: Dict[str, Any]) -> Tuple:
        """Extract key values from a row dictionary."""
        values = []
        for col in self.columns:
            val = row.get(col, "")
            if isinstance(val, str) and self.ignore_case:
                val = val.strip().upper()
            values.append(val)
        return tuple(values)


@dataclass
class Tolerance:
    """Numeric tolerance configuration for fuzzy matching.

    Attributes:
        absolute: Absolute tolerance value (e.g., 0.01).
        relative: Relative tolerance as a fraction (e.g., 0.001 for 0.1%).
    """

    absolute: float = 0.0
    relative: float = 0.0

    def is_within(self, value_a: float, value_b: float) -> bool:
        """Check if two values are within configured tolerance."""
        diff = abs(value_a - value_b)
        if diff <= self.absolute:
            return True
        if self.relative > 0 and diff <= abs(value_a) * self.relative:
            return True
        return False


@dataclass
class ReconcileResult:
    """Result of a single reconciliation comparison.

    Attributes:
        status: Match status - 'matched', 'mismatched', 'missing_a', 'missing_b'.
        key_values: The composite key values used for matching.
        row_a: Data from the first source (None if missing).
        row_b: Data from the second source (None if missing).
        differences: List of column-level differences found.
    """

    status: str
    key_values: Tuple
    row_a: Optional[Dict[str, Any]] = None
    row_b: Optional[Dict[str, Any]] = None
    differences: List[Dict[str, Any]] = field(default_factory=list)

    @property
    def is_match(self) -> bool:
        """Return True if the result is a perfect match."""
        return self.status == "matched"


@dataclass
class ReconcileSummary:
    """Summary statistics for a reconciliation run.

    Attributes:
        total_a: Total records in source A.
        total_b: Total records in source B.
        matched: Number of matched pairs.
        mismatched: Number of mismatched pairs.
        missing_in_a: Records in B but not in A.
        missing_in_b: Records in A but not in B.
        total_amount_diff: Total amount difference across all matches.
        details: Detailed list of ReconcileResult objects.
    """

    total_a: int = 0
    total_b: int = 0
    matched: int = 0
    mismatched: int = 0
    missing_in_a: int = 0
    missing_in_b: int = 0
    total_amount_diff: float = 0.0
    details: List[ReconcileResult] = field(default_factory=list)


@dataclass
class CleanOperation:
    """A single data cleaning operation.

    Attributes:
        operation: Type of operation (e.g., 'date_format', 'fill_na', 'dedup').
        columns: Target columns for the operation.
        params: Additional parameters for the operation.
    """

    operation: str
    columns: Optional[List[str]] = None
    params: Dict[str, Any] = field(default_factory=dict)


@dataclass
class CleanPipeline:
    """A sequence of cleaning operations to apply.

    Attributes:
        name: Pipeline name for logging.
        operations: Ordered list of CleanOperation objects.
    """

    name: str
    operations: List[CleanOperation] = field(default_factory=list)


@dataclass
class ReportConfig:
    """Configuration for report generation.

    Attributes:
        template_name: Name of the Jinja2 template to use.
        output_path: Path for the output file.
        group_by: Column(s) to split data into separate reports.
        include_charts: Whether to embed matplotlib charts.
        include_summary: Whether to include summary rows.
    """

    template_name: str
    output_path: str = "output.xlsx"
    group_by: Optional[List[str]] = None
    include_charts: bool = False
    include_summary: bool = True


@dataclass
class DiffDetail:
    """Column-level difference detail.

    Attributes:
        column: The column name where difference was found.
        value_a: Value from source A.
        value_b: Value from source B.
        diff_type: Type of difference ('value', 'amount', 'format').
    """

    column: str
    value_a: Any
    value_b: Any
    diff_type: str = "value"

# Created: 2026-09-08
