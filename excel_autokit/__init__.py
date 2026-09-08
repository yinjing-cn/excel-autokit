"""
excel_autokit - Excel report automation toolkit.

A pure Python CLI + library framework for Excel report automation,
abstracted from real business scenarios (reconciliation, data cleaning,
report generation).

Author: yinjing-cn
License: MIT
"""

__version__ = "0.1.0"

from excel_autokit.reconciler import Reconciler
from excel_autokit.cleaner import DataCleaner
from excel_autokit.reporter import ReportGenerator
from excel_autokit.rules import RuleEngine

__all__ = [
    "Reconciler",
    "DataCleaner",
    "ReportGenerator",
    "RuleEngine",
    "__version__",
]

# Created: 2026-09-08
