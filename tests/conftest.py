"""
Pytest configuration and shared fixtures for excel_autokit tests.

Provides reusable test data generators, temporary Excel files,
and common test configurations.
"""

import os
import tempfile
from pathlib import Path
from typing import Dict

import pandas as pd
import pytest
import yaml


# ---------------------------------------------------------------------------
# Temporary directory
# ---------------------------------------------------------------------------

@pytest.fixture
def tmp_dir(tmp_path: Path) -> Path:
    """Provide a temporary directory for test outputs."""
    return tmp_path


# ---------------------------------------------------------------------------
# Sample DataFrames
# ---------------------------------------------------------------------------

@pytest.fixture
def sample_order_data_a() -> pd.DataFrame:
    """Source A order data for reconciliation tests."""
    return pd.DataFrame({
        "order_id": ["ORD001", "ORD002", "ORD003", "ORD004", "ORD005"],
        "sku": ["SKU-A1", "SKU-B2", "SKU-C3", "SKU-D4", "SKU-E5"],
        "quantity": [10, 20, 30, 40, 50],
        "amount": [100.00, 200.00, 300.00, 400.00, 500.00],
        "order_date": ["2024-01-01", "2024-01-02", "2024-01-03", "2024-01-04", "2024-01-05"],
    })


@pytest.fixture
def sample_order_data_b() -> pd.DataFrame:
    """Source B order data with deliberate differences."""
    return pd.DataFrame({
        "order_id": ["ORD001", "ORD002", "ORD003", "ORD006"],
        "sku": ["SKU-A1", "SKU-B2", "SKU-C3", "SKU-F6"],
        "quantity": [10, 25, 30, 60],           # ORD002 qty differs (25 vs 20)
        "amount": [100.00, 205.00, 300.00, 600.00],  # ORD002 amount differs
        "order_date": ["2024-01-01", "2024-01-02", "2024-01-03", "2024-01-06"],
    })


@pytest.fixture
def sample_dirty_data() -> pd.DataFrame:
    """Dirty data for cleaning tests."""
    return pd.DataFrame({
        "id": [1, 2, 3, 4, 5, 2],  # Duplicate id=2
        "name": [" Alice ", "bob", "CHARLIE", " bob", "Eve", "bob"],
        "date_col": ["2024/01/01", "2024-01-02", "01/03/2024", "2024.01.04", "2024-01-05", "2024-01-02"],
        "amount": ["100", "200.5", "N/A", "400", None, "200.5"],
        "category": ["A", "B", None, "A", "B", "B"],
    })


@pytest.fixture
def sample_merge_left() -> pd.DataFrame:
    """Left DataFrame for merge tests."""
    return pd.DataFrame({
        "key": ["K1", "K2", "K3"],
        "val_left": [1, 2, 3],
    })


@pytest.fixture
def sample_merge_right() -> pd.DataFrame:
    """Right DataFrame for merge tests."""
    return pd.DataFrame({
        "key": ["K1", "K2", "K4"],
        "val_right": [10, 20, 40],
    })


@pytest.fixture
def sample_report_data() -> pd.DataFrame:
    """Data for report generation tests."""
    return pd.DataFrame({
        "region": ["North", "North", "South", "South", "East"],
        "product": ["Widget", "Gadget", "Widget", "Gadget", "Widget"],
        "revenue": [1000.0, 2000.0, 1500.0, 2500.0, 800.0],
        "units": [50, 100, 75, 125, 40],
    })


# ---------------------------------------------------------------------------
# Temporary Excel files
# ---------------------------------------------------------------------------

@pytest.fixture
def excel_file_a(sample_order_data_a: pd.DataFrame, tmp_dir: Path) -> Path:
    """Write sample source A data to a temporary Excel file."""
    path = tmp_dir / "source_a.xlsx"
    sample_order_data_a.to_excel(str(path), index=False)
    return path


@pytest.fixture
def excel_file_b(sample_order_data_b: pd.DataFrame, tmp_dir: Path) -> Path:
    """Write sample source B data to a temporary Excel file."""
    path = tmp_dir / "source_b.xlsx"
    sample_order_data_b.to_excel(str(path), index=False)
    return path


@pytest.fixture
def excel_dirty(sample_dirty_data: pd.DataFrame, tmp_dir: Path) -> Path:
    """Write dirty data to a temporary Excel file."""
    path = tmp_dir / "dirty_data.xlsx"
    sample_dirty_data.to_excel(str(path), index=False)
    return path


# ---------------------------------------------------------------------------
# YAML configurations
# ---------------------------------------------------------------------------

@pytest.fixture
def sample_recon_config() -> Dict:
    """Sample reconciliation YAML configuration as a dictionary."""
    return {
        "reconciliation": {
            "match_keys": {
                "columns": ["order_id", "sku"],
                "ignore_case": False,
            },
            "tolerance": {
                "absolute": 0.01,
                "relative": 0.0,
            },
            "compare_columns": [
                {"name": "quantity", "tolerance": 0},
                {"name": "amount", "tolerance": 0.01},
            ],
        },
    }


@pytest.fixture
def sample_clean_config() -> Dict:
    """Sample cleaning YAML configuration as a dictionary."""
    return {
        "cleaning": {
            "pipelines": [
                {
                    "name": "standard",
                    "operations": [
                        {
                            "operation": "string_format",
                            "columns": ["name"],
                            "params": {"strip": True, "lower": True},
                        },
                        {
                            "operation": "number_format",
                            "columns": ["amount"],
                        },
                        {
                            "operation": "dedup",
                            "params": {"subset": ["id"]},
                        },
                    ],
                },
            ],
        },
    }


@pytest.fixture
def yaml_config_file(sample_recon_config: Dict, tmp_dir: Path) -> Path:
    """Write reconciliation config to a temporary YAML file."""
    path = tmp_dir / "config.yaml"
    with open(path, "w", encoding="utf-8") as fh:
        yaml.dump(sample_recon_config, fh)
    return path


@pytest.fixture
def yaml_clean_file(sample_clean_config: Dict, tmp_dir: Path) -> Path:
    """Write cleaning config to a temporary YAML file."""
    path = tmp_dir / "clean_config.yaml"
    with open(path, "w", encoding="utf-8") as fh:
        yaml.dump(sample_clean_config, fh)
    return path

# Created: 2026-09-08
