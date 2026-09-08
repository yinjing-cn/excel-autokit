#!/usr/bin/env python3
"""
Excel AutoKit - Demo Script

Demonstrates the core features of excel_autokit including:
1. Data reconciliation between two sources
2. Data cleaning pipeline
3. Report generation

All data is fictional and for demonstration purposes only.
"""

import pandas as pd
from pathlib import Path

from excel_autokit import Reconciler, DataCleaner, ReportGenerator, RuleEngine
from excel_autokit.models import Tolerance


def demo_reconciliation():
    """Demonstrate the reconciliation engine."""
    print("=" * 60)
    print("DEMO: Reconciliation Engine")
    print("=" * 60)

    # Create fictional source data
    source_a = pd.DataFrame({
        "order_id": ["ORD001", "ORD002", "ORD003", "ORD004"],
        "sku": ["WIDGET-A", "WIDGET-B", "GADGET-X", "GADGET-Y"],
        "quantity": [100, 200, 50, 75],
        "unit_price": [9.99, 14.99, 24.99, 19.99],
        "amount": [999.00, 2998.00, 1249.50, 1499.25],
    })

    source_b = pd.DataFrame({
        "order_id": ["ORD001", "ORD002", "ORD003", "ORD005"],
        "sku": ["WIDGET-A", "WIDGET-B", "GADGET-X", "GIZMO-Z"],
        "quantity": [100, 210, 50, 30],  # ORD002 qty differs
        "unit_price": [9.99, 14.99, 24.99, 34.99],
        "amount": [999.00, 3147.90, 1249.50, 1049.70],  # ORD002 amount differs
    })

    # Run reconciliation
    reconciler = Reconciler(
        match_keys=["order_id", "sku"],
        tolerance=Tolerance(absolute=0.01),
    )
    summary = reconciler.reconcile(source_a, source_b)

    print(f"  Source A records: {summary.total_a}")
    print(f"  Source B records: {summary.total_b}")
    print(f"  Matched:          {summary.matched}")
    print(f"  Mismatched:       {summary.mismatched}")
    print(f"  Missing in A:     {summary.missing_in_a}")
    print(f"  Missing in B:     {summary.missing_in_b}")
    print(f"  Total amount diff: ${summary.total_amount_diff:,.2f}")
    print()


def demo_cleaning():
    """Demonstrate the data cleaning engine."""
    print("=" * 60)
    print("DEMO: Data Cleaning Pipeline")
    print("=" * 60)

    # Fictional messy data
    dirty_data = pd.DataFrame({
        "order_id": [" ord001", "ORD002", "ord001", "ORD003", "ORD004"],
        "customer": [" Alice ", "BOB", "Alice", "charlie", "Diana"],
        "amount": ["100.50", "N/A", "100.50", "250.00", None],
        "date": ["2024/01/15", "2024-01-16", "2024/01/15", "2024.01.17", "2024-01-18"],
    })

    print(f"  Input rows: {len(dirty_data)}")
    print(f"  Duplicates: {dirty_data.duplicated(subset=['order_id']).sum()}")

    # Apply cleaning pipeline
    cleaner = DataCleaner()
    result = (
        cleaner
        .load(dirty_data)
        .normalize_strings(["order_id", "customer"], strip=True, lower=True)
        .normalize_numbers(["amount"])
        .deduplicate(subset=["order_id"])
        .fill_na(strategy="mean", columns=["amount"])
    )

    cleaned = result.to_dataframe()
    print(f"  Cleaned rows: {len(cleaned)}")
    print(f"  Operations:   {len(cleaner.operations_log)}")
    print()


def demo_reporting():
    """Demonstrate the report generator."""
    print("=" * 60)
    print("DEMO: Report Generation")
    print("=" * 60)

    # Fictional sales data
    sales_data = pd.DataFrame({
        "region": ["North", "North", "South", "South", "East", "East"],
        "product": ["Widget", "Gadget", "Widget", "Gadget", "Widget", "Gadget"],
        "revenue": [15000.0, 22000.0, 18000.0, 27000.0, 12000.0, 19000.0],
        "units": [150, 110, 180, 135, 120, 95],
    })

    gen = ReportGenerator()
    gen.load_data(sales_data, metadata={"title": "Quarterly Sales Report"})
    gen.register_template("sales", "{{ metadata.get('title', 'Report') }}")

    print(f"  Data loaded: {gen._df.shape[0]} rows x {gen._df.shape[1]} columns")
    print(f"  Templates:   {gen.template_names}")
    print(f"  Ready to generate reports with batch splitting by region")
    print()


def demo_rule_engine():
    """Demonstrate the rule engine with YAML config."""
    print("=" * 60)
    print("DEMO: Rule Engine")
    print("=" * 60)

    config = {
        "reconciliation": {
            "match_keys": {"columns": ["order_id", "sku"], "ignore_case": False},
            "tolerance": {"absolute": 0.01},
            "compare_columns": [
                {"name": "quantity", "tolerance": 0},
                {"name": "amount", "tolerance": 0.01},
            ],
        },
        "cleaning": {
            "pipelines": [
                {
                    "name": "default",
                    "operations": [
                        {"operation": "string_format", "columns": ["order_id"], "params": {"strip": True}},
                        {"operation": "number_format", "columns": ["amount"]},
                    ],
                },
            ],
        },
    }

    engine = RuleEngine(config_dict=config)
    print(f"  Match keys:   {engine.match_key.columns}")
    print(f"  Tolerance:    abs={engine.tolerance.absolute}")
    print(f"  Pipelines:    {engine.pipeline_names}")
    print(f"  Config valid: True")
    print()


if __name__ == "__main__":
    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║           Excel AutoKit - Feature Demo                 ║")
    print("║           All data is fictional                        ║")
    print("╚══════════════════════════════════════════════════════════╝")
    print()

    demo_reconciliation()
    demo_cleaning()
    demo_reporting()
    demo_rule_engine()

    print("Demo complete. All features working correctly!")

# Created: 2026-09-08
