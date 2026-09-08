# Excel AutoKit

**A pure Python CLI + library framework for Excel report automation — reconciliation, data cleaning, and report generation driven by YAML configuration.**

[![Python](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Tests](https://img.shields.io/badge/tests-pytest-blueviolet)](https://docs.pytest.org/)

---

## Features

- **Reconciliation Engine** — Compare two Excel sources by composite keys, detect value differences with configurable numeric tolerance, output styled results with highlighted discrepancies
- **Data Cleaning Pipeline** — Chainable operations for date/number/string normalization, deduplication, missing value handling, and multi-table merging
- **Report Generator** — Jinja2 template-based reports with batch generation by group keys, auto-summary rows, and optional chart embedding
- **YAML Rule Engine** — Define matching, validation, and cleaning rules in YAML — no code changes needed
- **CLI Interface** — Full-featured command-line tool for all operations
- **Library API** — Clean Python API for programmatic use in scripts and pipelines

## Quick Start

### Installation

```bash
pip install -e .
```

Or install dependencies directly:

```bash
pip install -r requirements.txt
```

### Minimal Example

```python
import pandas as pd
from excel_autokit import Reconciler
from excel_autokit.models import Tolerance

# Two fictional data sources
source_a = pd.DataFrame({
    "order_id": ["ORD001", "ORD002", "ORD003"],
    "amount": [100.0, 200.0, 300.0],
})

source_b = pd.DataFrame({
    "order_id": ["ORD001", "ORD002", "ORD004"],
    "amount": [100.0, 205.0, 400.0],
})

# Reconcile
reconciler = Reconciler(
    match_keys=["order_id"],
    tolerance=Tolerance(absolute=0.01),
)
summary = reconciler.reconcile(source_a, source_b)
reconciler.write_result("result.xlsx")

print(f"Matched: {summary.matched}, Mismatched: {summary.mismatched}")
```

## Architecture

```
excel_autokit/
├── models.py       — Data models (dataclass)
├── rules.py        — YAML rule engine
├── reconciler.py   — Reconciliation engine
├── cleaner.py      — Data cleaning pipeline
├── reporter.py     — Report generator
├── cli.py          — CLI entry point (click)
└── utils.py        — Shared utilities & exceptions
```

See [docs/architecture.md](docs/architecture.md) for the full architecture overview.

## Configuration

Define rules in YAML — no code changes required:

```yaml
reconciliation:
  match_keys:
    columns: ["order_id", "sku"]
    ignore_case: false
  tolerance:
    absolute: 0.01
  compare_columns:
    - name: quantity
      tolerance: 0
    - name: amount
      tolerance: 0.01

cleaning:
  pipelines:
    - name: standard
      operations:
        - operation: string_format
          columns: ["order_id", "sku"]
          params: { strip: true }
        - operation: number_format
          columns: ["amount"]
        - operation: dedup
          params: { subset: ["order_id"] }
```

See [examples/reconciliation_rules.yaml](examples/reconciliation_rules.yaml) for a complete example.

## CLI Usage

### Reconcile

```bash
excel-autokit reconcile source_a.xlsx source_b.xlsx \
  --keys order_id,sku \
  --output result.xlsx \
  --tolerance 0.01
```

### Clean

```bash
excel-autokit clean raw_data.xlsx \
  --output cleaned.xlsx \
  --dedup-cols order_id \
  --number-cols amount,quantity \
  --fill-strategy mean
```

### Report

```bash
excel-autokit report data.xlsx \
  --output report.xlsx \
  --template template.txt \
  --with-summary
```

### Info

```bash
excel-autokit info
```

## API Usage

### Data Cleaning Pipeline

```python
from excel_autokit import DataCleaner

cleaner = DataCleaner()
result = (
    cleaner
    .load("raw_data.xlsx")
    .normalize_dates(["order_date"], fmt="%Y-%m-%d")
    .normalize_numbers(["amount", "quantity"])
    .normalize_strings(["order_id"], strip=True, lower=True)
    .deduplicate(subset=["order_id"])
    .fill_na(strategy="mean", columns=["amount"])
    .to_dataframe()
)
```

### Report Generation with Batching

```python
from excel_autokit import ReportGenerator

gen = ReportGenerator()
gen.load_data(sales_df, metadata={"title": "Monthly Sales"})
gen.register_template("report", "{{ metadata.get('title') }}")

# Generate one report per region
paths = gen.generate("report", output="reports/", group_by=["region"])
```

### Rule-Driven Workflow

```python
from excel_autokit import RuleEngine, Reconciler

engine = RuleEngine(config_path="rules.yaml")
reconciler = Reconciler.from_rule_engine(engine)
summary = reconciler.reconcile("source_a.xlsx", "source_b.xlsx")
```

## Examples

Run the demo script to see all features in action:

```bash
python examples/demo.py
```

## Development

### Setup

```bash
git clone https://github.com/yinjing-cn/excel-autokit.git
cd excel-autokit
pip install -e ".[dev]"
```

### Running Tests

```bash
pytest
```

With coverage:

```bash
pytest --cov=excel_autokit --cov-report=term-missing
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Write tests for your changes
4. Ensure all tests pass (`pytest`)
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

---

<p align="center">
  <sub>Built with ❤️ for automating the boring Excel stuff.</sub>
</p>
