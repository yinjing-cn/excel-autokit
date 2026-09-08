# Architecture

## Overview

Excel AutoKit follows a modular architecture with clear separation of concerns.
Each core module is independently usable and can be combined via YAML-driven
configuration.

```
┌─────────────────────────────────────────────────┐
│                   CLI Layer                      │
│               (click commands)                   │
├─────────────────────────────────────────────────┤
│                                                  │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐      │
│  │Reconciler│  │ Cleaner  │  │ Reporter │      │
│  │          │  │          │  │          │      │
│  │ Match    │  │ Format   │  │ Template │      │
│  │ Diff     │  │ Dedup    │  │ Charts   │      │
│  │ Summary  │  │ Fill NA  │  │ Batch    │      │
│  └────┬─────┘  └────┬─────┘  └────┬─────┘      │
│       │              │              │            │
│       └──────────┬───┴──────────────┘            │
│                  │                               │
│           ┌──────┴──────┐                        │
│           │ Rule Engine  │                       │
│           │ (YAML)       │                       │
│           └──────┬──────┘                        │
│                  │                               │
│           ┌──────┴──────┐                        │
│           │   Models    │                        │
│           │ (dataclass) │                        │
│           └─────────────┘                        │
│                                                  │
├─────────────────────────────────────────────────┤
│              Utilities / Exceptions              │
└─────────────────────────────────────────────────┘
```

## Module Responsibilities

### `models.py`
Pure data containers using `@dataclass`. No business logic.
- `MatchKey` — composite key definition and extraction
- `Tolerance` — numeric comparison tolerance
- `ReconcileResult` / `ReconcileSummary` — reconciliation outputs
- `CleanOperation` / `CleanPipeline` — cleaning pipeline definitions
- `ReportConfig` — report generation settings

### `rules.py`
YAML configuration loader and accessor.
- Validates configuration structure
- Provides typed access to match keys, tolerance, pipelines
- Supports both file-based and dictionary-based initialization

### `reconciler.py`
Core reconciliation engine.
- Loads two data sources (Excel or DataFrame)
- Builds composite-key lookup maps
- Performs row-level comparison with tolerance support
- Produces styled Excel output with difference highlighting
- Generates summary statistics

### `cleaner.py`
Chainable data cleaning engine.
- Format normalization (dates, numbers, strings, encoding)
- Deduplication (full-row or subset-based)
- Missing value strategies (zero, mean, median, forward/back fill, custom)
- Table merging (left, right, inner, outer joins)
- Pipeline support via `CleanPipeline` objects

### `reporter.py`
Report generation engine.
- Jinja2 template rendering for dynamic content
- Summary/subtotal row injection
- Batch generation by group keys
- Optional chart embedding via matplotlib

### `cli.py`
Command-line interface built with Click.
- `reconcile` — compare two Excel files
- `clean` — clean and standardize data
- `report` — generate formatted reports
- `info` — show version and environment

## Design Principles

1. **Chainable API** — Cleaner uses method chaining for fluent pipeline definition
2. **Configuration over code** — YAML rules drive behavior without code changes
3. **DataFrame interop** — All modules accept and return pandas DataFrames
4. **Fail fast** — Clear error messages with custom exception hierarchy
5. **Testability** — All modules work with in-memory DataFrames for easy testing
