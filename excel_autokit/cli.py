"""
CLI entry point for excel_autokit.

Provides command-line interface for reconciliation, data cleaning,
and report generation using click.
"""

import sys
from pathlib import Path
from typing import Optional

import click

from excel_autokit import __version__
from excel_autokit.utils import get_logger

logger = get_logger(__name__)


@click.group()
@click.version_option(version=__version__, prog_name="excel-autokit")
@click.option("-v", "--verbose", is_flag=True, help="Enable verbose logging")
def cli(verbose: bool) -> None:
    """Excel AutoKit - Automated Excel report processing toolkit.

    A pure Python CLI framework for Excel reconciliation, data cleaning,
    and report generation driven by YAML configuration.
    """
    import logging as _logging
    if verbose:
        logger.setLevel(_logging.DEBUG)


# ------------------------------------------------------------------
# Reconcile command
# ------------------------------------------------------------------

@cli.command()
@click.argument("source_a", type=click.Path(exists=True))
@click.argument("source_b", type=click.Path(exists=True))
@click.option("-o", "--output", default="reconcile_result.xlsx", help="Output file path")
@click.option("-k", "--keys", required=True, help="Comma-separated match key columns")
@click.option("-r", "--rules", type=click.Path(exists=True), help="YAML rules file")
@click.option("--tolerance", type=float, default=0.0, help="Numeric tolerance for comparison")
@click.option("--ignore-case", is_flag=True, help="Ignore case in string key matching")
@click.option("--sheet-a", default=0, help="Sheet name/index for source A")
@click.option("--sheet-b", default=0, help="Sheet name/index for source B")
def reconcile(
    source_a: str,
    source_b: str,
    output: str,
    keys: str,
    rules: Optional[str],
    tolerance: float,
    ignore_case: bool,
    sheet_a: str,
    sheet_b: str,
) -> None:
    """Reconcile two Excel files and output differences.

    SOURCE_A and SOURCE_B are paths to the Excel files to compare.
    """
    from excel_autokit.models import Tolerance
    from excel_autokit.reconciler import Reconciler

    match_keys = [k.strip() for k in keys.split(",")]

    if rules:
        from excel_autokit.rules import RuleEngine
        engine = RuleEngine(config_path=rules)
        reconciler = Reconciler.from_rule_engine(engine)
    else:
        reconciler = Reconciler(
            match_keys=match_keys,
            tolerance=Tolerance(absolute=tolerance),
            ignore_case=ignore_case,
        )

    click.echo(f"Reconciling: {source_a} vs {source_b}")
    summary = reconciler.reconcile(source_a, source_b)

    click.echo(f"  Source A: {summary.total_a} records")
    click.echo(f"  Source B: {summary.total_b} records")
    click.echo(f"  Matched:  {summary.matched}")
    click.echo(f"  Mismatched: {summary.mismatched}")
    click.echo(f"  Missing in A: {summary.missing_in_a}")
    click.echo(f"  Missing in B: {summary.missing_in_b}")
    click.echo(f"  Amount diff:  {summary.total_amount_diff:.4f}")

    result_path = reconciler.write_result(output)
    click.echo(f"Result written to: {result_path}")


# ------------------------------------------------------------------
# Clean command
# ------------------------------------------------------------------

@cli.command()
@click.argument("input_file", type=click.Path(exists=True))
@click.option("-o", "--output", default="cleaned_output.xlsx", help="Output file path")
@click.option("-r", "--rules", type=click.Path(exists=True), help="YAML rules file with cleaning pipelines")
@click.option("-p", "--pipeline", help="Name of the cleaning pipeline to apply")
@click.option("--dedup-cols", help="Comma-separated columns for deduplication")
@click.option("--date-cols", help="Comma-separated date columns to normalize")
@click.option("--number-cols", help="Comma-separated numeric columns to normalize")
@click.option("--fill-strategy", type=click.Choice(["zero", "mean", "median", "ffill", "bfill", "drop"]),
              default="zero", help="Missing value fill strategy")
def clean(
    input_file: str,
    output: str,
    rules: Optional[str],
    pipeline: Optional[str],
    dedup_cols: Optional[str],
    date_cols: Optional[str],
    number_cols: Optional[str],
    fill_strategy: str,
) -> None:
    """Clean and standardize an Excel file.

    INPUT_FILE is the path to the Excel file to clean.
    """
    from excel_autokit.cleaner import DataCleaner

    cleaner = DataCleaner()
    cleaner.load(input_file)

    if rules and pipeline:
        from excel_autokit.rules import RuleEngine
        engine = RuleEngine(config_path=rules)
        pipe = engine.get_pipeline(pipeline)
        cleaner.apply_pipeline(pipe)
    else:
        # Apply CLI-specified operations
        if date_cols:
            cols = [c.strip() for c in date_cols.split(",")]
            cleaner.normalize_dates(cols)
        if number_cols:
            cols = [c.strip() for c in number_cols.split(",")]
            cleaner.normalize_numbers(cols)
        if dedup_cols:
            cols = [c.strip() for c in dedup_cols.split(",")]
            cleaner.deduplicate(subset=cols)
        cleaner.fill_na(strategy=fill_strategy)

    result_path = cleaner.to_excel(output)
    click.echo(f"Cleaned data written to: {result_path}")
    click.echo(f"Operations applied: {len(cleaner.operations_log)}")


# ------------------------------------------------------------------
# Report command
# ------------------------------------------------------------------

@cli.command()
@click.argument("input_file", type=click.Path(exists=True))
@click.option("-o", "--output", default="report.xlsx", help="Output file path")
@click.option("-t", "--template", type=click.Path(exists=True), help="Jinja2 template file")
@click.option("--template-name", default="default", help="Template name for registration")
@click.option("--group-by", help="Comma-separated columns for batch report splitting")
@click.option("--with-summary", is_flag=True, default=True, help="Include summary rows")
@click.option("--charts", help="Comma-separated columns for chart generation")
def report(
    input_file: str,
    output: str,
    template: Optional[str],
    template_name: str,
    group_by: Optional[str],
    with_summary: bool,
    charts: Optional[str],
) -> None:
    """Generate a formatted Excel report.

    INPUT_FILE is the path to the source Excel file.
    """
    import pandas as pd

    from excel_autokit.reporter import ReportGenerator
    from excel_autokit.utils import read_excel_safe

    gen = ReportGenerator()
    df = read_excel_safe(input_file)
    gen.load_data(df)

    # Register template
    if template:
        gen.register_template_file(template_name, template)
    else:
        # Default template: use filename as sheet name
        default_tmpl = "{{ metadata.get('title', 'Report') }}"
        gen.register_template(template_name, default_tmpl)

    group_cols = [c.strip() for c in group_by.split(",")] if group_by else None
    chart_cols = [c.strip() for c in charts.split(",")] if charts else None

    result = gen.generate(
        template_name=template_name,
        output=output,
        group_by=group_cols,
        include_summary=with_summary,
        chart_columns=chart_cols,
    )

    if isinstance(result, list):
        click.echo(f"Batch reports generated: {len(result)} files")
    else:
        click.echo(f"Report generated: {result}")


# ------------------------------------------------------------------
# Info command
# ------------------------------------------------------------------

@cli.command()
def info() -> None:
    """Show excel_autokit version and environment info."""
    click.echo(f"excel-autokit v{__version__}")
    click.echo(f"Python: {sys.version}")
    try:
        import pandas
        click.echo(f"pandas: {pandas.__version__}")
    except ImportError:
        click.echo("pandas: not installed")
    try:
        import openpyxl
        click.echo(f"openpyxl: {openpyxl.__version__}")
    except ImportError:
        click.echo("openpyxl: not installed")


def main() -> None:
    """Main entry point."""
    cli()


if __name__ == "__main__":
    main()

# Created: 2026-09-08
