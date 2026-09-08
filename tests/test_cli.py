"""
Tests for the CLI interface (cli.py).
"""

from pathlib import Path

import pandas as pd
import pytest
from click.testing import CliRunner

from excel_autokit.cli import cli


@pytest.fixture
def cli_runner():
    """Provide a Click CLI test runner."""
    return CliRunner()


class TestCLIVersion:
    """Test CLI version and info commands."""

    def test_version(self, cli_runner):
        result = cli_runner.invoke(cli, ["--version"])
        assert result.exit_code == 0
        assert "excel-autokit" in result.output

    def test_info(self, cli_runner):
        result = cli_runner.invoke(cli, ["info"])
        assert result.exit_code == 0
        assert "excel-autokit" in result.output
        assert "Python" in result.output


class TestCLIReconcile:
    """Test the reconcile command."""

    def test_reconcile_basic(self, cli_runner, excel_file_a, excel_file_b, tmp_dir):
        output = str(tmp_dir / "cli_result.xlsx")
        result = cli_runner.invoke(cli, [
            "reconcile",
            str(excel_file_a),
            str(excel_file_b),
            "--keys", "order_id,sku",
            "--output", output,
        ])
        assert result.exit_code == 0
        assert "Matched" in result.output
        assert "Mismatched" in result.output
        assert Path(output).exists()

    def test_reconcile_with_tolerance(self, cli_runner, excel_file_a, excel_file_b, tmp_dir):
        output = str(tmp_dir / "cli_tol_result.xlsx")
        result = cli_runner.invoke(cli, [
            "reconcile",
            str(excel_file_a),
            str(excel_file_b),
            "--keys", "order_id,sku",
            "--output", output,
            "--tolerance", "0.1",
        ])
        assert result.exit_code == 0

    def test_reconcile_with_rules(self, cli_runner, excel_file_a, excel_file_b, yaml_config_file, tmp_dir):
        output = str(tmp_dir / "cli_rules_result.xlsx")
        result = cli_runner.invoke(cli, [
            "reconcile",
            str(excel_file_a),
            str(excel_file_b),
            "--keys", "order_id,sku",
            "--output", output,
            "--rules", str(yaml_config_file),
        ])
        assert result.exit_code == 0


class TestCLIClean:
    """Test the clean command."""

    def test_clean_basic(self, cli_runner, excel_dirty, tmp_dir):
        output = str(tmp_dir / "cli_cleaned.xlsx")
        result = cli_runner.invoke(cli, [
            "clean",
            str(excel_dirty),
            "--output", output,
            "--number-cols", "amount",
        ])
        assert result.exit_code == 0
        assert "Cleaned data" in result.output
        assert Path(output).exists()

    def test_clean_with_dedup(self, cli_runner, excel_dirty, tmp_dir):
        output = str(tmp_dir / "cli_dedup.xlsx")
        result = cli_runner.invoke(cli, [
            "clean",
            str(excel_dirty),
            "--output", output,
            "--dedup-cols", "id",
            "--fill-strategy", "zero",
        ])
        assert result.exit_code == 0

    def test_clean_with_rules(self, cli_runner, excel_dirty, yaml_clean_file, tmp_dir):
        output = str(tmp_dir / "cli_rules_clean.xlsx")
        result = cli_runner.invoke(cli, [
            "clean",
            str(excel_dirty),
            "--output", output,
            "--rules", str(yaml_clean_file),
            "--pipeline", "standard",
        ])
        assert result.exit_code == 0


class TestCLIReport:
    """Test the report command."""

    def test_report_basic(self, cli_runner, tmp_dir):
        # Create a simple input file
        df = pd.DataFrame({
            "name": ["Alice", "Bob"],
            "score": [95, 87],
        })
        input_path = tmp_dir / "input.xlsx"
        df.to_excel(str(input_path), index=False)
        output = str(tmp_dir / "cli_report.xlsx")

        result = cli_runner.invoke(cli, [
            "report",
            str(input_path),
            "--output", output,
        ])
        assert result.exit_code == 0
        assert "Report generated" in result.output
        assert Path(output).exists()

# Created: 2026-09-08
