"""
Tests for the report generator (reporter.py).
"""

import pandas as pd
import pytest

from excel_autokit.reporter import ReportGenerator
from excel_autokit.utils import DataValidationError, TemplateError


class TestReportGeneratorInit:
    """Test ReportGenerator initialization."""

    def test_default_init(self):
        gen = ReportGenerator()
        assert gen.template_names == []

    def test_load_data(self, sample_report_data):
        gen = ReportGenerator()
        result = gen.load_data(sample_report_data)
        assert result is gen  # Returns self for chaining


class TestTemplateRegistration:
    """Test template registration."""

    def test_register_template_string(self):
        gen = ReportGenerator()
        gen.register_template("test", "Hello {{ metadata.get('name', 'World') }}")
        assert "test" in gen.template_names

    def test_register_template_file(self, tmp_dir):
        tmpl_path = tmp_dir / "template.txt"
        tmpl_path.write_text("Report: {{ row_count }} rows", encoding="utf-8")

        gen = ReportGenerator()
        gen.register_template_file("file_tmpl", tmpl_path)
        assert "file_tmpl" in gen.template_names

    def test_register_template_file_not_found(self, tmp_dir):
        gen = ReportGenerator()
        with pytest.raises(TemplateError, match="not found"):
            gen.register_template_file("missing", tmp_dir / "nonexistent.txt")


class TestReportGeneration:
    """Test report generation."""

    def test_generate_basic(self, sample_report_data, tmp_dir):
        gen = ReportGenerator()
        gen.load_data(sample_report_data)
        gen.register_template("basic", "Sales Report")

        output = tmp_dir / "report.xlsx"
        path = gen.generate("basic", output=output, include_summary=False)
        assert path.exists()

        # Verify content
        df = pd.read_excel(str(path))
        assert len(df) == 5

    def test_generate_with_summary(self, sample_report_data, tmp_dir):
        gen = ReportGenerator()
        gen.load_data(sample_report_data)
        gen.register_template("summary", "Report")

        output = tmp_dir / "report_summary.xlsx"
        path = gen.generate("summary", output=output, include_summary=True)
        assert path.exists()

    def test_generate_without_data_raises(self, tmp_dir):
        gen = ReportGenerator()
        gen.register_template("t", "Test")
        with pytest.raises(DataValidationError, match="No data loaded"):
            gen.generate("t", output=tmp_dir / "out.xlsx")

    def test_generate_unregistered_template_raises(self, sample_report_data, tmp_dir):
        gen = ReportGenerator()
        gen.load_data(sample_report_data)
        with pytest.raises(TemplateError, match="not registered"):
            gen.generate("nonexistent", output=tmp_dir / "out.xlsx")

    def test_generate_with_metadata(self, sample_report_data, tmp_dir):
        gen = ReportGenerator()
        gen.load_data(sample_report_data, metadata={"title": "Q4 Sales"})
        gen.register_template("meta", "{{ metadata.get('title', 'Default') }}")

        output = tmp_dir / "meta_report.xlsx"
        path = gen.generate("meta", output=output)
        assert path.exists()


class TestBatchGeneration:
    """Test batch report generation by group."""

    def test_batch_by_region(self, sample_report_data, tmp_dir):
        gen = ReportGenerator()
        gen.load_data(sample_report_data)
        gen.register_template("region", "Region Report")

        output_dir = tmp_dir / "batch_output"
        paths = gen.generate("region", output=output_dir, group_by=["region"])
        assert len(paths) == 3  # North, South, East

        for p in paths:
            assert p.exists()
            df = pd.read_excel(str(p))
            assert len(df) > 0

    def test_batch_verify_group_data(self, sample_report_data, tmp_dir):
        gen = ReportGenerator()
        gen.load_data(sample_report_data)
        gen.register_template("region", "Region")

        output_dir = tmp_dir / "batch_verify"
        paths = gen.generate("region", output=output_dir, group_by=["region"], include_summary=False)

        # Check that each file has only its group's data
        for p in paths:
            df = pd.read_excel(str(p))
            regions = df["region"].dropna().unique()
            assert len(regions) == 1  # Each file has exactly one region


class TestSheetNameSanitization:
    """Test sheet name sanitization."""

    def test_sanitize_long_name(self):
        gen = ReportGenerator()
        result = gen._sanitize_sheet_name("A" * 50)
        assert len(result) == 31

    def test_sanitize_special_chars(self):
        gen = ReportGenerator()
        result = gen._sanitize_sheet_name("Report/2024*Q1")
        assert "/" not in result
        assert "*" not in result

    def test_sanitize_empty(self):
        gen = ReportGenerator()
        result = gen._sanitize_sheet_name("")
        assert result == "Sheet1"

# Created: 2026-09-08
