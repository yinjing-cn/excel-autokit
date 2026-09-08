"""
Tests for the data cleaning engine (cleaner.py).
"""

import pandas as pd
import pytest

from excel_autokit.cleaner import DataCleaner
from excel_autokit.utils import DataValidationError


class TestDataCleanerLoad:
    """Test DataCleaner data loading."""

    def test_load_from_dataframe(self, sample_dirty_data):
        cleaner = DataCleaner()
        result = cleaner.load(sample_dirty_data)
        assert result is cleaner  # Returns self for chaining
        df = cleaner.to_dataframe()
        assert len(df) == 6

    def test_load_from_file(self, excel_dirty):
        cleaner = DataCleaner()
        cleaner.load(excel_dirty)
        df = cleaner.to_dataframe()
        assert len(df) == 6

    def test_to_dataframe_without_load_raises(self):
        cleaner = DataCleaner()
        with pytest.raises(DataValidationError, match="No data loaded"):
            cleaner.to_dataframe()


class TestDateNormalization:
    """Test date format normalization."""

    def test_normalize_dates(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.normalize_dates(["date_col"])
        df = cleaner.to_dataframe()
        # Should have converted to datetime (some may be NaT due to bad formats)
        assert pd.api.types.is_datetime64_any_dtype(df["date_col"])

    def test_normalize_dates_invalid_columns_raises(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        with pytest.raises(DataValidationError):
            cleaner.normalize_dates(["nonexistent_col"])


class TestNumberNormalization:
    """Test numeric format normalization."""

    def test_normalize_numbers(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.normalize_numbers(["amount"])
        df = cleaner.to_dataframe()
        assert pd.api.types.is_numeric_dtype(df["amount"])
        # "N/A" should become NaN
        assert df["amount"].isna().sum() >= 1

    def test_normalize_numbers_invalid_columns_raises(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        with pytest.raises(DataValidationError):
            cleaner.normalize_numbers(["nonexistent"])


class TestStringNormalization:
    """Test string normalization."""

    def test_strip_whitespace(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.normalize_strings(["name"], strip=True)
        df = cleaner.to_dataframe()
        # " Alice " should become "Alice"
        assert "Alice" in df["name"].values

    def test_lowercase(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.normalize_strings(["name"], strip=True, lower=True)
        df = cleaner.to_dataframe()
        assert all(v == v.lower() for v in df["name"] if v != "nan")


class TestDeduplication:
    """Test deduplication."""

    def test_deduplicate_all_columns(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.deduplicate()
        df = cleaner.to_dataframe()
        assert len(df) < 6  # Should remove at least one duplicate

    def test_deduplicate_by_subset(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.deduplicate(subset=["id"])
        df = cleaner.to_dataframe()
        # id=2 appears twice, should be deduped
        assert len(df) == 5
        assert df["id"].nunique() == len(df)

    def test_detect_duplicates(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        dups = cleaner.detect_duplicates(subset=["id"])
        assert len(dups) == 2  # Two rows with id=2


class TestMissingValueHandling:
    """Test missing value strategies."""

    def test_fill_na_zero(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.normalize_numbers(["amount"])
        cleaner.fill_na(strategy="zero", columns=["amount"])
        df = cleaner.to_dataframe()
        assert df["amount"].isna().sum() == 0

    def test_fill_na_mean(self):
        df = pd.DataFrame({"val": [10.0, None, 30.0]})
        cleaner = DataCleaner()
        cleaner.load(df)
        cleaner.fill_na(strategy="mean", columns=["val"])
        result = cleaner.to_dataframe()
        assert result["val"].iloc[1] == 20.0  # mean of 10 and 30

    def test_fill_na_value(self):
        df = pd.DataFrame({"val": [1, None, 3]})
        cleaner = DataCleaner()
        cleaner.load(df)
        cleaner.fill_na(strategy="value", columns=["val"], value=99)
        result = cleaner.to_dataframe()
        assert result["val"].iloc[1] == 99

    def test_fill_na_drop(self):
        df = pd.DataFrame({"val": [1, None, 3]})
        cleaner = DataCleaner()
        cleaner.load(df)
        cleaner.fill_na(strategy="drop", columns=["val"])
        result = cleaner.to_dataframe()
        assert len(result) == 2

    def test_mark_missing(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.normalize_numbers(["amount"])
        cleaner.mark_missing(["amount"])
        df = cleaner.to_dataframe()
        assert "amount_was_missing" in df.columns


class TestMerging:
    """Test table merging."""

    def test_merge_left(self, sample_merge_left, sample_merge_right):
        cleaner = DataCleaner()
        cleaner.load(sample_merge_left)
        cleaner.merge(sample_merge_right, on=["key"], how="left")
        df = cleaner.to_dataframe()
        assert len(df) == 3
        assert "val_right" in df.columns
        # K3 has no match in right
        assert df.loc[df["key"] == "K3", "val_right"].isna().iloc[0]

    def test_merge_inner(self, sample_merge_left, sample_merge_right):
        cleaner = DataCleaner()
        cleaner.load(sample_merge_left)
        cleaner.merge(sample_merge_right, on=["key"], how="inner")
        df = cleaner.to_dataframe()
        assert len(df) == 2  # Only K1 and K2

    def test_merge_outer(self, sample_merge_left, sample_merge_right):
        cleaner = DataCleaner()
        cleaner.load(sample_merge_left)
        cleaner.merge(sample_merge_right, on=["key"], how="outer")
        df = cleaner.to_dataframe()
        assert len(df) == 4  # K1, K2, K3, K4

    def test_merge_from_file(self, sample_merge_left, sample_merge_right, tmp_dir):
        right_path = tmp_dir / "right.xlsx"
        sample_merge_right.to_excel(str(right_path), index=False)

        cleaner = DataCleaner()
        cleaner.load(sample_merge_left)
        cleaner.merge(right_path, on=["key"], how="inner")
        df = cleaner.to_dataframe()
        assert len(df) == 2


class TestPipeline:
    """Test cleaning pipeline application."""

    def test_apply_pipeline(self, sample_dirty_data, sample_clean_config):
        from excel_autokit.models import CleanOperation, CleanPipeline
        from excel_autokit.rules import RuleEngine

        engine = RuleEngine(config_dict=sample_clean_config)
        pipeline = engine.get_pipeline("standard")

        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.apply_pipeline(pipeline)
        df = cleaner.to_dataframe()

        # After pipeline: strings stripped+lowered, numbers coerced, deduped by id
        assert len(df) == 5  # Duplicate id=2 removed
        assert "alice" in df["name"].values

    def test_operations_log(self, sample_dirty_data):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.normalize_strings(["name"])
        cleaner.deduplicate(subset=["id"])
        assert len(cleaner.operations_log) == 3  # load + 2 ops

    def test_to_excel_output(self, sample_dirty_data, tmp_dir):
        cleaner = DataCleaner()
        cleaner.load(sample_dirty_data)
        cleaner.normalize_strings(["name"])
        output = tmp_dir / "cleaned.xlsx"
        path = cleaner.to_excel(output)
        assert path.exists()

        df = pd.read_excel(str(path))
        assert len(df) == 6

    def test_chaining(self, sample_dirty_data):
        """Verify method chaining works correctly."""
        cleaner = DataCleaner()
        result = (
            cleaner
            .load(sample_dirty_data)
            .normalize_strings(["name"], strip=True, lower=True)
            .normalize_numbers(["amount"])
            .deduplicate(subset=["id"])
        )
        assert result is cleaner
        df = cleaner.to_dataframe()
        assert len(df) == 5

# Created: 2026-09-08
