"""
Tests for the rule engine (rules.py).
"""

import pytest
import yaml

from excel_autokit.rules import RuleEngine
from excel_autokit.utils import ConfigurationError


class TestRuleEngineInit:
    """Test RuleEngine initialization."""

    def test_init_from_dict(self, sample_recon_config):
        engine = RuleEngine(config_dict=sample_recon_config)
        assert engine.raw_config == sample_recon_config

    def test_init_from_yaml_file(self, yaml_config_file):
        engine = RuleEngine(config_path=yaml_config_file)
        assert "reconciliation" in engine.raw_config

    def test_init_no_args_raises(self):
        with pytest.raises(ConfigurationError, match="Either config_path or config_dict"):
            RuleEngine()

    def test_init_missing_file_raises(self, tmp_dir):
        with pytest.raises(ConfigurationError, match="not found"):
            RuleEngine(config_path=tmp_dir / "nonexistent.yaml")

    def test_init_invalid_yaml_raises(self, tmp_dir):
        bad_yaml = tmp_dir / "bad.yaml"
        bad_yaml.write_text("{{invalid yaml content::", encoding="utf-8")
        with pytest.raises(ConfigurationError, match="Invalid YAML"):
            RuleEngine(config_path=bad_yaml)

    def test_init_non_mapping_yaml_raises(self, tmp_dir):
        list_yaml = tmp_dir / "list.yaml"
        list_yaml.write_text("- item1\n- item2\n", encoding="utf-8")
        with pytest.raises(ConfigurationError, match="YAML root must be a mapping"):
            RuleEngine(config_path=list_yaml)


class TestMatchKey:
    """Test match key extraction from config."""

    def test_match_key_basic(self, sample_recon_config):
        engine = RuleEngine(config_dict=sample_recon_config)
        mk = engine.match_key
        assert mk.columns == ["order_id", "sku"]
        assert mk.ignore_case is False

    def test_match_key_missing_raises(self):
        engine = RuleEngine(config_dict={"reconciliation": {}})
        with pytest.raises(ConfigurationError, match="Missing"):
            _ = engine.match_key

    def test_match_key_empty_columns_raises(self):
        engine = RuleEngine(config_dict={"reconciliation": {"match_keys": {"columns": []}}})
        with pytest.raises(ConfigurationError, match="must not be empty"):
            _ = engine.match_key

    def test_match_key_extract(self, sample_recon_config):
        engine = RuleEngine(config_dict=sample_recon_config)
        mk = engine.match_key
        row = {"order_id": "ORD001", "sku": "SKU-A1", "amount": 100}
        key = mk.extract(row)
        assert key == ("ORD001", "SKU-A1")

    def test_match_key_extract_ignore_case(self):
        cfg = {"reconciliation": {"match_keys": {"columns": ["id"], "ignore_case": True}}}
        engine = RuleEngine(config_dict=cfg)
        mk = engine.match_key
        key = mk.extract({"id": "hello"})
        assert key == ("HELLO",)


class TestTolerance:
    """Test tolerance configuration."""

    def test_default_tolerance(self, sample_recon_config):
        engine = RuleEngine(config_dict=sample_recon_config)
        tol = engine.tolerance
        assert tol.absolute == 0.01
        assert tol.relative == 0.0

    def test_tolerance_missing_section(self):
        engine = RuleEngine(config_dict={"reconciliation": {"match_keys": {"columns": ["id"]}}})
        tol = engine.tolerance
        assert tol.absolute == 0.0
        assert tol.relative == 0.0


class TestCompareColumns:
    """Test compare columns configuration."""

    def test_compare_columns(self, sample_recon_config):
        engine = RuleEngine(config_dict=sample_recon_config)
        cols = engine.compare_columns
        assert len(cols) == 2
        assert cols[0]["name"] == "quantity"
        assert cols[0]["tolerance"] == 0


class TestCleaningPipelines:
    """Test cleaning pipeline retrieval."""

    def test_get_pipeline(self, sample_clean_config):
        engine = RuleEngine(config_dict=sample_clean_config)
        pipeline = engine.get_pipeline("standard")
        assert pipeline.name == "standard"
        assert len(pipeline.operations) == 3
        assert pipeline.operations[0].operation == "string_format"

    def test_get_pipeline_not_found(self, sample_clean_config):
        engine = RuleEngine(config_dict=sample_clean_config)
        with pytest.raises(ConfigurationError, match="not found"):
            engine.get_pipeline("nonexistent")

    def test_pipeline_names(self, sample_clean_config):
        engine = RuleEngine(config_dict=sample_clean_config)
        names = engine.pipeline_names
        assert names == ["standard"]


class TestRawConfig:
    """Test raw config access."""

    def test_raw_config_returns_copy(self, sample_recon_config):
        engine = RuleEngine(config_dict=sample_recon_config)
        raw = engine.raw_config
        assert raw == sample_recon_config
        raw["new_key"] = "value"
        assert "new_key" not in engine.raw_config

    def test_section_properties_empty(self):
        engine = RuleEngine(config_dict={})
        assert engine.reconciliation == {}
        assert engine.cleaning == {}
        assert engine.reporting == {}
        assert engine.validation == {}

# Created: 2026-09-08
