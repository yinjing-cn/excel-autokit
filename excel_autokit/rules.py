"""
Rule engine for excel_autokit.

Loads matching, validation, and cleaning rules from YAML configuration
files and provides a unified interface for rule evaluation.
"""

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import yaml

from excel_autokit.models import CleanOperation, CleanPipeline, MatchKey, Tolerance
from excel_autokit.utils import ConfigurationError, ensure_path, get_logger

logger = get_logger(__name__)


class RuleEngine:
    """YAML-driven rule engine for reconciliation and cleaning.

    Loads rule definitions from YAML files and provides typed access
    to match keys, tolerance settings, validation rules, and cleaning
    pipelines.

    Example YAML structure::

        reconciliation:
          match_keys:
            columns: ["order_id", "sku"]
            ignore_case: true
          tolerance:
            absolute: 0.01
            relative: 0.001
          compare_columns:
            - name: quantity
              tolerance: 0
            - name: amount
              tolerance: 0.01
        cleaning:
          pipelines:
            - name: standard_clean
              operations:
                - operation: date_format
                  columns: ["order_date", "ship_date"]
                  params:
                    format: "%Y-%m-%d"
    """

    def __init__(self, config_path: Optional[Union[str, Path]] = None, config_dict: Optional[Dict[str, Any]] = None):
        """Initialize the rule engine.

        Args:
            config_path: Path to a YAML configuration file.
            config_dict: Direct dictionary configuration (for testing).

        Raises:
            ConfigurationError: If neither path nor dict is provided,
                or if the YAML file cannot be loaded.
        """
        if config_dict is not None:
            self._config = config_dict
        elif config_path is not None:
            self._config = self._load_yaml(config_path)
        else:
            raise ConfigurationError("Either config_path or config_dict must be provided")

        self._validate_config()
        logger.info("RuleEngine initialized with %d top-level sections", len(self._config))

    @staticmethod
    def _load_yaml(path: Union[str, Path]) -> Dict[str, Any]:
        """Load and parse a YAML configuration file.

        Args:
            path: Path to the YAML file.

        Returns:
            Parsed configuration dictionary.

        Raises:
            ConfigurationError: If file not found or YAML is malformed.
        """
        p = ensure_path(path, create_parent=False)
        if not p.exists():
            raise ConfigurationError(f"Configuration file not found: {p}")
        try:
            with open(p, "r", encoding="utf-8") as fh:
                data = yaml.safe_load(fh)
            if not isinstance(data, dict):
                raise ConfigurationError(f"YAML root must be a mapping, got {type(data).__name__}")
            return data
        except yaml.YAMLError as exc:
            raise ConfigurationError(f"Invalid YAML in {p}: {exc}") from exc

    def _validate_config(self) -> None:
        """Validate the top-level structure of the configuration."""
        valid_sections = {"reconciliation", "cleaning", "reporting", "validation"}
        for key in self._config:
            if key not in valid_sections:
                logger.warning("Unknown configuration section: '%s'", key)

    # ------------------------------------------------------------------
    # Reconciliation rules
    # ------------------------------------------------------------------

    @property
    def reconciliation(self) -> Dict[str, Any]:
        """Return the reconciliation configuration section."""
        return self._config.get("reconciliation", {})

    @property
    def match_key(self) -> MatchKey:
        """Build and return a MatchKey from configuration.

        Returns:
            MatchKey instance.

        Raises:
            ConfigurationError: If match_keys section is missing or invalid.
        """
        recon = self.reconciliation
        mk_cfg = recon.get("match_keys")
        if not mk_cfg:
            raise ConfigurationError("Missing 'reconciliation.match_keys' in config")
        columns = mk_cfg.get("columns", [])
        if not columns:
            raise ConfigurationError("'reconciliation.match_keys.columns' must not be empty")
        return MatchKey(columns=columns, ignore_case=mk_cfg.get("ignore_case", False))

    @property
    def tolerance(self) -> Tolerance:
        """Build and return a Tolerance from configuration."""
        tol_cfg = self.reconciliation.get("tolerance", {})
        return Tolerance(
            absolute=tol_cfg.get("absolute", 0.0),
            relative=tol_cfg.get("relative", 0.0),
        )

    @property
    def compare_columns(self) -> List[Dict[str, Any]]:
        """Return the list of column comparison configs.

        Each entry is a dict with at least 'name' and optionally 'tolerance'.
        """
        return self.reconciliation.get("compare_columns", [])

    # ------------------------------------------------------------------
    # Cleaning rules
    # ------------------------------------------------------------------

    @property
    def cleaning(self) -> Dict[str, Any]:
        """Return the cleaning configuration section."""
        return self._config.get("cleaning", {})

    def get_pipeline(self, name: str) -> CleanPipeline:
        """Retrieve a named cleaning pipeline.

        Args:
            name: Pipeline name as defined in YAML.

        Returns:
            CleanPipeline instance.

        Raises:
            ConfigurationError: If the named pipeline is not found.
        """
        pipelines = self.cleaning.get("pipelines", [])
        for p_cfg in pipelines:
            if p_cfg.get("name") == name:
                ops = []
                for op_cfg in p_cfg.get("operations", []):
                    ops.append(
                        CleanOperation(
                            operation=op_cfg["operation"],
                            columns=op_cfg.get("columns"),
                            params=op_cfg.get("params", {}),
                        )
                    )
                return CleanPipeline(name=name, operations=ops)
        raise ConfigurationError(f"Cleaning pipeline '{name}' not found in config")

    @property
    def pipeline_names(self) -> List[str]:
        """Return all available cleaning pipeline names."""
        return [p.get("name", "") for p in self.cleaning.get("pipelines", [])]

    # ------------------------------------------------------------------
    # Reporting rules
    # ------------------------------------------------------------------

    @property
    def reporting(self) -> Dict[str, Any]:
        """Return the reporting configuration section."""
        return self._config.get("reporting", {})

    # ------------------------------------------------------------------
    # Validation rules
    # ------------------------------------------------------------------

    @property
    def validation(self) -> Dict[str, Any]:
        """Return the validation configuration section."""
        return self._config.get("validation", {})

    # ------------------------------------------------------------------
    # Raw access
    # ------------------------------------------------------------------

    @property
    def raw_config(self) -> Dict[str, Any]:
        """Return the raw configuration dictionary."""
        return dict(self._config)

# Created: 2026-09-08
