"""Config: loads report-generator settings from a YAML file.

Default values live in  <project-root>/config.yaml.
Pass a custom YAML to Config.load(path) to override any subset of keys.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

_CONFIGS_DIR = Path(__file__).parent / 'configs'

# Bundled app constants (not for users to edit)
_APP_CONFIG_YAML = _CONFIGS_DIR / 'app_config.yaml'
# User-tunable settings (ships with sensible defaults, users edit this)
_USER_CONFIG_YAML = _CONFIGS_DIR / 'config.yaml'


# ── Sub-configs ────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class ColorsConfig:
    header: str
    mean: str
    positive: str
    negative: str

    @classmethod
    def _from_dict(cls, d: dict[str, str]) -> 'ColorsConfig':
        return cls(
            header=d['header'],
            mean=d['mean'],
            positive=d['positive'],
            negative=d['negative'],
        )


@dataclass(frozen=True)
class SheetNamesConfig:
    model1: str
    model2: str
    comparison: str
    excluded: str

    @classmethod
    def _from_dict(cls, d: dict[str, str]) -> 'SheetNamesConfig':
        return cls(
            model1=d['model1'],
            model2=d['model2'],
            comparison=d['comparison'],
            excluded=d['excluded'],
        )


@dataclass(frozen=True)
class BusinessConfig:
    target_perebrak: float
    target_nedobrak: float
    comparison_pct_threshold: float
    verdict_score_threshold: float
    column_translations: dict[str, str]

    # Source column names (what the raw Excel files contain)
    perebrak_col: str
    nedobrak_col: str
    f1_col: str

    # Goal / binary-flag column display names (Russian)
    col_perebrak_target: str
    col_nedobrak_target: str
    col_perebrak_gross: str
    col_nedobrak_gross: str

    @property
    def goal_cols(self) -> list[str]:
        return [
            self.col_perebrak_target,
            self.col_nedobrak_target,
            self.col_perebrak_gross,
            self.col_nedobrak_gross,
        ]

    @classmethod
    def _from_dict(cls, d: dict[str, Any]) -> 'BusinessConfig':
        return cls(
            target_perebrak=float(d.get('target_perebrak', 0.3)),
            target_nedobrak=float(d.get('target_nedobrak', 0.2)),
            comparison_pct_threshold=float(d.get('comparison_pct_threshold', 5.0)),
            verdict_score_threshold=float(d.get('verdict_score_threshold', 0.05)),
            column_translations=dict(d.get('column_translations') or {}),
            perebrak_col=str(d.get('perebrak_col', 'perebrak')),
            nedobrak_col=str(d.get('nedobrak_col', 'nedobrak')),
            f1_col=str(d.get('f1_col', 'f1_score')),
            col_perebrak_target=str(d.get('col_perebrak_target', 'Перебраковка в целях')),
            col_nedobrak_target=str(d.get('col_nedobrak_target', 'Недобраковка в целях')),
            col_perebrak_gross=str(d.get('col_perebrak_gross', 'Перебраковка грубо не в целях')),
            col_nedobrak_gross=str(d.get('col_nedobrak_gross', 'Недобраковка грубо не в целях')),
        )


# ── Main config ────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class Config:
    min_train_count: int
    degradation_threshold: float
    excluded_from_mean: frozenset[str]
    ratio_cols: frozenset[str]
    int_cols: frozenset[str]
    better_higher_cols: frozenset[str]
    better_lower_cols: frozenset[str]
    colors: ColorsConfig
    sheet_names: SheetNamesConfig
    business: BusinessConfig

    # ── Factories ──────────────────────────────────────────────────────────────

    @classmethod
    def default(cls) -> 'Config':
        """Load the bundled default config.yaml."""
        return cls.load()

    @classmethod
    def load(cls, path: str | Path | None = None) -> 'Config':
        """Load config by merging app constants with user settings.

        Load order (each layer overrides the previous):
        1. app_config.yaml  — bundled app constants, not for user editing
        2. config.yaml      — default user settings shipped with the tool
        3. *path*           — optional run-specific user override

        Parameters
        ----------
        path:
            Path to a user-supplied YAML file with additional overrides.
        """
        with open(_APP_CONFIG_YAML, encoding='utf-8') as fh:
            data: dict[str, Any] = yaml.safe_load(fh) or {}

        with open(_USER_CONFIG_YAML, encoding='utf-8') as fh:
            data = _deep_merge(data, yaml.safe_load(fh) or {})

        if path is not None:
            with open(path, encoding='utf-8') as fh:
                data = _deep_merge(data, yaml.safe_load(fh) or {})

        return cls._from_dict(data)

    @classmethod
    def _from_dict(cls, d: dict[str, Any]) -> 'Config':
        return cls(
            min_train_count=int(d['min_train_count']),
            degradation_threshold=float(d['degradation_threshold']),
            excluded_from_mean=frozenset(d['excluded_from_mean']),
            ratio_cols=frozenset(d['ratio_cols']),
            int_cols=frozenset(d['int_cols']),
            better_higher_cols=frozenset(d['better_higher_cols']),
            better_lower_cols=frozenset(d['better_lower_cols']),
            colors=ColorsConfig._from_dict(d['colors']),
            sheet_names=SheetNamesConfig._from_dict(d['sheet_names']),
            business=BusinessConfig._from_dict(d.get('business') or {}),
        )


# ── Helpers ────────────────────────────────────────────────────────────────────


def _deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge *override* into a copy of *base*."""
    result = base.copy()
    for key, val in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(val, dict):
            result[key] = _deep_merge(result[key], val)
        else:
            result[key] = val
    return result
