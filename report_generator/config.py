"""Config: loads report-generator settings from a YAML file.

Default values live in  <project-root>/config.yaml.
Pass a custom YAML to Config.load(path) to override any subset of keys.
"""

import math
import re
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
    verdict: str

    @classmethod
    def _from_dict(cls, d: dict[str, str]) -> 'SheetNamesConfig':
        return cls(
            model1=d['model1'],
            model2=d['model2'],
            comparison=d['comparison'],
            excluded=d['excluded'],
            verdict=d['verdict'],
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
            data: dict[str, Any] = _mapping(yaml.safe_load(fh), 'app config')

        with open(_USER_CONFIG_YAML, encoding='utf-8') as fh:
            data = _deep_merge(data, _mapping(yaml.safe_load(fh), 'config'), validate_keys=False)

        if path is not None:
            with open(path, encoding='utf-8') as fh:
                data = _deep_merge(data, _mapping(yaml.safe_load(fh), 'config'))

        _validate(data)
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


def _deep_merge(base: dict, override: dict, *, validate_keys: bool = True) -> dict:
    """Recursively merge *override* into a copy of *base*."""
    override = _mapping(override, 'configuration override')
    unknown = set(override) - set(base)
    if validate_keys and unknown:
        raise ValueError(f'Unknown configuration keys: {sorted(unknown)}')
    result = base.copy()
    for key, val in override.items():
        if key == 'column_translations':
            result[key] = {**result[key], **_mapping(val, key)}
        elif key in result and isinstance(result[key], dict) and isinstance(val, dict):
            result[key] = _deep_merge(result[key], val, validate_keys=validate_keys)
        else:
            result[key] = val
    return result


def _mapping(value: object, label: str) -> dict:
    if not isinstance(value, dict):
        raise ValueError(f'{label} must be a mapping')
    if not all(isinstance(key, str) for key in value):
        raise ValueError(f'{label} keys must be strings')
    return value


def _validate(data: dict) -> None:
    def number(value: object, label: str, low: float, high: float | None = None) -> None:
        if (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not math.isfinite(value)
            or value < low
            or (high is not None and value > high)
        ):
            raise ValueError(f'{label} must be a finite number in range {low}..{high}')

    number(data['min_train_count'], 'min_train_count', 0)
    if not isinstance(data['min_train_count'], int):
        raise ValueError('min_train_count must be an integer')
    number(data['degradation_threshold'], 'degradation_threshold', 0, 1)
    for key in [
        'excluded_from_mean',
        'ratio_cols',
        'int_cols',
        'better_higher_cols',
        'better_lower_cols',
    ]:
        value = data[key]
        if not isinstance(value, list) or not all(isinstance(v, str) and v for v in value):
            raise ValueError(f'{key} must be a list of nonempty column names')
    overlap = set(data['better_higher_cols']) & set(data['better_lower_cols'])
    if overlap:
        raise ValueError(f'Conflicting metric directions: {sorted(overlap)}')
    colors = _mapping(data['colors'], 'colors')
    for key, value in colors.items():
        if not isinstance(value, str) or not re.fullmatch(
            r'[0-9a-fA-F]{6}([0-9a-fA-F]{2})?', value
        ):
            raise ValueError(f'colors.{key} must be a 6 or 8 digit hex color')
    names = _mapping(data['sheet_names'], 'sheet_names')
    for key, value in names.items():
        if (
            not isinstance(value, str)
            or not value.strip()
            or len(value) > 31
            or re.search(r'[\\/*?:\[\]]', value)
            or value.startswith("'")
            or value.endswith("'")
        ):
            raise ValueError(f'Invalid Excel worksheet name: sheet_names.{key}')
    if len({value.casefold() for value in names.values()}) != len(names):
        raise ValueError('Worksheet names must be unique (case insensitive)')
    biz = _mapping(data['business'], 'business')
    for key in ['target_perebrak', 'target_nedobrak', 'verdict_score_threshold']:
        number(biz[key], f'business.{key}', 0, 1)
    number(biz['comparison_pct_threshold'], 'business.comparison_pct_threshold', 0, 100)
    for key, value in biz.items():
        if key.endswith('_col') or key.startswith('col_'):
            if not isinstance(value, str) or not value.strip():
                raise ValueError(f'business.{key} must be a nonempty column name')
    translations = _mapping(biz['column_translations'], 'business.column_translations')
    if not all(
        isinstance(k, str) and k and isinstance(v, str) and v.strip()
        for k, v in translations.items()
    ):
        raise ValueError('column_translations must map nonempty strings to nonempty strings')
    source_cols = [biz[key] for key in ('perebrak_col', 'nedobrak_col', 'f1_col')]
    error_sources = {biz['perebrak_col'], biz['nedobrak_col']}
    if (
        error_sources & set(data['better_higher_cols'])
        or biz['f1_col'] in data['better_lower_cols']
    ):
        raise ValueError('Business source role conflicts with configured metric direction')
    metadata = set(data['int_cols']) | {
        'Класс',
        'ID',
        'confidence',
        'tp',
        'fp',
        'fn',
        'Причина',
        'Модель',
        'Число примеров',
    }
    for source in source_cols:
        identity = source.casefold()
        if (
            source in metadata
            or identity in {'id', 'confidence', 'tp', 'fp', 'fn'}
            or 'count' in identity
            or 'пример' in identity
            or re.search(r'(^|[_\s])n([_\s]|$)', identity)
        ):
            raise ValueError(f'Business source cannot alias immutable metadata: {source!r}')
    goal_cols = [
        biz[key]
        for key in (
            'col_perebrak_target',
            'col_nedobrak_target',
            'col_perebrak_gross',
            'col_nedobrak_gross',
        )
    ]
    reserved = set().union(
        *[
            set(data[key])
            for key in (
                'ratio_cols',
                'int_cols',
                'better_higher_cols',
                'better_lower_cols',
            )
        ],
        {'Класс', 'Причина', 'Модель', 'Число примеров'},
    )
    if (
        len(set(source_cols)) != 3
        or len(set(goal_cols)) != 4
        or set(goal_cols) & (reserved | set(source_cols))
    ):
        raise ValueError('Generated/source column name collision')
    identities = set().union(
        *(
            set(data[key])
            for key in [
                'excluded_from_mean',
                'ratio_cols',
                'int_cols',
                'better_higher_cols',
                'better_lower_cols',
            ]
        )
    )
    identities.update([biz[key] for key in biz if key.endswith('_col') or key.startswith('col_')])
    identities.update(['Класс', 'Число примеров', 'Причина', 'Модель'])
    identities.update(translations)
    identities = {
        key
        for key in identities
        if not (key not in translations and 'пример' in key and key in translations.values())
    }
    displayed = [translations.get(key, key) for key in identities]
    if len(displayed) != len(set(displayed)):
        raise ValueError('Translated column name collision')
