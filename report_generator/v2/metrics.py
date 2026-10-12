"""Structured metrics keep class strings exact; the legacy reader remains separate."""

import math
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from ..app_config import CLASS_COL


def _valid_id(value: object) -> bool:
    if isinstance(value, str):
        return bool(value.strip())
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


@dataclass(frozen=True)
class MetricsTable:
    """A defensive table snapshot whose lexical class names are not normalized.

    Unlike the legacy Excel reader, numeric-looking strings, leading zeros,
    literal NA and significant surrounding whitespace are retained. Numeric
    footer cells and aggregate rows must be removed by the input reader.
    """

    frame: pd.DataFrame
    source_name: str
    source_path: Path | None = None

    def __post_init__(self) -> None:
        frame = self.frame.copy(deep=True)
        if not isinstance(self.source_name, str) or not self.source_name:
            raise ValueError('Metrics table requires a nonempty source_name')
        if CLASS_COL not in frame:
            raise ValueError(f'Metrics table requires {CLASS_COL!r}: {self.source_name}')
        if not frame.columns.is_unique:
            raise ValueError(f'Duplicate metric columns: {self.source_name}')
        labels = frame[CLASS_COL]
        if not labels.map(
            lambda value: isinstance(value, str) and bool(value.strip()) and value != 'Среднее'
        ).all():
            raise ValueError(f'Metrics table requires nonempty lexical classes: {self.source_name}')
        duplicates = frame.loc[labels.duplicated(), CLASS_COL].tolist()
        if duplicates:
            raise ValueError(
                f'Duplicate classes in {self.source_path or self.source_name}: {duplicates}'
            )
        if 'ID' in frame:
            ids = frame['ID']
            if ids.isna().any() or not ids.map(_valid_id).all() or ids.duplicated().any():
                raise ValueError(
                    'ID values must be nonempty unique strings or numbers: '
                    f'{self.source_path or self.source_name}'
                )
        object.__setattr__(self, 'frame', frame.reset_index(drop=True))
        if self.source_path is not None:
            object.__setattr__(self, 'source_path', Path(self.source_path))


def read_metrics_table(file_path: str | Path, sheet_name: str | int = 0) -> MetricsTable:
    """Read exact text class cells while retaining pandas missing metric handling."""
    path = Path(file_path)
    with pd.ExcelFile(path) as workbook:
        frame = workbook.parse(sheet_name)
        if len(frame.columns) == 0:
            raise ValueError(f'Empty worksheet: {sheet_name!r} in {path}')
        labels = workbook.parse(sheet_name, usecols=[0], dtype=object, keep_default_na=False).iloc[
            :, 0
        ]
    if len(labels) != len(frame):
        raise ValueError(f'Dashboard class/metric rows disagree: {path}')
    frame = frame.rename(columns={frame.columns[0]: CLASS_COL})
    frame[CLASS_COL] = labels
    keep = labels.map(
        lambda value: isinstance(value, str) and bool(value.strip()) and value != 'Среднее'
    )
    return MetricsTable(frame.loc[keep], path.name, path)
