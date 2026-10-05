"""MeanRowCalculator, ComparisonCalculator, ClassFilter."""

import math
import re
import warnings

import pandas as pd

from ..app_config import CLASS_COL, REASON_COL, TRAIN_COUNT_DISPLAY_COL
from ..config import Config


def _find_train_col(df: pd.DataFrame) -> str | None:
    """Return the column that holds the training-set example count, or None."""
    candidates = [
        col for col in df.columns if 'train' in str(col).lower() and 'пример' in str(col).lower()
    ]
    if len(candidates) > 1:
        raise ValueError(f'Multiple training count columns: {candidates}')
    return candidates[0] if candidates else None


class ClassFilter:
    """Splits a DataFrame into included and excluded classes.

    Classes with training-example count < config.min_train_count are excluded.
    The excluded DataFrame keeps ID (if present), class name, and train count.
    """

    def __init__(self, config: Config) -> None:
        self._min_train_count = config.min_train_count

    def split(self, df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
        """Return (included_df, excluded_summary_df).

        excluded_summary_df columns: CLASS_COL, TRAIN_COUNT_DISPLAY_COL, REASON_COL.
        """
        train_col = _find_train_col(df)
        if train_col is None:
            warnings.warn(
                'No recognized training count column; training filter disabled',
                UserWarning,
                stacklevel=2,
            )
            return df, pd.DataFrame(columns=[CLASS_COL, TRAIN_COUNT_DISPLAY_COL, REASON_COL])

        counts = pd.to_numeric(df[train_col], errors='coerce')
        invalid = (
            ~counts.map(lambda value: pd.notna(value) and math.isfinite(value))
            | (counts < 0)
            | (counts % 1 != 0)
            | df[train_col].map(lambda value: isinstance(value, bool))
        )
        if invalid.any():
            classes = df.loc[invalid, CLASS_COL].tolist()
            raise ValueError(f'Invalid training counts in {train_col!r} for classes: {classes}')
        mask = counts >= self._min_train_count

        included = df[mask].reset_index(drop=True)
        excluded_raw = df[~mask].reset_index(drop=True)

        excluded = excluded_raw[[CLASS_COL, train_col]].rename(
            columns={train_col: TRAIN_COUNT_DISPLAY_COL}
        )
        excluded[REASON_COL] = f'Мало примеров train (< {self._min_train_count})'
        return included, excluded


def exclude_unmatched(
    df1: pd.DataFrame, df2: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Remove classes absent from either model and return an excluded summary.

    Returns (df1_common, df2_common, excluded_unmatched_df) where
    excluded_unmatched_df has columns CLASS_COL, TRAIN_COUNT_DISPLAY_COL, REASON_COL.
    """
    classes1 = set(df1[CLASS_COL])
    classes2 = set(df2[CLASS_COL])
    only_in_1 = classes1 - classes2
    only_in_2 = classes2 - classes1

    train_col1 = _find_train_col(df1)
    train_col2 = _find_train_col(df2)

    rows: list[dict] = []
    for cls in sorted(only_in_1):
        row = df1.loc[df1[CLASS_COL] == cls].iloc[0]
        count = row[train_col1] if train_col1 else None
        rows.append(
            {CLASS_COL: cls, TRAIN_COUNT_DISPLAY_COL: count, REASON_COL: 'Только в новой модели'}
        )  # noqa: E501
    for cls in sorted(only_in_2):
        row = df2.loc[df2[CLASS_COL] == cls].iloc[0]
        count = row[train_col2] if train_col2 else None
        rows.append(
            {CLASS_COL: cls, TRAIN_COUNT_DISPLAY_COL: count, REASON_COL: 'Только в прод модели'}
        )  # noqa: E501

    cols = [CLASS_COL, TRAIN_COUNT_DISPLAY_COL, REASON_COL]
    excluded_unmatched = pd.DataFrame(rows, columns=cols)

    common = classes1 & classes2
    df1_common = df1[df1[CLASS_COL].isin(common)].reset_index(drop=True)
    df2_common = df2[df2[CLASS_COL].isin(common)].reset_index(drop=True)
    return df1_common, df2_common, excluded_unmatched


class MeanRowCalculator:
    """Appends a 'Среднее' row to a DataFrame.

    Assumes low-training-count classes have already been removed by
    ClassFilter. Columns listed in config.excluded_from_mean remain missing internally
    and are displayed as NA.
    """

    def __init__(self, config: Config) -> None:
        self._excluded_from_mean = config.excluded_from_mean
        self._metric_cols = config.ratio_cols | config.better_higher_cols | config.better_lower_cols

    def append_mean_row(self, df: pd.DataFrame) -> pd.DataFrame:
        mean_series = self._compute(df)
        return pd.concat([df, mean_series.to_frame().T], ignore_index=True)

    def _compute(self, df: pd.DataFrame) -> pd.Series:
        row: dict[str, object] = {CLASS_COL: 'Среднее'}

        for col in df.columns:
            if col == CLASS_COL:
                continue
            if col in self._excluded_from_mean:
                row[col] = None
                continue
            numeric = pd.to_numeric(df[col], errors='coerce').replace(
                [math.inf, -math.inf], float('nan')
            )
            finite = numeric.dropna()
            row[col] = float((finite / len(finite)).sum()) if len(finite) else None
            if col in self._metric_cols or numeric.notna().any() or df[col].isna().all():
                coverage = f'{col} coverage'
                if coverage in df.columns:
                    raise ValueError(f'Generated coverage column collision: {coverage}')
                row[coverage] = f'{numeric.notna().sum()}/{len(df)}'

        return pd.Series(row)


class ComparisonCalculator:
    """Computes metric differences between model1 and model2 (model1 − model2).

    Counts, IDs, confidence and nonnumeric columns are omitted. Unavailable
    operands remain unavailable differences. All eligible classes are included;
    missing operands remain numeric NaN.
    """

    def __init__(self, config: Config) -> None:
        self._non_metrics = config.int_cols | {'ID', 'confidence', 'tp', 'fp', 'fn'}
        self._metrics = config.ratio_cols | config.better_higher_cols | config.better_lower_cols

    def compute(self, df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
        idx1 = df1.set_index(CLASS_COL)
        idx2 = df2.set_index(CLASS_COL)

        classes = idx1.index.union(idx2.index, sort=True)
        idx1 = idx1.reindex(classes)
        idx2 = idx2.reindex(classes)

        result = pd.DataFrame(index=classes)
        for col in idx1.columns.union(idx2.columns, sort=False):
            identity = str(col).casefold()
            is_count = (
                'пример' in identity
                or 'count' in identity
                or re.search(r'(^|[_\s])n([_\s]|$)', identity) is not None
            )
            if col in self._non_metrics or is_count:
                continue
            raw1 = idx1[col] if col in idx1 else pd.Series(float('nan'), index=classes)
            raw2 = idx2[col] if col in idx2 else pd.Series(float('nan'), index=classes)
            v1 = pd.to_numeric(raw1, errors='coerce').replace([math.inf, -math.inf], float('nan'))
            v2 = pd.to_numeric(raw2, errors='coerce').replace([math.inf, -math.inf], float('nan'))
            if col in self._metrics or v1.notna().any() or v2.notna().any():
                result[col] = (v1 - v2).replace([math.inf, -math.inf], float('nan'))
        return result.reset_index()
