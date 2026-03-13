"""MeanRowCalculator, ComparisonCalculator, ClassFilter."""

from __future__ import annotations

import pandas as pd

from ..app_config import CLASS_COL, REASON_COL, TRAIN_COUNT_DISPLAY_COL
from ..config import Config


def _find_train_col(df: pd.DataFrame) -> str | None:
    """Return the column that holds the training-set example count, or None."""
    for col in df.columns:
        s = str(col).lower()
        if 'train' in s and ('примеров' in s or 'пример' in s):
            return col
    return None


class ClassFilter:
    """Splits a DataFrame into included and excluded classes.

    Classes with training-example count <= config.min_train_count are excluded.
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
            return df, pd.DataFrame(columns=[CLASS_COL, TRAIN_COUNT_DISPLAY_COL, REASON_COL])

        counts = pd.to_numeric(df[train_col], errors='coerce').fillna(0)
        mask = counts > self._min_train_count

        included = df[mask].reset_index(drop=True)
        excluded_raw = df[~mask].reset_index(drop=True)

        excluded = excluded_raw[[CLASS_COL, train_col]].rename(
            columns={train_col: TRAIN_COUNT_DISPLAY_COL}
        )
        excluded[REASON_COL] = f'Мало примеров train (≤ {self._min_train_count})'
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
    ClassFilter.  Columns listed in config.excluded_from_mean are left blank.
    """

    def __init__(self, config: Config) -> None:
        self._excluded_from_mean = config.excluded_from_mean

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
            numeric = pd.to_numeric(df[col], errors='coerce')
            row[col] = numeric.mean() if numeric.notna().any() else None

        return pd.Series(row)


class ComparisonCalculator:
    """Computes metric differences between model1 and model2 (model1 − model2).

    Columns in config.excluded_from_mean (counts, tp/fp/fn, confidence) are
    kept from model1 unchanged so that the 'Среднее' filter still works.
    Only classes present in both files are included.
    """

    def __init__(self, config: Config) -> None:
        self._excluded_from_mean = config.excluded_from_mean

    def compute(self, df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
        idx1 = df1.set_index(CLASS_COL)
        idx2 = df2.set_index(CLASS_COL)

        common = idx1.index.intersection(idx2.index)
        idx1 = idx1.loc[common].copy()
        idx2 = idx2.loc[common]

        for col in idx1.columns:
            if col in self._excluded_from_mean:
                continue
            if col not in idx2.columns:
                continue
            v1 = pd.to_numeric(idx1[col], errors='coerce')
            v2 = pd.to_numeric(idx2[col], errors='coerce')
            if v1.notna().any() and v2.notna().any():
                idx1[col] = v1 - v2

        return idx1.reset_index()
