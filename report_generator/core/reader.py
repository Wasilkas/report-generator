"""MetricsReader: loads raw metrics from an Excel file into a DataFrame."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

CLASS_COL = 'Класс'


class MetricsReader:
    """Reads one sheet of an Excel metrics file into a clean DataFrame.

    The first column (class name, no header in source files) is renamed to
    CLASS_COL. Non-data rows are dropped:
    - empty / NaN class names
    - rows whose class name is a number (threshold rows at the bottom)
    - pre-existing 'Среднее' aggregate rows
    """

    def __init__(self, file_path: str | Path, sheet_name: str | int = 0) -> None:
        self.file_path = Path(file_path)
        self.sheet_name = sheet_name

    def read(self) -> pd.DataFrame:
        df = pd.read_excel(
            self.file_path,
            sheet_name=self.sheet_name,
            header=0,
        )
        # The first column has no header in source files → pandas names it
        # something like "Unnamed: 0"; rename it to a stable internal name.
        first_col = df.columns[0]
        df = df.rename(columns={first_col: CLASS_COL})

        df = df[df[CLASS_COL].apply(_is_valid_class)].reset_index(drop=True)
        return df

    def __repr__(self) -> str:
        return f'MetricsReader({self.file_path.name!r}, sheet={self.sheet_name!r})'


def _is_valid_class(val: object) -> bool:
    """Return True for genuine class-name strings."""
    if val is None:
        return False
    try:
        if pd.isna(val):
            return False
    except (TypeError, ValueError):
        pass
    if isinstance(val, (int, float)):
        return False
    return str(val).strip() not in ('Среднее', '')
