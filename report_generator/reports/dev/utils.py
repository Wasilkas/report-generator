"""Utility functions for the developer Excel report."""

from __future__ import annotations

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from ...config import Config
from ...core.reader import CLASS_COL


def write_excluded_sheet(ws: Worksheet, excluded: pd.DataFrame, config: Config) -> None:
    header_font = Font(bold=True, name='Calibri')
    header_fill = PatternFill('solid', fgColor=config.colors.header)
    center = Alignment(horizontal='center', vertical='center')

    ws.cell(1, 1, 'Удаленные классы (число примеров train ≤ порога)').font = Font(
        bold=True, size=13, name='Calibri'
    )

    display = {CLASS_COL: 'Класс'}
    for c_idx, col in enumerate(excluded.columns, 1):
        cell = ws.cell(2, c_idx, display.get(col, col))
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center

    for r, (_, row) in enumerate(excluded.iterrows(), 3):
        for c_idx, col in enumerate(excluded.columns, 1):
            val = row[col]
            cell = ws.cell(r, c_idx, val)
            cell.font = Font(name='Calibri')
            cell.alignment = center
            if isinstance(val, (int, float)):
                cell.number_format = '0'

    col_widths = {'ID': 8, CLASS_COL: 30}
    for c_idx, col in enumerate(excluded.columns, 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = col_widths.get(col, 20)

    ws.freeze_panes = ws.cell(3, 1)
