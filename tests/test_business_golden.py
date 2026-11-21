"""Golden: the business workbook matches the report owner's template cell by cell.

The fixture holds the template as it was supplied (anonymized; see its provenance
block) and the two input tables rebuilt from its model sheets. The test builds the
report from those inputs and compares every cell's value, number format and fill.

The template was a report edited by hand, so it differs from the decided layout in
a few places. Each difference is one entry of EXCEPTIONS below, tied to the layout
decision that settles it; nothing else may differ. Not compared, because the
template's values were set by hand in Excel and the decision is that the writer
keeps its own: column widths, freeze panes, fonts and alignments.
"""

import json
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

from report_generator.config import Config
from report_generator.core.reader import MetricsReader
from report_generator.reports.business.builder import BusinessReportBuilder

FIXTURE = Path(__file__).parent / 'fixtures/business-golden.json'
METRIC_SHEETS = ('Новая модель', 'Прод модель', 'Сравнение')
MEAN_ROW = 40  # 37 classes below the title and header rows
NOTE = [
    'Средние и вердикт: собственный набор допустимых классов каждой модели (новая: 37, прод: 37);',
    'среднее разниц: общих классов 37.',
    'Наборы могут различаться;',
    'NA — отсутствующие или отфильтрованные данные.',
    'Причины отсутствия сравнения приведены на листе исключений.',
]
NOTE_START = {'Новая модель': 42, 'Прод модель': 42, 'Сравнение': 48, 'Вердикт': 13}
LEGEND = [
    ('Улучшение', 'C6EFCE'),
    ('Ухудшение: доли > 0.05; перебраковка/недобраковка > 5.0 п.п.', 'FFC7CE'),
    ('Незначительное изменение', None),
    ('NA — разница недоступна; цветовая оценка не применяется', None),
]
MEAN_FILL = 'FFFFCC'


def _column(coordinate):
    return column_index_from_string(coordinate_from_string(coordinate)[0])


def _row(coordinate):
    return coordinate_from_string(coordinate)[1]


def _no_availability_columns(sheets, formulas):
    """Decision "Availability: none".

    The template kept the old coverage columns R:AA on the new-model sheet (most
    of them hidden); the layout ends with the example counts in column Q.
    """
    sheet = sheets['Новая модель']
    for coordinate in [c for c in sheet if _column(c) > _column('Q1')]:
        del sheet[coordinate]


def _mean_row_na_for_excluded_columns(sheets, formulas):
    """Decision "Среднее row: columns named in excluded_from_mean are NA".

    The template has hand-typed AVERAGE formulas (model sheets) and differences of
    them (comparison) for tp, fp, fn and the threshold.
    """
    for title in METRIC_SHEETS:
        for letter in 'BCDE':
            assert f'{letter}{MEAN_ROW}' in formulas[title]
            sheets[title][f'{letter}{MEAN_ROW}'] = ('NA', 'General', None)


def _computed_values_instead_of_formulas(sheets, formulas):
    """Decision "Cells: computed numbers, never Excel formulas".

    The fixture already stores the value Excel cached for each formula, so values
    compare directly. A computed number in the mean row is styled like every other
    mean value, with the mean fill; the template's hand-typed formula cells there
    (example counts on the model sheets, goal and count columns on the comparison)
    had no fill.
    """
    for title in METRIC_SHEETS:
        for coordinate in formulas[title]:
            value, number_format, _ = sheets[title][coordinate]
            if _row(coordinate) == MEAN_ROW and value != 'NA':
                sheets[title][coordinate] = (value, number_format, MEAN_FILL)


def _one_number_format_per_column(sheets, formulas):
    """Decision "Formats: one number format per column on all three metric sheets".

    The template's comparison formulas showed the threshold and the example counts
    with two decimals; they take the formats of the model sheets.
    """
    sheet = sheets['Сравнение']
    for coordinate, (value, _, fill) in list(sheet.items()):
        if 3 <= _row(coordinate) <= MEAN_ROW and isinstance(value, (int, float)):
            letter = get_column_letter(_column(coordinate))
            if letter == 'E':
                sheet[coordinate] = (value, '0.0000', fill)
            elif letter in 'OPQ':
                sheet[coordinate] = (value, '0', fill)


def _legend_text_on_the_coloured_cell(sheets, formulas):
    """Decision "Сравнение legend: four text lines, text in column A on the coloured cell".

    The template lost the legend text and painted A:E of the two coloured rows.
    """
    sheet = sheets['Сравнение']
    for row, (text, fill) in enumerate(LEGEND, 43):
        for letter in 'BCDE':
            sheet.pop(f'{letter}{row}', None)
        sheet[f'A{row}'] = (text, 'General', fill)


def _full_note_wording(sheets, formulas):
    """Decision "Population note: five rows in column A with the full original wording".

    The template's first row lost its prefix and two rows lost their semicolons.
    """
    for title, start in NOTE_START.items():
        for row, line in enumerate(NOTE, start):
            assert f'A{row}' in sheets[title]
            sheets[title][f'A{row}'] = (line, 'General', None)


EXCEPTIONS = [
    _no_availability_columns,
    _mean_row_na_for_excluded_columns,
    _computed_values_instead_of_formulas,
    _one_number_format_per_column,
    _legend_text_on_the_coloured_cell,
    _full_note_wording,
]


def _expected(golden):
    sheets = {
        title: {coordinate: tuple(cell) for coordinate, *cell, _ in cells}
        for title, cells in golden['sheets'].items()
    }
    formulas = {
        title: {coordinate for coordinate, *_, formula in cells if formula}
        for title, cells in golden['sheets'].items()
    }
    for exception in EXCEPTIONS:
        exception(sheets, formulas)
    return sheets


def _actual(path):
    workbook = load_workbook(path)
    sheets = {}
    for sheet in workbook:
        cells = {}
        for row in sheet:
            for cell in row:
                assert cell.data_type != 'f', f'{sheet.title}!{cell.coordinate} is a formula'
                fill = cell.fill.fgColor.rgb[2:] if cell.fill.patternType else None
                if cell.value is not None or fill is not None:
                    cells[cell.coordinate] = (cell.value, cell.number_format, fill)
        sheets[sheet.title] = cells
    return sheets


def test_generated_workbook_matches_the_template(tmp_path):
    golden = json.loads(FIXTURE.read_text(encoding='utf-8'))
    readers = []
    for role in ('new', 'prod'):
        table = golden['inputs'][role]
        # The title rows name the input files, so the inputs keep the template's names.
        path = tmp_path / table['file']
        pd.DataFrame(table['rows'], columns=table['columns']).to_excel(path, index=False)
        readers.append(MetricsReader(path))
    output = tmp_path / 'business.xlsx'
    BusinessReportBuilder(*readers, Config.default()).build(output)

    expected, actual = _expected(golden), _actual(output)
    assert list(actual) == list(expected)
    for title, cells in expected.items():
        assert sorted(actual[title]) == sorted(cells), title
        for coordinate, (value, number_format, fill) in cells.items():
            where = f'{title}!{coordinate}'
            got_value, got_format, got_fill = actual[title][coordinate]
            if isinstance(value, (int, float)):
                assert got_value == pytest.approx(value, rel=1e-12, abs=1e-12), where
            else:
                assert got_value == value, where
            assert got_format == number_format, where
            assert got_fill == fill, where
