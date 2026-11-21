"""Reopen saved business reports to verify the fixed-column layout."""

import pandas as pd
import pytest
from click.testing import CliRunner
from openpyxl import load_workbook

from report_generator import v2
from report_generator.app_config import CLASS_COL
from report_generator.cli import cli
from report_generator.config import Config
from report_generator.core.reader import MetricsReader
from report_generator.reports.business.builder import BusinessReportBuilder

HEADERS = [
    'Класс',
    'tp',
    'fp',
    'fn',
    'Порог',
    'precision',
    'recall',
    'f1_score',
    'Перебраковка, %',
    'Недобраковка, %',
    'Перебраковка в целях',
    'Недобраковка в целях',
    'Перебраковка грубо не в целях',
    'Недобраковка грубо не в целях',
    'Число примеров test',
    'Число примеров train',
    'Число примеров val',
]


def dashboard(rows):
    """Build a dashboard with every column digital-metrics writes and extra ones.

    Each row is (class, tp, fp, fn, threshold, test count, train count, val count).
    Ratios follow from the counts, so expected values can be worked out by hand.
    """
    frame = pd.DataFrame(
        rows, columns=[CLASS_COL, 'tp', 'fp', 'fn', 'confidence', 'test', 'train', 'val']
    )
    precision = frame['tp'] / (frame['tp'] + frame['fp'])
    recall = frame['tp'] / (frame['tp'] + frame['fn'])
    return pd.DataFrame(
        {
            CLASS_COL: frame[CLASS_COL],
            'counts_observed': True,
            'tp': frame['tp'],
            'fp': frame['fp'],
            'fn': frame['fn'],
            'confidence': frame['confidence'],
            'ap50': 0.5,
            'cohen_kappa': 0.4,
            'precision': precision,
            'recall': recall,
            'f1_score': 2 * precision * recall / (precision + recall),
            'perebrak': 1 - precision,
            'nedobrak': 1 - recall,
            'precision_ci_lower': 0.1,
            'Количество примеров train': frame['train'],
            'Количество примеров test': frame['test'],
            'Количество примеров val': frame['val'],
        }
    )


NEW = [
    ('shared', 8, 2, 2, 0.5, 10, 40, 12),
    ('weak', 5, 5, 15, 0.3, 20, 30, 8),
    ('new', 9, 1, 1, 0.7, 10, 25, 5),
]
PROD = [
    ('shared', 6, 4, 4, 0.4, 10, 40, 12),
    ('weak', 10, 10, 10, 0.2, 20, 50, 8),
    ('deleted', 4, 1, 6, 0.6, 10, 60, 9),
]


def build(tmp_path, new, prod, config=None):
    config = config or Config.default()
    paths = [tmp_path / 'new.xlsx', tmp_path / 'prod.xlsx']
    for data, path in zip([new, prod], paths):
        data.to_excel(path, index=False)
    output = tmp_path / 'business.xlsx'
    BusinessReportBuilder(*(MetricsReader(path) for path in paths), config).build(output)
    return load_workbook(output), config


def metric_sheets(workbook, config):
    names = config.sheet_names
    return [workbook[name] for name in (names.model1, names.model2, names.comparison)]


def headers(sheet):
    return [cell.value for cell in sheet[2]]


def test_metric_sheets_carry_only_the_fixed_columns_in_order(tmp_path):
    workbook, config = build(tmp_path, dashboard(NEW), dashboard(PROD))
    for sheet in metric_sheets(workbook, config):
        assert headers(sheet) == HEADERS
        assert sheet.max_column == len(HEADERS)
    text = ' '.join(str(c.value) for ws in workbook for row in ws for c in row if c.value)
    assert 'coverage' not in text
    assert 'Valid class counts' not in text
    assert 'valid-class-count' not in text


@pytest.mark.parametrize('missing', ['tp', 'confidence', 'precision', 'Количество примеров val'])
def test_listed_column_missing_from_both_inputs_is_omitted(tmp_path, missing):
    new, prod = dashboard(NEW).drop(columns=missing), dashboard(PROD).drop(columns=missing)
    workbook, config = build(tmp_path, new, prod)
    display = config.business.column_translations.get(missing, missing)
    for sheet in metric_sheets(workbook, config):
        assert headers(sheet) == [header for header in HEADERS if header != display]


def test_listed_column_present_in_one_input_is_kept(tmp_path):
    workbook, config = build(tmp_path, dashboard(NEW), dashboard(PROD).drop(columns='tp'))
    new_sheet, prod_sheet, comparison = metric_sheets(workbook, config)
    assert headers(new_sheet) == headers(prod_sheet) == headers(comparison) == HEADERS
    # Rows are the sorted union: deleted, new, shared, weak.
    assert [new_sheet.cell(row, 1).value for row in range(3, 7)] == [
        'deleted',
        'new',
        'shared',
        'weak',
    ]
    assert new_sheet['B5'].value == 8
    assert prod_sheet['B5'].value == 'NA'
    assert comparison['B5'].value == 'NA'


def rows(sheet):
    """Map class name to {header: cell} for the metric table of a sheet."""
    names = headers(sheet)
    table = {}
    for row in sheet.iter_rows(min_row=3):
        if row[0].value is None:
            break
        table[row[0].value] = dict(zip(names, row))
    return table


def test_unsplit_count_stands_in_when_no_input_has_split_counts(tmp_path):
    split = ['Количество примеров train', 'Количество примеров test', 'Количество примеров val']
    new, prod = dashboard(NEW).drop(columns=split), dashboard(PROD).drop(columns=split)
    new['Количество примеров'] = [100, 200, 300]
    prod['Количество примеров'] = [110, 190, 50]
    with pytest.warns(UserWarning, match='training filter disabled'):
        workbook, config = build(tmp_path, new, prod)
    new_sheet, prod_sheet, comparison = metric_sheets(workbook, config)
    for sheet in (new_sheet, prod_sheet, comparison):
        assert headers(sheet) == [*HEADERS[:14], 'Количество примеров']
    cell = rows(new_sheet)['shared']['Количество примеров']
    assert (cell.value, cell.number_format) == (100, '0')
    assert rows(new_sheet)['Среднее']['Количество примеров'].value == 600
    assert rows(comparison)['shared']['Количество примеров'].value == -10


def test_unsplit_count_is_not_shown_beside_a_split_count(tmp_path):
    new, prod = dashboard(NEW), dashboard(PROD)
    new['Количество примеров'] = [100, 200, 300]
    workbook, config = build(tmp_path, new, prod)
    for sheet in metric_sheets(workbook, config):
        assert headers(sheet) == HEADERS


def test_custom_translations_rename_the_remaining_headers(tmp_path):
    settings = tmp_path / 'config.yaml'
    settings.write_text(
        'business:\n  column_translations:\n    tp: Верно\n    f1_score: F1\n'
        '    Количество примеров train: Обучение\n    ap50: AP\n',
        encoding='utf-8',
    )
    workbook, config = build(tmp_path, dashboard(NEW), dashboard(PROD), Config.load(settings))
    expected = list(HEADERS)
    for old, new in [('tp', 'Верно'), ('f1_score', 'F1'), ('Число примеров train', 'Обучение')]:
        expected[expected.index(old)] = new
    new_sheet, _, comparison = metric_sheets(workbook, config)
    for sheet in metric_sheets(workbook, config):
        assert headers(sheet) == expected
    assert rows(new_sheet)['shared']['Верно'].number_format == '0'
    assert rows(new_sheet)['Среднее']['Обучение'].value == 95
    # f1 of "shared": new 0.8, prod 0.6; the renamed column keeps its colour rule.
    cell = rows(comparison)['shared']['F1']
    assert cell.value == pytest.approx(0.2)
    assert cell.fill.fgColor.rgb.endswith(config.colors.positive)


def values(row, names):
    return [row[name].value for name in names]


def test_mean_row_is_na_for_excluded_columns_sum_for_counts_and_mean_for_the_rest(tmp_path):
    workbook, config = build(tmp_path, dashboard(NEW), dashboard(PROD))
    mean = rows(workbook[config.sheet_names.model1])['Среднее']
    for name in ['tp', 'fp', 'fn', 'Порог']:
        assert mean[name].value == 'NA'
        assert mean[name].fill.patternType is None
    # Classes of the new model: precision 0.8, 0.5, 0.9 and recall 0.8, 0.25, 0.9.
    assert mean['precision'].value == pytest.approx(2.2 / 3)
    assert mean['recall'].value == pytest.approx(1.95 / 3)
    assert mean['Перебраковка, %'].value == pytest.approx(80 / 3)
    assert mean['Недобраковка, %'].value == pytest.approx(35)
    assert values(mean, HEADERS[10:14]) == pytest.approx([2 / 3, 2 / 3, 0, 1 / 3])
    counts = ['Число примеров test', 'Число примеров train', 'Число примеров val']
    assert values(mean, counts) == [40, 95, 25]
    for name in counts:
        assert mean[name].number_format == '0'
        assert mean[name].fill.fgColor.rgb.endswith(config.colors.mean)
    prod_mean = rows(workbook[config.sheet_names.model2])['Среднее']
    assert values(prod_mean, counts) == [40, 150, 29]


def test_mean_row_counts_cover_only_the_models_own_eligible_classes(tmp_path):
    new = dashboard([*NEW, ('rare', 1, 1, 1, 0.5, 7, 5, 3)])
    workbook, config = build(tmp_path, new, dashboard([*PROD, ('rare', 2, 1, 1, 0.5, 7, 80, 3)]))
    new_rows = rows(workbook[config.sheet_names.model1])
    prod_rows = rows(workbook[config.sheet_names.model2])
    counts = ['Число примеров test', 'Число примеров train', 'Число примеров val']
    assert new_rows['rare']['tp'].value == 'NA'
    assert values(new_rows['Среднее'], counts) == [40, 95, 25]
    assert values(prod_rows['Среднее'], counts) == [47, 230, 32]


def test_mean_row_follows_a_changed_excluded_from_mean_list(tmp_path):
    settings = tmp_path / 'config.yaml'
    settings.write_text(
        'excluded_from_mean: [precision, Количество примеров train]\n', encoding='utf-8'
    )
    workbook, config = build(tmp_path, dashboard(NEW), dashboard(PROD), Config.load(settings))
    mean = rows(workbook[config.sheet_names.model1])['Среднее']
    assert mean['precision'].value == 'NA'
    assert mean['tp'].value == pytest.approx(22 / 3)
    assert mean['Порог'].value == pytest.approx(0.5)
    # Example counts are summed whatever the list says.
    assert mean['Число примеров train'].value == 95
    assert mean['Число примеров test'].value == 40


def test_mean_row_count_is_na_without_an_eligible_class(tmp_path):
    rare = [('rare', 1, 1, 1, 0.5, 7, 5, 3)]
    workbook, config = build(tmp_path, dashboard(rare), dashboard(PROD))
    mean = rows(workbook[config.sheet_names.model1])['Среднее']
    assert mean['Число примеров train'].value == 'NA'
    assert mean['precision'].value == 'NA'


def test_comparison_is_new_minus_prod_for_every_column(tmp_path):
    workbook, config = build(tmp_path, dashboard(NEW), dashboard(PROD))
    comparison = rows(workbook[config.sheet_names.comparison])
    shared = comparison['shared']
    # shared: new tp/fp/fn 8/2/2 at 0.5, prod 6/4/4 at 0.4; counts are equal.
    assert values(shared, ['tp', 'fp', 'fn']) == [2, -2, -2]
    assert shared['Порог'].value == pytest.approx(0.1)
    assert shared['precision'].value == pytest.approx(0.2)
    assert shared['recall'].value == pytest.approx(0.2)
    assert shared['Перебраковка, %'].value == pytest.approx(-20)
    assert shared['Недобраковка, %'].value == pytest.approx(-20)
    # Goals of shared: new 1, 1, 0, 0; prod (perebrak 0.4, nedobrak 0.4) 0, 0, 0, 0.
    assert values(shared, HEADERS[10:14]) == [1, 1, 0, 0]
    assert values(shared, HEADERS[14:]) == [0, 0, 0]
    # weak: train count 30 against 50.
    assert comparison['weak']['Число примеров train'].value == -20
    assert comparison['weak']['tp'].value == -5
    formats = {name: shared[name].number_format for name in HEADERS[1:]}
    assert formats == {
        **dict.fromkeys(['tp', 'fp', 'fn', *HEADERS[14:]], '0'),
        **dict.fromkeys(['Порог', 'precision', 'recall', 'f1_score'], '0.0000'),
        **dict.fromkeys(HEADERS[8:14], '0.00'),
    }
    for sheet in workbook:
        for row in sheet:
            for cell in row:
                assert cell.data_type != 'f', cell.coordinate


def test_comparison_gives_na_for_a_missing_operand(tmp_path):
    prod = dashboard(PROD)
    prod.loc[prod[CLASS_COL] == 'shared', ['tp', 'confidence']] = float('nan')
    workbook, config = build(tmp_path, dashboard(NEW), prod)
    comparison = rows(workbook[config.sheet_names.comparison])
    for name in ['new', 'deleted']:
        for header in HEADERS[1:]:
            assert comparison[name][header].value == 'NA', (name, header)
            assert comparison[name][header].fill.patternType is None
    assert comparison['shared']['tp'].value == 'NA'
    assert comparison['shared']['Порог'].value == 'NA'
    assert comparison['shared']['fp'].value == -2


def test_comparison_mean_row_aggregates_differences_over_shared_classes(tmp_path):
    workbook, config = build(tmp_path, dashboard(NEW), dashboard(PROD))
    comparison = rows(workbook[config.sheet_names.comparison])
    mean = comparison['Среднее']
    for name in ['tp', 'fp', 'fn', 'Порог']:
        assert mean[name].value == 'NA'
    # Shared classes are "shared" and "weak": recall differences 0.2 and -0.25.
    # The difference of the model means would be 0.65 - 0.5 = 0.15 instead.
    assert mean['recall'].value == pytest.approx(-0.025)
    assert mean['precision'].value == pytest.approx(0.1)
    assert mean['Перебраковка, %'].value == pytest.approx(-10)
    # Goal differences: shared 1, 1, 0, 0 and weak 0, 0, 0, 0.
    assert values(mean, HEADERS[10:14]) == pytest.approx([0.5, 0.5, 0, 0])
    # Counts are summed: test 0 + 0, train 0 - 20, val 0 + 0.
    assert values(mean, HEADERS[14:]) == [0, -20, 0]


def test_comparison_colours_only_quality_and_percentage_columns(tmp_path):
    workbook, config = build(tmp_path, dashboard(NEW), dashboard(PROD))
    comparison = rows(workbook[config.sheet_names.comparison])
    coloured = ['precision', 'recall', 'f1_score', 'Перебраковка, %', 'Недобраковка, %']
    for name in ['shared', 'weak']:
        for header in HEADERS[1:]:
            cell = comparison[name][header]
            if header not in coloured:
                assert cell.fill.patternType is None, (name, header)
    shared = comparison['shared']
    assert shared['precision'].fill.fgColor.rgb.endswith(config.colors.positive)
    assert shared['Перебраковка, %'].fill.fgColor.rgb.endswith(config.colors.positive)
    # weak: recall 0.25 against 0.5, nedobrak 75 % against 50 %.
    weak = comparison['weak']
    assert weak['recall'].fill.fgColor.rgb.endswith(config.colors.negative)
    assert weak['Недобраковка, %'].fill.fgColor.rgb.endswith(config.colors.negative)
    assert weak['tp'].value == -5


NOTE = [
    'Средние и вердикт: собственный набор допустимых классов каждой модели (новая: 3, прод: 3);',
    'среднее разниц: общих классов 2.',
    'Наборы могут различаться;',
    'NA — отсутствующие или отфильтрованные данные.',
    'Причины отсутствия сравнения приведены на листе исключений.',
]


def column_a(sheet, first, last):
    return [sheet.cell(row, 1).value for row in range(first, last + 1)]


def test_population_note_is_five_rows_in_column_a_on_four_sheets(tmp_path):
    workbook, config = build(tmp_path, dashboard(NEW), dashboard(PROD))
    names = config.sheet_names
    # Metric tables end with the mean row 7; the comparison legend takes rows 9-13.
    for title, first in [
        (names.model1, 9),
        (names.model2, 9),
        (names.comparison, 15),
        (names.verdict, 13),
    ]:
        sheet = workbook[title]
        assert column_a(sheet, first - 1, first + 4) == [None, *NOTE], title
        assert sheet.max_row == first + 4
        for row in range(first, first + 5):
            assert all(cell.value is None for cell in sheet[row][1:]), (title, row)
    excluded = ' '.join(str(c.value) for row in workbook[names.excluded] for c in row)
    assert 'собственный набор' not in excluded


def test_comparison_legend_is_four_text_lines_in_column_a(tmp_path):
    workbook, config = build(tmp_path, dashboard(NEW), dashboard(PROD))
    sheet = workbook[config.sheet_names.comparison]
    assert column_a(sheet, 8, 14) == [
        None,
        'Легенда:',
        'Улучшение',
        'Ухудшение: доли > 0.05; перебраковка/недобраковка > 5.0 п.п.',
        'Незначительное изменение',
        'NA — разница недоступна; цветовая оценка не применяется',
        None,
    ]
    assert sheet['A10'].fill.fgColor.rgb.endswith(config.colors.positive)
    assert sheet['A11'].fill.fgColor.rgb.endswith(config.colors.negative)
    assert sheet['A12'].fill.patternType is None
    assert sheet['A13'].fill.patternType is None
    for row in range(9, 14):
        for cell in sheet[row][1:]:
            assert cell.value is None and cell.fill.patternType is None, cell.coordinate


def assert_same_cells(actual, expected):
    """Compare values, formats and fills; numbers may differ by float rounding."""
    assert actual.sheetnames == expected.sheetnames
    for one, two in zip(actual, expected):
        assert one.dimensions == two.dimensions, one.title
        for row_one, row_two in zip(one, two):
            for a, b in zip(row_one, row_two):
                where = f'{one.title}!{a.coordinate}'
                if isinstance(b.value, (int, float)):
                    assert a.value == pytest.approx(b.value, rel=1e-12, abs=1e-12), where
                else:
                    assert a.value == b.value, where
                assert a.number_format == b.number_format, where
                assert a.fill.fgColor.rgb == b.fill.fgColor.rgb, where


def test_structured_builder_writes_the_same_workbook_from_tables(tmp_path):
    legacy, config = build(tmp_path, dashboard(NEW), dashboard(PROD))
    tables = [
        v2.MetricsTable(dashboard(NEW), 'new.xlsx'),
        v2.MetricsTable(dashboard(PROD), 'prod.xlsx'),
    ]
    before = [table.frame.copy(deep=True) for table in tables]
    output = tmp_path / 'structured.xlsx'
    v2.BusinessReportBuilder(*tables, config).build(output)
    # The legacy inputs went through an Excel file, the tables did not.
    assert_same_cells(load_workbook(output), legacy)
    for table, frame in zip(tables, before):
        pd.testing.assert_frame_equal(table.frame, frame)


def test_structured_builder_takes_no_presentation():
    table = v2.MetricsTable(dashboard(NEW), 'new.xlsx')
    with pytest.raises(TypeError, match='presentation'):
        v2.BusinessReportBuilder(
            table,
            table,
            Config.default(),
            presentation=v2.PresentationOptions(counts='below_table'),
        )


def test_structured_builder_never_replaces_an_input_workbook(tmp_path):
    source = tmp_path / 'source.xlsx'
    dashboard(NEW).to_excel(source, index=False)
    table = v2.read_metrics_table(source)
    builder = v2.BusinessReportBuilder(table, table, Config.default())
    link = tmp_path / 'link.xlsx'
    link.symlink_to(source)
    for target in (source, link):
        with pytest.raises(ValueError, match='alias'):
            builder.build(target, overwrite=True)
    output = tmp_path / 'output.xlsx'
    builder.build(output)
    with pytest.raises(FileExistsError):
        builder.build(output)
    builder.build(output, overwrite=True)


def test_command_line_writes_the_fixed_column_layout(tmp_path):
    paths = [tmp_path / 'new.xlsx', tmp_path / 'prod.xlsx']
    for data, path in zip([dashboard(NEW), dashboard(PROD)], paths):
        data.to_excel(path, index=False)
    output = tmp_path / 'out.xlsx'
    result = CliRunner().invoke(cli, ['business', *map(str, paths), str(output)])
    assert result.exit_code == 0, result.output
    workbook = load_workbook(output)
    config = Config.default()
    for sheet in metric_sheets(workbook, config):
        assert headers(sheet) == HEADERS
    assert rows(workbook[config.sheet_names.model1])['Среднее']['tp'].value == 'NA'
