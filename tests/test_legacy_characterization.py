"""Capture saved legacy artifacts and public call surfaces before extraction."""

import dataclasses
import inspect
import json
from pathlib import Path

import pandas as pd
import pytest
from click.testing import CliRunner
from openpyxl import load_workbook

from report_generator.cli import cli
from report_generator.config import BusinessConfig, ColorsConfig, Config, SheetNamesConfig
from report_generator.core.reader import MetricsReader
from report_generator.core.writer import ExcelSheetWriter
from report_generator.reports.base import BaseReportBuilder
from report_generator.reports.business.builder import BusinessReportBuilder
from report_generator.reports.dev.builder import DevReportBuilder

FIXTURE = Path(__file__).parent / 'fixtures/legacy-characterization.json'
BUSINESS_FIXTURE = Path(__file__).parent / 'fixtures/business-verdict-excluded.json'
SCENARIOS = ['shared', 'disjoint', 'asymmetric', 'empty']


def scenario_sources(tmp_path, scenario):
    """Write the new and prod input workbooks of one characterization scenario."""
    sources = []
    for role in ['new', 'prod']:
        names = ['shared', role, 'filtered', 'emptyAP']
        if scenario == 'disjoint':
            names = [role]
        scores = [0.8 if role == 'new' else 0.5] * len(names)
        counts = [30] * len(names)
        if scenario == 'asymmetric' and role == 'new':
            counts[2] = 0
        if scenario == 'empty':
            counts = [0] * len(names)
        frame = pd.DataFrame(
            {
                'Класс': names,
                'Количество примеров train': counts,
                'f1_score': scores,
                'ap50': scores,
                'perebrak': [0.1] * len(names),
                'nedobrak': [0.2] * len(names),
                'confidence': [0.5] * len(names),
            }
        )
        frame.loc[frame['Класс'] == 'emptyAP', 'ap50'] = float('nan')
        path = tmp_path / (role + '.xlsx')
        frame.to_excel(path, index=False)
        sources.append(path)
    return sources


def characterize(tmp_path):
    result = {'signatures': {}, 'cli': {}, 'workbooks': {}}
    for cls in [
        MetricsReader,
        ExcelSheetWriter,
        BaseReportBuilder,
        DevReportBuilder,
        BusinessReportBuilder,
        Config,
        ColorsConfig,
        SheetNamesConfig,
        BusinessConfig,
    ]:
        result['signatures'][cls.__name__] = str(inspect.signature(cls))
        for method in ['read', 'build', 'write', 'load', 'default', 'with_comparison_colors']:
            if hasattr(cls, method):
                result['signatures'][cls.__name__ + '.' + method] = str(
                    inspect.signature(getattr(cls, method))
                )

    def serial(value):
        if isinstance(value, dict):
            return {k: serial(v) for k, v in value.items()}
        if isinstance(value, (set, frozenset)):
            return sorted(value)
        return value

    result['defaults'] = serial(dataclasses.asdict(Config.load()))
    for args in [[], ['dev'], ['business']]:
        invocation = CliRunner().invoke(cli, [*args, '--help'])
        result['cli'][' '.join(args) or 'root'] = invocation.output
    for scenario in SCENARIOS:
        sources = scenario_sources(tmp_path, scenario)
        # The business workbook has its own golden; only the developer one is frozen here.
        for builder in [DevReportBuilder]:
            target = tmp_path / 'report.xlsx'
            builder(*(MetricsReader(p) for p in sources), Config.load()).build(
                target, overwrite=True
            )
            workbook = load_workbook(target)
            sheets = {}
            for sheet in workbook:
                sheets[sheet.title] = {
                    'cells': [
                        [
                            cell.coordinate,
                            cell.value,
                            cell.data_type,
                            cell.number_format,
                            cell.style_id,
                        ]
                        for row in sheet
                        for cell in row
                        if cell.value is not None
                    ],
                    'styles': [str(style) for style in workbook._cell_styles],
                    'widths': {col: dim.width for col, dim in sheet.column_dimensions.items()},
                    'heights': {row: dim.height for row, dim in sheet.row_dimensions.items()},
                    'freeze': sheet.freeze_panes,
                    'merged': str(sheet.merged_cells),
                }
            result['workbooks'][scenario + ':' + builder.__name__] = sheets
            workbook.close()
    reader_path = tmp_path / 'reader.xlsx'
    pd.DataFrame(
        {'label': [' cat ', '0001', '1e-3', 0.5, 'Среднее', None, 'dog'], 'metric': list(range(7))}
    ).to_excel(reader_path, index=False)
    result['reader'] = MetricsReader(reader_path).read().to_dict(orient='list')
    return json.loads(json.dumps(result, ensure_ascii=False, default=str))


def test_legacy_characterization(tmp_path):
    assert characterize(tmp_path) == json.loads(FIXTURE.read_text())


@pytest.mark.parametrize('scenario', SCENARIOS)
def test_business_verdict_and_excluded_sheets_equal_the_previous_builder(tmp_path, scenario):
    """The fixed-column layout leaves the verdict and the excluded classes as they were."""
    sources = scenario_sources(tmp_path, scenario)
    target = tmp_path / 'report.xlsx'
    BusinessReportBuilder(*(MetricsReader(p) for p in sources), Config.load()).build(target)
    workbook = load_workbook(target)
    expected = json.loads(BUSINESS_FIXTURE.read_text(encoding='utf-8'))['workbooks'][scenario]
    for title, frozen in expected.items():
        sheet = workbook[title]
        # Below the verdict legend (row 11) only the population note follows.
        last_row = 11 if title == 'Вердикт' else sheet.max_row
        cells = [
            [cell.coordinate, cell.value, cell.data_type, cell.number_format]
            for row in sheet.iter_rows(max_row=last_row)
            for cell in row
            if cell.value is not None
        ]
        assert cells == frozen['cells'], title
        assert {col: dim.width for col, dim in sheet.column_dimensions.items()} == frozen[
            'widths'
        ], title
        assert sheet.freeze_panes == frozen['freeze'], title
        assert str(sheet.merged_cells) == frozen['merged'], title
    workbook.close()
