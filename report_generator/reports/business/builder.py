"""BusinessReportBuilder: assembles the business-oriented Excel report."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from ...config import Config
from ...core.reader import MetricsReader
from ...core.writer import ExcelSheetWriter
from ..base import BaseReportBuilder
from .utils import (
    add_goal_columns,
    build_comparison_color_fn,
    build_format_map,
    to_percentage,
    translate_columns,
    write_excluded_sheet,
    write_verdict_sheet,
)


class BusinessReportBuilder(BaseReportBuilder):
    """Builds the business-oriented Excel report.

    Sheet 1 — new model metrics (Russian headers, % perebrak/nedobrak, goal columns).
    Sheet 2 — prod model metrics, same structure.
    Sheet 3 — Сравнение: numeric diff (new − prod) with comparison coloring.
    Sheet 4 — Вердикт: F1/classes-in-target/gross-not-target verdict table.
    Sheet 5 — Удаленные классы: classes excluded due to insufficient training data.
    """

    def __init__(
        self,
        model1_reader: MetricsReader,
        model2_reader: MetricsReader,
        config: Config,
    ) -> None:
        super().__init__(model1_reader, model2_reader, config)

        fmt = build_format_map(config)
        self._sheet_writer = ExcelSheetWriter(config, format_map=fmt)
        self._diff_writer = ExcelSheetWriter(
            config,
            color_fn=build_comparison_color_fn(config),
            format_map=fmt,
        )

    def build(self, output_path: str | Path) -> None:
        output_path = Path(output_path)
        biz = self._config.business
        translations = dict(biz.column_translations)

        df1_raw = self.model1_reader.read()
        df2_raw = self.model2_reader.read()

        df1, excluded = self._filter.split(df1_raw)
        df2, _ = self._filter.split(df2_raw)

        # Comparison diff (no goal columns — diff of 0/1 flags is meaningless)
        df_diff_display = translate_columns(
            to_percentage(self._mean_calc.append_mean_row(self._cmp_calc.compute(df1, df2)), biz),
            translations,
        )

        # Goal columns on sheets 1 & 2
        df1_goals = add_goal_columns(df1, biz)
        df2_goals = add_goal_columns(df2, biz)

        df1_display = translate_columns(
            to_percentage(self._mean_calc.append_mean_row(df1_goals), biz), translations
        )
        df2_display = translate_columns(
            to_percentage(self._mean_calc.append_mean_row(df2_goals), biz), translations
        )

        wb = Workbook()
        ws1 = wb.active
        ws1.title = self._config.sheet_names.model1
        ws2 = wb.create_sheet(self._config.sheet_names.model2)
        ws3 = wb.create_sheet(self._config.sheet_names.comparison)
        ws_verdict = wb.create_sheet('Вердикт')
        ws_excl = wb.create_sheet(self._config.sheet_names.excluded)

        self._sheet_writer.write(
            ws1,
            df1_display,
            sheet_title=f'Метрики новой модели — {self.model1_reader.file_path.name}',
        )
        self._sheet_writer.write(
            ws2,
            df2_display,
            sheet_title=f'Метрики прод модели — {self.model2_reader.file_path.name}',
        )
        self._diff_writer.write(ws3, df_diff_display, sheet_title='Сравнение (новая − прод)')
        write_verdict_sheet(ws_verdict, df1_goals, df2_goals, self._config)
        write_excluded_sheet(ws_excl, excluded, self._config)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        print(f'Business report saved → {output_path}')
