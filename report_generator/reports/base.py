"""BaseReportBuilder: abstract base class for all report builders."""

from abc import ABC, abstractmethod
from pathlib import Path

import pandas as pd

from ..app_config import CLASS_COL, REASON_COL, TRAIN_COUNT_DISPLAY_COL
from ..config import Config
from ..core.calculator import ClassFilter, ComparisonCalculator, MeanRowCalculator, _find_train_col
from ..core.reader import MetricsReader


class BaseReportBuilder(ABC):
    """Abstract base for Excel report builders.

    Subclasses must implement :meth:`build`.
    """

    def __init__(
        self,
        model1_reader: MetricsReader,
        model2_reader: MetricsReader,
        config: Config,
    ) -> None:
        self.model1_reader = model1_reader
        self.model2_reader = model2_reader
        self._config = config
        self._filter = ClassFilter(config)
        self._mean_calc = MeanRowCalculator(config)
        self._cmp_calc = ComparisonCalculator(config)

    def _check_output(self, output_path: Path, overwrite: bool) -> None:
        for reader in (self.model1_reader, self.model2_reader):
            source = reader.file_path
            if output_path.resolve() == source.resolve() or (
                output_path.exists() and source.exists() and output_path.samefile(source)
            ):
                raise ValueError(f'Output must not alias an input workbook: {output_path}')
        if output_path.exists() and not overwrite:
            raise FileExistsError(
                f'Output already exists: {output_path}; use --force to replace it'
            )

    def _prepare(self, raw1: pd.DataFrame, raw2: pd.DataFrame):
        df1, train1 = self._filter.split(raw1)
        df2, train2 = self._filter.split(raw2)
        common = set(df1[CLASS_COL]) & set(df2[CLASS_COL])
        rows = []
        for raw, peer, train, model, unmatched_reason in [
            (raw1, raw2, train1, 'Новая модель', 'Только в новой модели'),
            (raw2, raw1, train2, 'Прод модель', 'Только в прод модели'),
        ]:
            train_col = _find_train_col(raw)
            train_reasons = dict(zip(train[CLASS_COL], train[REASON_COL]))
            peer_classes = set(peer[CLASS_COL])
            for _, row in raw.iterrows():
                cls = row[CLASS_COL]
                if cls in common:
                    continue
                reasons = []
                if cls in train_reasons:
                    reasons.append(train_reasons[cls])
                if cls not in peer_classes:
                    reasons.append(unmatched_reason)
                if not reasons:
                    reasons.append('Исключен из сравнения: мало примеров train в другой модели')
                rows.append(
                    {
                        CLASS_COL: cls,
                        TRAIN_COUNT_DISPLAY_COL: row[train_col] if train_col else None,
                        REASON_COL: '; '.join(reasons),
                        'Модель': model,
                    }
                )
        excluded = pd.DataFrame(
            rows, columns=[CLASS_COL, TRAIN_COUNT_DISPLAY_COL, REASON_COL, 'Модель']
        )
        return (
            df1[df1[CLASS_COL].isin(common)].reset_index(drop=True),
            df2[df2[CLASS_COL].isin(common)].reset_index(drop=True),
            excluded,
        )

    @abstractmethod
    def build(self, output_path: str | Path, *, overwrite: bool = False) -> None: ...
