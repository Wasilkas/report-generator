"""BaseReportBuilder: abstract base class for all report builders."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from ..config import Config
from ..core.calculator import ClassFilter, ComparisonCalculator, MeanRowCalculator
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

    @abstractmethod
    def build(self, output_path: str | Path) -> None: ...
