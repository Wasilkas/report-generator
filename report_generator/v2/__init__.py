"""Additive structured reporting API; legacy modules remain supported unchanged."""

from .comparison import write_comparison_workbook
from .layout import compact_report_counts
from .metrics import MetricsTable, read_metrics_table
from .reports import BusinessReportBuilder, DevReportBuilder, PresentationOptions
from .workbook import (
    MetadataFormat,
    annotate_workbook,
    original_workbook,
    package_fingerprint,
    read_metadata,
)

__all__ = [
    'MetricsTable',
    'read_metrics_table',
    'PresentationOptions',
    'DevReportBuilder',
    'BusinessReportBuilder',
    'write_comparison_workbook',
    'compact_report_counts',
    'MetadataFormat',
    'annotate_workbook',
    'original_workbook',
    'read_metadata',
    'package_fingerprint',
]
