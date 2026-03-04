"""CLI entry point for the report generator (powered by Click).

Usage examples
--------------
# Dev report using bundled default config:
    python -m report_generator.cli dev аммп_новая.xlsx аммп_прод.xlsx report_dev.xlsx

# Dev report with a custom config override:
    python -m report_generator.cli --config my_config.yaml dev m1.xlsx m2.xlsx out.xlsx

# Business report:
    python -m report_generator.cli business аммп_новая.xlsx аммп_прод.xlsx report_biz.xlsx

# Explicit sheet names inside the Excel files:
    python -m report_generator.cli dev m1.xlsx m2.xlsx out.xlsx \\
        --sheet1 "Новая" --sheet2 "Sheet1"
"""

from __future__ import annotations

from pathlib import Path

import click

from .config import Config
from .core.reader import MetricsReader
from .reports.business.builder import BusinessReportBuilder
from .reports.dev.builder import DevReportBuilder


def _coerce_sheet(value: str) -> str | int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return value


# ── Shared options ─────────────────────────────────────────────────────────────


def _shared_options(f):
    for decorator in reversed(
        [
            click.argument('model1', type=click.Path(exists=True, dir_okay=False, path_type=Path)),
            click.argument('model2', type=click.Path(exists=True, dir_okay=False, path_type=Path)),
            click.argument('output', type=click.Path(dir_okay=False, path_type=Path)),
            click.option(
                '--sheet1',
                default='0',
                show_default=True,
                help='Sheet name or 0-based index for MODEL1.',
            ),
            click.option(
                '--sheet2',
                default='0',
                show_default=True,
                help='Sheet name or 0-based index for MODEL2.',
            ),
        ]
    ):
        f = decorator(f)
    return f


# ── CLI group ──────────────────────────────────────────────────────────────────


@click.group()
@click.option(
    '--config',
    'config_path',
    default=None,
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    help='Path to a YAML config file. Overrides keys from the default config.yaml.',
)
@click.pass_context
def cli(ctx: click.Context, config_path: Path | None) -> None:
    """Report generator: build Excel reports from two metrics files."""
    ctx.ensure_object(dict)
    ctx.obj['config'] = Config.load(config_path)


@cli.command('dev')
@_shared_options
@click.pass_context
def dev_cmd(
    ctx: click.Context,
    model1: Path,
    model2: Path,
    output: Path,
    sheet1: str,
    sheet2: str,
) -> None:
    """Build a 4-sheet developer report (metrics + comparison + excluded classes)."""
    config: Config = ctx.obj['config']
    reader1 = MetricsReader(model1, sheet_name=_coerce_sheet(sheet1))
    reader2 = MetricsReader(model2, sheet_name=_coerce_sheet(sheet2))
    DevReportBuilder(reader1, reader2, config).build(output)


@cli.command('business')
@_shared_options
@click.pass_context
def business_cmd(
    ctx: click.Context,
    model1: Path,
    model2: Path,
    output: Path,
    sheet1: str,
    sheet2: str,
) -> None:
    """Build a business-oriented report."""
    config: Config = ctx.obj['config']
    reader1 = MetricsReader(model1, sheet_name=_coerce_sheet(sheet1))
    reader2 = MetricsReader(model2, sheet_name=_coerce_sheet(sheet2))
    BusinessReportBuilder(reader1, reader2, config).build(output)


def main() -> None:
    cli()


if __name__ == '__main__':
    main()
