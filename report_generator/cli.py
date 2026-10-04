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

from pathlib import Path
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

import click
import yaml
from openpyxl.utils.exceptions import InvalidFileException

from .config import Config
from .core.reader import MetricsReader
from .reports.business.builder import BusinessReportBuilder
from .reports.dev.builder import DevReportBuilder


def _coerce_sheet(value: str) -> str | int:
    if value.startswith('name:'):
        name = value[5:]
        if not name:
            raise ValueError('Sheet name must not be empty')
        return name
    explicit_index = value.startswith('index:')
    if explicit_index:
        value = value[6:]
    try:
        index = int(value)
        if index < 0:
            raise ValueError('Sheet index must be nonnegative')
        return index
    except (TypeError, ValueError):
        if explicit_index or value.lstrip('-').isdigit():
            raise ValueError(f'Invalid sheet index: {value!r}') from None
        return value


# ── Shared options ─────────────────────────────────────────────────────────────


def _shared_options(f):
    for decorator in reversed(
        [
            click.argument('model1', type=click.Path(exists=True, dir_okay=False, path_type=Path)),
            click.argument('model2', type=click.Path(exists=True, dir_okay=False, path_type=Path)),
            click.argument('output', type=click.Path(dir_okay=False, path_type=Path)),
            click.option(
                '--force', is_flag=True, help='Replace an existing output; never an input.'
            ),
            click.option(
                '--sheet1',
                default='0',
                show_default=True,
                help='name:NAME or index:N (bare numeric index also accepted) for MODEL1.',
            ),
            click.option(
                '--sheet2',
                default='0',
                show_default=True,
                help='name:NAME or index:N (bare numeric index also accepted) for MODEL2.',
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
    try:
        ctx.obj['config'] = Config.load(config_path)
    except (ValueError, OSError, yaml.YAMLError) as exc:
        raise click.ClickException(f'Invalid configuration: {exc}') from exc


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
    force: bool,
) -> None:
    """Build a 4-sheet developer report (metrics + comparison + excluded classes)."""
    config: Config = ctx.obj['config']
    try:
        reader1 = MetricsReader(model1, sheet_name=_coerce_sheet(sheet1))
        reader2 = MetricsReader(model2, sheet_name=_coerce_sheet(sheet2))
        DevReportBuilder(reader1, reader2, config).build(output, overwrite=force)
    except (ValueError, OSError, BadZipFile, InvalidFileException, ParseError) as exc:
        raise click.ClickException(f'Cannot generate report: {exc}') from exc


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
    force: bool,
) -> None:
    """Build a business-oriented report."""
    config: Config = ctx.obj['config']
    try:
        reader1 = MetricsReader(model1, sheet_name=_coerce_sheet(sheet1))
        reader2 = MetricsReader(model2, sheet_name=_coerce_sheet(sheet2))
        BusinessReportBuilder(reader1, reader2, config).build(output, overwrite=force)
    except (ValueError, OSError, BadZipFile, InvalidFileException, ParseError) as exc:
        raise click.ClickException(f'Cannot generate report: {exc}') from exc


def main() -> None:
    cli()


if __name__ == '__main__':
    main()
