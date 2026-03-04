# report-generator

Builds formatted Excel comparison reports from two model metrics files.

## Features

- **Dev report** — 4-sheet Excel workbook:
  - Sheet 1 & 2: per-class metrics for each model (precision, recall, F1, AP50, AP75, AP50-95), colour-coded headers and a weighted mean row
  - Sheet 3 (*Сравнение*): side-by-side diff with green/red highlights for improvements and degradations beyond a configurable threshold
  - Sheet 4 (*Удаленные классы*): classes excluded due to insufficient training examples

- **Business report** — stakeholder-friendly view:
  - Perebrak / nedobrak columns converted to percentages
  - Binary goal flags: whether each class meets the configured perebrak/nedobrak targets
  - Verdict sheet summarising model comparison with a score

- **Configurable** — thresholds, targets, colours, and sheet names are controlled via a YAML config file. Override any subset of keys without touching the defaults.

## Installation

```bash
pip install git+https://github.com/Wasilkas/report-generator.git
```

Or, for development:

```bash
git clone https://github.com/Wasilkas/report-generator.git
cd report-generator
pip install -e .
```

## Usage

```bash
# Developer report (default config)
generate-report dev model_new.xlsx model_prod.xlsx report_dev.xlsx

# Business report
generate-report business model_new.xlsx model_prod.xlsx report_business.xlsx

# Select a specific sheet inside the Excel files (name or 0-based index)
generate-report dev model_new.xlsx model_prod.xlsx out.xlsx --sheet1 "Metrics" --sheet2 1

# Override config values at runtime
generate-report --config my_config.yaml dev model_new.xlsx model_prod.xlsx out.xlsx
```

The tool is also runnable as a module:

```bash
python -m report_generator.cli dev model_new.xlsx model_prod.xlsx out.xlsx
```

## Configuration

Edit `report_generator/configs/config.yaml` to change default behaviour, or pass `--config path/to/override.yaml` to override specific keys at runtime.

| Key | Default | Description |
|-----|---------|-------------|
| `min_train_count` | `20` | Classes with fewer training examples are excluded |
| `degradation_threshold` | `0.05` | Absolute diff threshold for red highlighting on the comparison sheet |
| `business.target_perebrak` | `0.3` | Perebrak goal threshold |
| `business.target_nedobrak` | `0.2` | Nedobrak goal threshold |
| `business.comparison_pct_threshold` | `5.0` | Pct-point threshold for red/green on the business comparison sheet |
| `business.verdict_score_threshold` | `0.05` | Relative-diff threshold for verdict scoring |

## Project structure

```
report_generator/
    cli.py          # Click CLI entry point
    config.py       # Config loader (YAML → frozen dataclasses)
    configs/        # Bundled default YAML files
    core/
        reader.py       # MetricsReader — reads Excel into DataFrames
        calculator.py   # Mean row and comparison calculators
        writer.py       # ExcelSheetWriter — writes/styles sheets
    reports/
        base.py         # BaseReportBuilder (abstract)
        dev/
            builder.py  # DevReportBuilder
            utils.py    # Sheet helpers for the dev report
        business/
            builder.py  # BusinessReportBuilder
            utils.py    # Data prep and verdict helpers
```
