# Changelog

## [Unreleased]

## [0.3.0] - 2026-11-21

### Changed

- The business report has one fixed-column layout, written by `generate-report business`,
  the legacy `BusinessReportBuilder` and `v2.BusinessReportBuilder`. The metric sheets show
  only `Класс`, `tp`, `fp`, `fn`, `Порог`, `precision`, `recall`, `f1_score`,
  `Перебраковка, %`, `Недобраковка, %`, the four goal columns and the example counts by
  split, in that order. A listed column that neither input supplies is omitted, and the
  unsplit `Количество примеров` column stands in when no input has a split count.
  `business.column_translations` still renames the headers that remain.
- In the business `Среднее` row, columns in `excluded_from_mean` are `NA`, example-count
  columns show the sum over the model's own eligible classes, and the other columns keep
  their mean.
- The business `Сравнение` sheet shows new − prod for every column, including `tp`, `fp`,
  `fn`, `Порог`, the goal columns and the example counts. Its `Среднее` row aggregates the
  differences over shared classes. Only the quality and percentage columns are coloured.
  The legend is four text lines in column A.
- The business population note is five rows in column A on the model, comparison and
  verdict sheets.
- `v2.BusinessReportBuilder(model1, model2, config)` no longer takes `presentation`.

### Removed

- The previous business layout: every input column, the `<metric> coverage` columns and
  the `Valid class counts` block are gone from the business report. The developer report
  is unchanged.
- The `translations` argument of `compact_report_counts`, which only business reports used.

### Fixed

- `annotate_workbook` keeps workbooks saved by Excel openable: a rewritten part keeps the
  declaration of every namespace prefix that its markup-compatibility attributes name
  (`mc:Ignorable`, `mc:ProcessContent`, `mc:MustUnderstand`, `mc:PreserveElements`,
  `mc:PreserveAttributes` and `Requires` on `mc:Choice`). Workbooks written by openpyxl are
  annotated byte for byte as before.

## [0.2.0] - 2026-11-09

### Added

- `report_generator.v2`: structured `DevReportBuilder`/`BusinessReportBuilder` with
  `PresentationOptions`, `MetricsTable` input, `write_comparison_workbook`,
  `compact_report_counts` and reversible workbook annotation (`annotate_workbook`,
  `original_workbook`, `read_metadata`, `package_fingerprint`). Legacy module interfaces
  are unchanged.
- `annotate_workbook(placement="bottom")` appends caller banners one empty row after each
  worksheet's last cell, merged range or drawing without moving original coordinates, with
  optional per-line http(s) `banner_links` and a printed `page_footer`. The default
  prepend placement is unchanged.

## [0.1.2] - 2026-10-05

### Fixed

- Retain new and deleted classes in developer and business metric sheets, preserving
  the available model's measurements and rendering unavailable model values and deltas as `NA`.
- Apply training eligibility per model and calculate model means and business verdict
  inputs before aligning display rows; disclose the model and comparison populations.
- Preserve missing-value semantics, real zeros, class identities and neutral styling
  for unavailable comparisons, including disjoint model vocabularies.

## [0.1.1] - 2026-10-05

### Added
- **Class name matching** — only classes present in both models are compared. Classes absent from either model are routed to the *Удаленные классы* sheet with a reason (`Только в новой модели` / `Только в прод модели`).
- **Reason column** (`Причина`) on the excluded-classes sheet — distinguishes low-training-count exclusions from unmatched-class exclusions.
- **Comparison sheet legend** — explanation of green/red colouring (improvement vs. significant degradation) written below the data on the *Сравнение* sheet (both report types).
- **Verdict sheet legend** — explanation of 0/1/2 scores (`Хуже` / `Без изменений или незначительно лучше` / `Лучше`) written below the verdict table on the business report.
- **Excluded-classes sheet title** describes mixed reasons and source models.
- **`app_config.py`** — centralised string constants (`CLASS_COL`, `TRAIN_COUNT_DISPLAY_COL`, `REASON_COL`) shared across all modules.
- `ExcelSheetWriter.write()` now returns the last written row index so callers can position content (e.g. legends) below the data.
- `write_comparison_legend()` helper in `core/writer.py`.
- `exclude_unmatched()` function in `core/calculator.py`.

### Changed
- Excluded-classes sheet columns standardised to `[Класс, Число примеров, Причина, Модель]` across both report types (ID column removed).
- Dev report: supplied valid IDs are preserved; IDs are generated only when absent, before filtering.

### Fixed

- Normalize class identity and reject duplicates; filter numeric-looking threshold labels.
- Validate configuration shapes, domains, sheet names, identity collisions and directions.
- Retain the minimum training count; reject malformed counts and preserve both models’ exclusion provenance.
- Compute finite arithmetic macro means with availability coverage; leave unavailable metric differences blank.
- Keep count/ID/confidence metadata out of comparisons independently of mean exclusions.
- Preserve unavailable business values, inclusive targets, strict OR gross violations, and mandatory rollout criteria.
- Handle equality/zero baselines correctly and preserve styling under translations.
- Support explicit sheet selectors and configurable verdict names; correct legends and exclusion headings.
- Reject input/output aliases and require explicit permission to replace existing reports or samples.
- Make sample generation import-safe, deterministic and destination-aware; report expected CLI errors actionably.

See [the dated remediation ledger](docs/remediation-20261004.md) for focused verification and limits.
