# Changelog

## [Unreleased]

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
