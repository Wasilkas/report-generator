# Changelog

## [Unreleased]

### Added
- **Class name matching** — only classes present in both models are compared. Classes absent from either model are routed to the *Удаленные классы* sheet with a reason (`Только в новой модели` / `Только в прод модели`).
- **Reason column** (`Причина`) on the excluded-classes sheet — distinguishes low-training-count exclusions from unmatched-class exclusions.
- **Comparison sheet legend** — explanation of green/red colouring (improvement vs. significant degradation) written below the data on the *Сравнение* sheet (both report types).
- **Verdict sheet legend** — explanation of 0/1/2 scores (`Хуже` / `Незначительно лучше` / `Лучше`) written below the verdict table on the business report.
- **Excluded-classes sheet title** now shows the actual threshold value (`≤ N`) for interpretability.
- **`app_config.py`** — centralised string constants (`CLASS_COL`, `TRAIN_COUNT_DISPLAY_COL`, `REASON_COL`) shared across all modules.
- `ExcelSheetWriter.write()` now returns the last written row index so callers can position content (e.g. legends) below the data.
- `write_comparison_legend()` helper in `core/writer.py`.
- `exclude_unmatched()` function in `core/calculator.py`.

### Changed
- Excluded-classes sheet columns standardised to `[Класс, Число примеров, Причина]` across both report types (ID column removed).
- Dev report: global class IDs are inserted *before* filtering so excluded classes preserve their original index gaps.
