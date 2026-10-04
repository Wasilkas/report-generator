# Report generator remediation — 2026-10-04

Scope: all 25 findings from the source review of revision
`2172d3945723233b087b1375c2c821d020462f00`. This document records the accepted
contracts, implementation and focused verification without depending on an unpublished
audit. Historical reproduction scripts assert old symptoms and are not regression tests.

The unfinished candidate patch was inspected before extension. Its 30 regression cases
passed. Twelve additional cases then reproduced gaps in finite means/coverage, null
configuration, generated identity conflicts, direction conflicts, business warnings,
translated class styling, zero-threshold equality and sample destination preflight.
Those cases passed after remediation. The final tests exercise public helpers, CLI
commands and reopened saved workbooks rather than claiming Excel GUI rendering.

## Findings and dispositions

1. **cli-output-overwrites-source — fixed.** Previously, choosing either source as
   output destroyed it. Both builders preflight resolved paths and existing-file identity;
   explicit replacement permission never permits an input alias. Verification:
   `test_workbook_provenance_id_and_output` covers both inputs, symlinks and hard links,
   on both builders. Limits: no concurrent filesystem replacement race test.

2. **core-comparison-unavailable — fixed.** Candidate values survived when the baseline
   was absent or malformed, producing false improvements. Comparison now builds a fresh
   metric-only table, coerces invalid/nonfinite operands to unavailable, and preserves
   unavailable differences. Verification: `test_macro_and_numeric_differences` and
   `test_saved_workbook_unavailable_and_legends` reopen both report modes and confirm
   blank values without positive fill. Limits: synthetic workbooks.

3. **core-exclusion-provenance — fixed.** Production-side training reasons/counts were
   lost or mislabelled as unmatched. Preparation now records excluded rows separately
   for each original model, all applicable reasons and original counts; a peer's training
   exclusion is distinguished from true absence. Verification:
   `test_workbook_provenance_id_and_output` checks production count/reason/model in saved
   workbooks. Limits: synthetic asymmetric fixture.

4. **dev-existing-id — fixed.** Unconditional ID insertion crashed on valid supplied
   identifiers. Readers reject empty, invalid or duplicate supplied IDs; developer reports
   preserve valid IDs and generate them only when the column is absent, before filtering.
   Verification: `test_id_preserved`, `test_invalid_ids`, and workbook provenance tests.
   Limits: no production identifier schema is assumed beyond strings/finite numbers.

5. **core-unweighted-means — resolved and fixed.** Documentation promised weighted means
   while implementation computed arithmetic means. The accepted contract is a finite
   arithmetic macro mean, explicitly documented in Russian README. Each averaged metric
   has a separate `<metric> coverage` column containing `valid/total` in the mean row.
   Verification: `test_macro_and_numeric_differences`,
   `test_finite_macro_coverage_and_comparison_identity` and saved AP coverage tests.
   Limits: coverage describes availability, not confidence or weighting.

6. **input-numeric-sheet-name-shadowed — fixed.** Numeric names were consumed as indices.
   Selectors now distinguish `name:NAME` from `index:N`; bare numeric values remain
   nonnegative indices. Verification: `test_cli_selectors_errors_and_force` and
   `test_numeric_sheet_name_and_custom_verdict` run both modes against a named `1` sheet.
   Limits: API already accepts strings as names and integers as indices.

7. **config-validation — fixed.** Unknown keys and invalid types/domains passed through
   or crashed internally. Loading now validates mapping shapes, explicit nulls, string
   keys, numeric finiteness/domains, lists, colors, Excel sheet names and uniqueness,
   translations, generated identities and opposing metric directions. Runtime input
   collisions are also rejected. Verification: `test_config_rejects_invalid`,
   `test_config_null_identity_conflicts`, translation tests. Limits: frozen dataclass
   replacement is an internal Python escape hatch, not a YAML validation interface.

8. **business-invalid-goal-values — fixed.** Invalid ratios fabricated zero flags.
   Required business ratios are coerced only when finite, nonboolean and in `[0,1]`;
   missing/invalid values remain unavailable with warnings and empty flags. Intentional
   AP NaNs are preserved. Verification: `test_business_boundaries_invalid_and_gross`,
   `test_business_unavailable_warns_and_equality_zero` and saved workbook tests.
   Limits: warnings use Python's standard warning mechanism.

9. **business-comparison-legend-threshold — fixed.** The legend reported a ratio threshold
   for percentage-point columns. Business legend explicitly lists ratio and percentage-point
   thresholds. Verification: saved workbook legend asserts `5.0 п.п.`. Limits: no Excel GUI.

10. **business-excluded-sheet-labels — fixed.** Empty class header and training-only title
    misrepresented mixed exclusions. Both report modes now show a class heading and a title
    describing reasons and source models; custom business heading translation is retained.
    Verification: saved provenance workbook tests. Limits: no font rendering comparison.

11. **business-verdict-equality-label — fixed.** Equality was described as improvement and
    zero thresholds scored equality as 2. Equality now always scores 1, with a legend that
    includes unchanged values. Verification: `test_verdict_zero_baselines` and
    `test_business_unavailable_warns_and_equality_zero`. Limits: none for tested boundaries.

12. **cli-raw-traceback — fixed.** Expected invalid-workbook, configuration, sheet and
    output failures escaped as internal diagnostics. CLI boundaries convert known input
    and filesystem exceptions into actionable Click errors. Unexpected runtime defects
    still propagate. Verification: CLI malformed-workbook/configuration/output tests and
    `test_cli_unexpected_internal_error_is_not_masked`. Limits: not every third-party
    parser exception has a synthetic fixture.

13. **business-verdict-skips-unscorable — resolved and fixed.** Partial criteria could
    recommend rollout. All three criteria are required, and every included class must
    supply valid required values. Missing coverage yields `Недостаточно данных` even
    when the macro mean is finite. A zero baseline has unavailable relative percentage;
    equality scores 1 and zero-to-positive scores by direction. Verification:
    `test_rollout_partial_class_and_zero_gross`, `test_verdict_zero_baselines`, saved tests.
    Limits: synthetic mandatory-criteria scenarios.

14. **business-gross-aggregation-ambiguous — resolved and fixed.** Gross failures counted
    only simultaneous violations. Accepted policy counts either error metric strictly
    above twice its target. Verification: `test_business_boundaries_invalid_and_gross`
    and `test_rollout_partial_class_and_zero_gross`. Limits: none for tested OR boundary.

15. **core-duplicate-classes — resolved and fixed.** Duplicate labels broadcast or crashed
    depending on model order. Reader now rejects duplicate normalized identities before
    alignment. Verification: `test_reader_normalization`. Limits: case remains significant.

16. **core-training-schema — resolved and fixed.** Missing, ambiguous and malformed counts
    had implicit behavior. Missing counts warn and disable filtering; multiple recognized
    count columns and malformed, negative, nonfinite, fractional or boolean counts fail.
    Verification: `test_invalid_training_count` and `test_training_schema_cutoff`.
    Limits: recognition intentionally requires `train` and `пример` in the column name.

17. **input-class-whitespace-not-normalized — resolved and fixed.** Outer spaces generated
    false unmatched classes. Reader trims before matching and duplicate detection.
    Verification: `test_reader_normalization`. Limits: internal whitespace is preserved.

18. **core-training-cutoff — resolved and fixed.** Exactly the configured minimum was
    excluded. Accepted contract retains equality and excludes strictly fewer examples;
    the reason label uses `<`. Verification: `test_training_schema_cutoff` and saved
    workbook tests with count 20. Limits: none for tested boundary.

19. **core-comparison-counts — resolved and fixed.** Metadata/counts appeared as candidate
    passthrough values and comparison policy depended on mean exclusions. Comparisons
    now omit IDs, confidence, configured integer columns and count identities, including
    unfamiliar `count`, `пример`, and separate `n` names. Metric comparison is independent
    of `excluded_from_mean`. Verification: `test_macro_and_numeric_differences` and
    `test_finite_macro_coverage_and_comparison_identity`. Limits: arbitrary unlabelled
    numeric columns cannot be inferred as counts without a naming/configuration contract.

20. **business-goal-boundaries-ambiguous — resolved and fixed.** Equality at the target
    and doubled target lacked a contract. Targets are inclusive (`≤`); gross violations
    are strict (`> 2 × target`). Verification: business boundary tests. Limits: none for
    tested finite ratio values.

21. **business-custom-translation-styling — resolved and fixed.** Accepted translations
    disconnected number formats, comparison colors and class/mean identity styling.
    Formatting/color maps follow translated source identities and writers retain the
    translated class identity. Header text and mean rows stay styled. Collisions fail.
    Verification: `test_custom_translations` and
    `test_translated_class_mean_saved_styles` inspect reopened workbook cells. Limits:
    no visual screenshot test; source semantic names determine behavior.

22. **business-verdict-sheet-name — resolved and fixed.** A hardcoded verdict name ignored
    sheet configuration. `sheet_names.verdict` now controls workbook topology and shares
    Excel-name validation. Verification: `test_numeric_sheet_name_and_custom_verdict`.
    Limits: none for tested custom name.

23. **input-numeric-text-class-retained — resolved and fixed.** Numeric Excel cells were
    removed but numeric-looking text remained. Both forms are now threshold rows;
    numeric-only text class names are explicitly unsupported. Verification:
    `test_reader_normalization` covers decimal/scientific text. Limits: Python numeric
    parsing defines numeric-looking text; localized comma decimals are ordinary labels.

24. **cli-existing-output-overwritten — resolved and fixed.** Existing output was replaced
    silently. Both CLI and builder APIs require explicit overwrite permission for existing
    outputs while preserving mandatory alias rejection. Verification:
    `test_cli_selectors_errors_and_force`, both-builder output tests. Limits: preflight
    does not promise transactional protection against concurrent filesystem mutation.

25. **sample-fixed-output-overwrite — resolved and fixed.** Import seeded global random
    state and wrote fixed files; execution overwrote destinations. Sample generation now
    has a guarded entrypoint, local deterministic random generator, configurable distinct
    `.xlsx` destinations, explicit overwrite and all-destination preflight before writes.
    Verification: `test_sample_import_and_generation` and
    `test_sample_preflight_prevents_partial_write`. Limits: unexpected storage failure
    during a write is not a transactional two-file operation.

## Verification and documentation

Commands run in this repository:

```bash
uv run --with pytest pytest -q
uv run --with ruff ruff check report_generator create_test_data.py tests
uv run --with ruff ruff format --check report_generator create_test_data.py tests
git diff --check
```

Final focused suite after review corrections: **83 passed**, including wheel/source-distribution construction,
installed wheel resources and the real generated console wrapper in
`test_wheel_resources_installed_entrypoint`. Tests use temporary directories and clean up
samples, installed wheel target and generated reports. Two warnings in the boundary helper
fixture explicitly demonstrate unavailable business data. No services were started.

README was rewritten in Russian to describe the chosen identity, count, arithmetic,
availability, boundaries, overwrite, selector and configuration contracts. This ledger
provides all findings' observation, disposition, changes, checks and limitations directly.
Historical bulky audit artifacts remain local and unchanged. No dependency reference,
project version, upstream `digital-metrics`, commit or push was changed by this work.

Scope limits: one current Python/dependency stack, synthetic files, no production workbook,
minimum-version/platform matrix, ClearML execution or Excel GUI rendering. Parent-owned
integration verification and fresh independent review are separate acceptance stages.

## Independent review corrections — 2026-10-04

A fresh parent-owned review returned fix-first with three additional boundary failures.
The new focused regressions reproduced 21 failing cases before correction; a legitimate
metric-alias control passed. All 22 new cases and the prior 55 cases now pass.

- **Recursive mixed configuration keys (finding 7):** nested unknown integer/string keys
  reached sorting and raised `TypeError`. Every recursive merge now validates string keys
  before computing or sorting unknown keys. `test_recursive_config_keys_have_actionable_cli_error`
  checks both API rejection and actionable CLI configuration diagnostics.
- **Business sources aliasing metadata (finding 7/21):** `f1_col: Класс` erased class labels,
  and `perebrak_col: ID` scaled IDs as percentages. Configuration now rejects source aliases
  to class identity, IDs, confidence, count metadata and configured integer columns.
  `test_business_source_alias_cannot_mutate_metadata` covers all three business source roles
  with six metadata identities; `test_business_legitimate_metric_source_alias` confirms that
  valid existing/custom metric names remain configurable.
- **Nested sample destinations (finding 25):** choosing one workbook as the other workbook's
  parent passed preflight and caused a partial write. Resolved ancestor/descendant destinations
  now fail before directory creation or file writes in either order.
  `test_sample_rejects_nested_destinations_before_write` verifies both orders leave the first
  destination absent.

Full suite, Ruff lint/format, Markdown local-link/fence checks and `git diff --check`
passed after these corrections. No commits or pushes were made. Parent-owned fresh
independent re-review remains the next acceptance step.

## Semantic alias direction correction — 2026-10-04

The next fresh review identified a conflicting business error alias: using `precision` as
`perebrak_col` interpreted a +40 percentage-point deterioration through the configured
higher-is-better color direction. Configuration now rejects either business error source
in `better_higher_cols` and the F1 source in `better_lower_cols`. Distinct legitimate
custom metric aliases remain supported, and the configured F1 source receives the
higher-is-better business color semantics even when its name is custom.

Five configuration regressions failed before this correction (one already rejected by
source-identity collision but lacked the direction diagnostic). After correction,
`test_business_source_direction_conflicts_rejected` checks both default and custom conflicts,
and `test_custom_business_error_alias_saved_red` reopens a saved workbook to verify a
custom error's +40 percentage-point difference is red and a custom quality improvement
is green. The full suite now has 83 passing cases; Ruff lint/format, Markdown local-link
and fence checks, and `git diff --check` pass. Fresh parent-owned review remains pending.

## Parent acceptance

The parent independently reran all 83 tests, Ruff and whitespace checks after review
corrections. Fresh independent review returned `ship` on 2026-10-04. The proposed
report package also generated developer and business workbooks in real CPU and
single-GPU YOLO comparison runs; those published workbooks were downloaded. Normal
integration used fresh destinations and did not opt into overwrite.
