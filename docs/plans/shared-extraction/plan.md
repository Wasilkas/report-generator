# Shared report extraction

The legacy file-reader, constructor, configuration and CLI contracts remain unchanged. The additive `report_generator.v2` API accepts validated tables with exact lexical class labels and explicit presentation options. Both entry points share calculators and table rendering; legacy output retains coverage columns while v2 can render counts directly below each table. Generic OOXML operations know metadata location/envelope and supplied text, never application identity schemas. Statistical comparison rendering consumes precomputed result columns without computing statistics.

## Architecture

[Machine-readable architecture](architecture.json) and [API contract](api.md) defines package isolation and write ownership. Package code may depend on pandas/openpyxl and the standard library, never clearml_yolo, digital_metrics, ClearML or inference backends. Builders own workbook filesystem effects; caller owns evaluation and identity schema validation.

## Ordered work

- [x] Inspect existing inputs/readers/calculators/builders and application rendering/identity adapters.
- [x] Run legacy tests and capture report workbook/config/signature/CLI characterizations before internal edits.
- [x] Validate architecture planning record.
- [x] Add structured inputs/presentation wrappers sharing unchanged computations and explicit rendering hooks.
- [x] Add standalone precomputed statistical comparison renderer and generic OOXML banner/metadata utilities.
- [x] Verify legacy characterizations, new table rendering and OOXML recovery/fingerprint; exercise prohibited-import negative case.
- [x] Update interface documentation and run Markdown/local-link checks.
- [ ] Obtain parent integrated independent review and reconcile architecture acceptance record before delivery.

Parent application integration can proceed after interface agreement while this package implementation proceeds. Integration, final verification and independent review are sequential. This bounded implementation does not delegate or commit.
