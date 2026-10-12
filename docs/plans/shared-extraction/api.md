# Structured spreadsheet interfaces

`report_generator.v2` is additive. Legacy imports, signatures, constructor defaults,
configuration fields, YAML loading and CLI entry points retain their existing behavior.
The legacy reader still trims labels and excludes numeric-looking strings; structured
input retains exact lexical labels and rejects duplicates without trimming.

`MetricsTable(frame, source_name, source_path=None)` takes a defensive copy of a pandas
DataFrame containing `Класс`. Labels must be nonempty strings distinct by exact value;
`Среднее` is reserved for aggregate rows. Optional IDs must be nonempty unique strings
or finite numbers. The stored frame remains a pandas DataFrame; callers should treat it
as a snapshot. Builders copy it before calculations and do not mutate it. A supplied
source path enforces workbook alias protection; table-only callers need only a display
source name. `read_metrics_table(file_path, sheet_name=0)` obtains text labels from
original Excel cells, retains their lexical spelling, and filters numeric footer cells.
Missing metric handling remains pandas' normal handling.

`DevReportBuilder(table1, table2, config, *, presentation=...)` and
`BusinessReportBuilder(...)` accept explicit `PresentationOptions(counts=...)`.
`counts` is `columns` or `below_table`; the latter separates computed coverage summaries
before rendering and writes valid/eligible counts directly below tables. Titles, class
sorting, filtering, missing placeholders, means, comparison eligibility, business goals,
verdicts, color and number-format calculations remain shared with legacy builders.
Counts preserve original aggregate populations, including zero and missing values.
`count_title` and `count_labels` allow explicit display wording; default AP labels are
AP50, AP75 and AP50-95. `build(path, overwrite=False)` preserves alias and replacement
protection. `compact_report_counts` remains available for consumers of old layouts.

`write_comparison_workbook(rows, excluded, methodology, path)` is independent from
those report builders. It consumes all canonical columns in
`report_generator.v2.comparison.COMPARISON_COLUMNS` plus `is_pooled`, places pooled
results last, applies the existing Russian labels/colors, preserves tiny p-values,
and writes exclusions and methodology CSVs. It calculates no evaluation statistics.
It returns the CSV path mapping and does not mutate input tables.

`MetadataFormat(path, namespace, root_name, version, relationship_id)` describes an
existing custom XML envelope. `annotate_workbook(path, *, banner_lines, metadata_texts,
format, validate=None)` adds literal text banners, adapts worksheet coordinates and
associated OOXML references, and embeds the original ZIP bytes. `metadata_texts` maps
XML local names to caller-serialized text; the package owns `original` and the
`package_sha256` attribute. `read_metadata(content, format, validate=None)` checks the
envelope, fingerprint presence, optional caller schema validation and complete package
fingerprint in that order. `original_workbook(content, format, validate=None)` recovers
the exact original bytes. `package_fingerprint(parts, format)` retains the existing
algorithm over sorted part names and SHA-256 digests, excluding only its own attribute.
These utilities do not import, define or validate application identity models.

An existing version-1 application envelope is exercised by a characterized artifact;
the shared utilities introduce no replacement storage format. Generic nonapplication
metadata is tested separately. Characterized payloads cover cached original bytes,
formulas, merged cells, tables, validation, charts, filters, frozen panes, styles,
page breaks and repeated print headings. The original package must remain authenticated
before schema recovery or reannotation can use its embedded historical metric cells.

See [the architecture record](architecture.json), [the ordered plan](plan.md), and
[dated verification evidence](evidence-20261012.md).
