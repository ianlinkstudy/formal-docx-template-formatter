# Validation record

Validation date: 2026-10-27

## Functional tests

`python3 -m unittest discover -s tests -v` passes 21 tests:

- first-use template list contains only the built-in profile and requires selection;
- unconfirmed candidates cannot be saved or used;
- amendments invalidate the previous hash-bound confirmation;
- the three-template limit, archive, restore, rename, and duplicate-name rules;
- `.dotx` handling without modifying the source;
- style inheritance extraction;
- style-template formatting with formula and OLE preservation;
- shell-template formatting with fixed cover, header, footer, and anchor mapping;
- explicit body-range materialization;
- missing, duplicate, unknown, and header anchors;
- nested and merged table integrity;
- oversized inline-image scaling;
- tampered formula detection;
- CLI formatting, verification, and JSON lifecycle output;
- Skill identity, release version, runtime assets, and brand metadata.

## End-to-end examples

All examples are fictional and anonymized.

- `examples/before/01-project-report.docx` was formatted with the built-in GB reference profile;
- the same source was formatted with a synthetic company shell template;
- both outputs passed the format-time semantic check and the independent verify command;
- both reports mark content, page, paragraph typography, spacing, indent, and table checks as passed;
- template inspection and confirmation used a temporary library, not the user's default template store.

## Visual validation

- the source, GB result, and company-shell result were rendered through LibreOffice and Poppler;
- Chinese glyphs, table borders, merged cells, signatures, dates, headers, and footers were visually inspected;
- the public four-page usage guide was rendered to PNG and inspected page by page;
- Source Han Sans CN is embedded in the PDF under the SIL Open Font License;
- no font files are stored in the repository or release archive.

## Remaining limits

- Template role inference remains heuristic; confirmation is required to resolve ambiguity.
- Shell mode supports a single-section source document in the first release.
- Floating objects, nested tables, content controls, tracked changes, text boxes, comments, and footnotes require warning-driven review.
- Exact pagination must be checked on the target machine with the organization's installed fonts.
- The built-in GB profile is a reference style, not a legal-compliance determination.
