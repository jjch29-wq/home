# 비파괴검사보고서

This folder contains the NDT report application.

## Layout

- `src/비파괴검사보고서.py`: application implementation.
- `src/ndt_history_import.py`: read-only site history selector for NDT imports.
- `tests/`: focused regression tests for site history imports and report links.

The legacy entry file remains at `home/src/비파괴검사보고서.py` as a launcher so
existing shortcuts and IDE run targets continue to work.

## Report issue history

The top toolbar's `보고서 발행 이력` button (also beside RT report generation)
searches generated report snapshots by report
number, inspection date, line/drawing and Joint, and opens the generated workbook.
RT/KOGAS, PT, MT, PMI and PAUT save a history entry after successful workbook
generation. The number is read from the cover; blank numbers display as
`(번호 미입력)`. Each generation keeps its own item snapshot, including reissues.
History is stored in `data/report_issue_history.json` beside this README.
Previously generated reports are not backfilled. This records file generation,
not delivery or customer approval. Source work logs are never modified.

Double-click an item or choose `선택 항목 편집` to edit the issue's report number
and output path plus the selected item's date, drawing, Joint, section, size and
result. Report number/path changes apply to every item in that issue. Item and
whole-issue deletion are available separately. Edits retain previous values in
`revisions`; whole-issue deletion hides the entry using `deleted_at`. Existing
Excel outputs and source logs are never edited or deleted. Concurrent stale
edits are rejected so the user can refresh before saving again.
