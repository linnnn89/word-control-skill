# Recovery and cleanup

- `no running Word instance found`: open the intended document in desktop Word.
- `active document path mismatch`: bring the intended document to the front and rerun `status`.
- `selection ... changed`: rerun `selection-info`; use story type, start, end and hash from the same result.
- `target fingerprint changed`: inspect that table with `tables --table N` or rerun `equations`. An index alone is not an identity guarantee.
- `table index out of range`: refresh `tables --detail summary` before selecting an index; document edits can change indices.
- `--table` with `--max`: choose either a single table or a prefix of the collection.
- Cell pagination requires `--table` and `--detail text`; its linear indices refer to the current table, including merged cells. Follow `next_cell` only when `read_errors` is empty and the document has not changed.
- `find-text` has no matches: check `story`, `story_available`, case-sensitive literal input and `from` before concluding text is absent. Headers, footers and text boxes are outside this query's scope. A failed Find-settings restoration is an error, not a successful read.
- `set-cell` returns `verified:false`: inspect its `applied`, `readback` and `errors`. `applied:true` means the write completed but verification failed; `applied:null` means a thrown write may already have changed the document. No reusable fingerprint is returned. Inspect before retrying and do not assume rollback.
- `create-table` fails: `applied:true` means the table was created; use `fill_complete` and `errors` to distinguish an interrupted fill from failed readback or final inspection. `applied:null` means creation itself threw with unknown effects. Inspect for an existing or partial table before retrying; the command does not delete it automatically.
- A border command fails: preserve both `failures`/`rollback_failures` and final-inspection `errors`. A readable post-check does not turn failed writes into success; failed results have no reusable fingerprint. `applied:null` or `rolled_back:true` does not authorize automatic replay. Inspect the current borders first.
- Summary/text mode has no fingerprint: request full detail for the specific table. This is an intentional omission, not a failed COM read.
- Paragraph/equation has `text:null` and `read_errors`: the read failed; keep that content unknown and re-read the affected scope after resolving the error. An `ok:true` envelope can contain partial inspection results.
- Full inspection has `fingerprint:null`, `inspection_complete:false` or `read_errors`: do not mutate from that result. Resolve the read failure and inspect again.
- Merged table has unknown row/column counts: inspect full `linear_cells` and use `set-cell --cell`. Do not force structural row/column operations on an irregular table.
- `selection is collapsed`: select the intended text; use `--allow-insert` only for intended cursor insertion.
- Backup fails with unsaved changes: do not silently save or use an older disk copy. Obtain the user's decision unless the existing request already authorizes the necessary action.
- Word dialog blocks automation: have the visible dialog resolved; do not dismiss it or close the user's session automatically.
- Non-ASCII console text is corrupted: use UTF-8 input and `--output` files.
- Missing response or nonzero exit: inspect the current target before retrying. A failure after writing may leave changed content; report any `failures` or rollback failures. This bridge has no durable operation receipts.

## Cleanup

Delete only task-generated status, selection, table, equation, conversion and input scratch files after verification. Keep original DOCX files, explicit backups, requested PDF/PNG previews and anything needed for rollback or audit. Never quit the user's Word application.
