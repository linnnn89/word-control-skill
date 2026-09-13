# Recovery and cleanup

- `open` fails: retain its `opened`, `automation_security_restored` and `errors`. An open or restoration failure may leave a document open; inspect before retrying. Do not lower macro security to bypass a failed preparation check. The command restores only the Word application's previous automation setting; it does not edit Trust Center or registry policy.

- `no running Word instance found`: open the intended document in desktop Word.
- `active document path mismatch`: bring the intended document to the front and rerun `status`.
- `selection ... changed`: rerun `selection-info`; use story type, start, end and hash from the same result.
- `target fingerprint changed`: inspect that table with `tables --table N` or rerun `equations`. An index alone is not an identity guarantee.
- `table index out of range`: refresh `tables --detail summary` before selecting an index; document edits can change indices.
- `--table` with `--max`: choose either a single table or a prefix of the collection.
- Cell pagination requires `--table` and `--detail text`; its linear indices refer to the current table, including merged cells. Follow `next_cell` only when `read_errors` is empty and the document has not changed.
- `find-text` has no matches: check `story`, `story_available`, case-sensitive literal input and `from` before concluding text is absent. Headers, footers and text boxes are outside this query's scope. A failed Find-settings restoration is an error, not a successful read.
- `set-cell` returns `verified:false`: inspect its `applied`, `readback` and `errors`. `applied:true` means the write completed but verification failed; `applied:null` means a thrown write may already have changed the document. No reusable fingerprint is returned. Inspect before retrying and do not assume rollback.
- `swap-cell-text` or `set-cell-shading` fails: retain `applied`, `readback`, `failures`, `rollback_failures` and final-inspection `errors`. `applied:true` can accompany a successful write followed by incomplete inspection; `applied:false` with `rolled_back:true` means the original text or shading values were read back after restoration; `applied:null` leaves the effect uncertain. No failed result supplies a reusable guard. Inspect before retrying, especially because repeating a successful swap reverses it.
- `create-table` fails: `applied:true` means the table was created; use `fill_complete` and `errors` to distinguish an interrupted fill from failed readback or final inspection. `applied:null` means creation itself threw with unknown effects. Inspect for an existing or partial table before retrying; the command does not delete it automatically.
- Row/column insertion or deletion fails: inspect `applied`, actual `rows`/`cols`, expected dimensions in `readback`, and all `errors`. A returned Word method can be followed by failed readback; a thrown method leaves its effects uncertain. Keep the changed document open for inspection. Do not replay deletion or automatically issue an inverse insertion/deletion; that cannot recover lost user content and may target a different row or column after indices shift.
- Whole-table or equation deletion is unverified: inspect the [deletion result](usage.md#deletion-results), remaining objects and revision state. `applied:true` with an unchanged count can represent a pending tracked deletion. Equation `prepared_inline:true` or null also requires checking its current format, even if deletion did not complete. Keep the document open for review; do not repeat deletion, switch tracking off or accept revisions to force the count to change. A successful deletion also shifts later object indices, so obtain a fresh inspection before editing another target.
- A border command fails: preserve both `failures`/`rollback_failures` and final-inspection `errors`. A readable post-check does not turn failed writes into success; failed results have no reusable fingerprint. `applied:null` or `rolled_back:true` does not authorize automatic replay. Inspect the current borders first.
- Summary/text mode has no fingerprint: request full detail for the specific table. This is an intentional omission, not a failed COM read.
- `--compare-xml` fails: inspect `xml_comparison.errors` and `read_errors`; the diagnostic provides no reusable fingerprint on failure. For ordinary editing, omit the flag and obtain a fresh complete full inspection. Never use an `xml-v1:` candidate as a mutation guard; it remains diagnostic even when two reads are stable.
- Paragraph/equation has `text:null` and `read_errors`: the read failed; keep that content unknown and re-read the affected scope after resolving the error. An `ok:true` envelope can contain partial inspection results.
- Full inspection has `fingerprint:null`, `inspection_complete:false` or `read_errors`: do not mutate from that result. Resolve the read failure and inspect again.
- Merged table has unknown row/column counts: inspect full `linear_cells` and use `set-cell --cell`. Do not force structural row/column operations on an irregular table.
- `selection is collapsed`: select the intended text; use `--allow-insert` only for intended cursor insertion.
- Backup fails or cannot include unsaved changes: stop dependent edits, preserve the open document, and follow the [backup policy](workflow.md#backup-before-editing) for an already authorized save or an explicit waiver. Do not silently save or use an older disk copy.
- Save is cancelled, pending or cannot be confirmed: leave the document open, inspect Word's save state and resolve the cause before retrying. Never follow a failed save with `close-active --discard` to finish a save request.
- Backup/PDF/smoke publication or restoration fails: inspect the final path and any reported `previous.<extension>` recovery file. Preserve the old content and resolve the file lock or access error before restoring; never overwrite a concurrent output blindly. After interruption, check the destination's `.word-control-*.tmp` directories before retrying.
- Output reports `cleanup_warning`: publication succeeded but temporary/recovery files remain at the reported path. Verify the final output before removing those files; the warning is not a request to regenerate the output.
- Word dialog blocks automation: have the visible dialog resolved; do not dismiss it or close the user's session automatically.
- Non-ASCII console text is corrupted: use UTF-8 input and `--output` files.
- Missing response or nonzero exit: inspect the current target before retrying. A failure after writing may leave changed content; report any `failures` or rollback failures. This bridge has no durable operation receipts.

## Cleanup

Delete only task-generated status, selection, table, equation, conversion and input scratch files after verification. Keep original DOCX files, explicit backups, requested PDF/PNG previews and anything needed for rollback or audit. Never quit the user's Word application.
