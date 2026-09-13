# Word Control Command Guide

Set the entry point once per PowerShell session:

```powershell
$wc = "$env:USERPROFILE\.agents\skills\word-control\scripts\word_control.js"
```

## Inspect First

```powershell
cscript //nologo $wc status --output status.json
cscript //nologo $wc selection-info --output selection-info.json
cscript //nologo $wc tables --detail summary --output tables-summary.json
cscript //nologo $wc tables --detail text --output tables-text.json
cscript //nologo $wc tables --table 2 --output table-2.json
cscript //nologo $wc paragraphs --from 81 --max 40 --output paragraphs-81.json
cscript //nologo $wc find-text --input query.txt --max 5 --context 80 --output matches.json
cscript //nologo $wc tables --table 2 --detail text --cell-from 5 --cell-max 10 --output cells.json
cscript //nologo $wc equations --index 1 --output equation-1.json
```

Read-only commands are `help`, `status`, `selection`, `selection-info`, `document-text`, `find-text`, `paragraphs`, `tables`, `equations`, and `convert-equation`.

Use the active document path returned by `status` as `<doc>`. For an unsaved document, use its exact name with `--expect-name` instead.

To open a document, use `open --path "<doc>"`. The command temporarily sets Word's [AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity) to force-disable document macros, verifies that setting before opening, and then restores the original value. Failure to establish the guard prevents the open call. This controls programmatic document macros, not every Office parser, add-in or external-content behavior.

The open result retains `action` and `path`, and adds `opened`, `automation_security_restored` and `errors`. `opened:true` means Word returned an opened document; false means opening was not attempted, and null means the call threw with uncertain effects. Any preparation, open or restoration error returns a nonzero exit. Inspect the current documents and security state before retrying; a restoration failure can leave the requested document open. Failed opens only clean up an empty Word instance created by that command, never an existing user instance or an instance with documents.

### Scoped queries

- `paragraphs --from N --max count` reads only the requested entries of Word's document paragraph collection. Indices are 1-based and remain absolute in the result. `--max` defaults to 80. With `--from`, the result adds `from` and `next_from`; `next_from:null` means the end. Starting beyond the end returns an empty page. These are live indices, not stable cursors across edits; refresh after document changes.
- `tables --table N` reads only that table, in full detail by default. `table_count` remains the whole collection count and `returned` is 1. An out-of-range index fails. Do not combine `--table` and `--max`; existing `tables --max N` still reads the first N tables.
- `tables --detail summary` reads table dimensions and cell counts without retrieving cell text, formatting or fingerprints. It works with `--table` or `--max`. Every summary uses `fingerprint:null`, `inspection_complete:false`, `layout:"unverified"`, and empty cell arrays. Unknown dimensions/counts remain null with warnings/read errors. Counts alone do not prove that a table is rectangular.
- `tables --detail text` reads cell contents, including `linear_cells` for irregular/merged tables, without inspecting formatting. It works with `--table` or `--max`. It intentionally returns `fingerprint:null` and `inspection_complete:false`; check `read_errors` for missing text. Use it when reading table values is the task.
- Use `tables --table N` or explicit `--detail full` before editing. Require a complete full inspection and its fingerprint; summary/text modes cannot authorize writes.
- `tables --table N --detail text --cell-from N --cell-max count` reads only a page of Word's linear cell collection, including merged cells. Cell indices are 1-based; defaults are 1 and 40, with at most 200 cells per page. `cell_count` is the whole table count; `returned_cells` and `next_cell` describe the page. Past the end returns an empty page. Results include document identity, use `linear_cells`, and deliberately leave layout unverified and fingerprint null. Read errors invalidate pagination evidence; resolve them before continuing. New query indices are live and can shift after edits.
- `equations --index N` reads only that equation, retaining its absolute index and the whole `equation_count`. An out-of-range index fails. Omitting `--index` retains the existing all-equation output.

### Literal text location

`find-text --input query.txt` finds case-sensitive, non-overlapping literal text in the active document without selecting it. Full-width and half-width characters are distinct. The UTF-8 input must contain 1-200 characters with no control characters or final newline; `^p` is literal text, not a Word search code. This command preserves the Find settings it changes and the user's selection.

Choose `--story main` (default), `footnotes`, or `endnotes`; only that story is searched. Headers, footers, text boxes and other stories are not covered. An absent note story returns `story_available:false` with no matches. Output includes document identity, `story_type`, match `start`/`end`, bounded context, `has_more`, and `next_from`. Positions are Word character positions within that story, not indices into normalized JSON text. `--from` defaults to 0; `--max` defaults to 20 and is capped at 100; `--context` defaults to 80 characters on each side and is capped at 500. A full page checks one extra match to determine `has_more`; follow `next_from` only while the document is unchanged. `inspection_complete:true` means the requested query completed, not that every story or every page of matches was returned.

Search/read/Find-restoration failures return nonzero exit with `ok:false`, partial `matches`, `read_errors`, and unknown `has_more`. Search results are discovery evidence, not mutation guards. `find-text` and paged table queries optionally accept `--expect-path` or `--expect-name` to reject the wrong active document.

Without the new options, successful paragraph/table limits, full-detail fields and output shapes are unchanged. Failed paragraph reads return `text:null`, with top-level `inspection_complete:false` and indexed `read_errors`; they are not empty paragraphs. Failed equation reads return `text:null`, `fingerprint:null`, `inspection_complete:false` and `read_errors` on that equation. Successful equation text and its fingerprint come from one read. Do not treat a partial response as complete document evidence; resolve the error and re-read the affected scope.

Paragraph results do not provide selection mutation guards. For writing, read the [editing workflow](workflow.md) and obtain the appropriate fresh selection or object inspection.

### Experimental XML comparison

For maintenance experiments, `tables --table N --compare-xml --expect-path "<doc>"` adds `xml_comparison` to a single full table inspection. Ordinary queries do not run this comparison. The optional result contains a versioned `candidate_hash`, `stable_repeat`, character counts, elapsed milliseconds and errors; it never emits the XML itself. The ordinary `fingerprint` still comes from the existing full COM inspection.

The diagnostic reads `Range.WordOpenXML` twice, removes only editing-session `rsid` attributes and the settings `rsids` list, and compares the normalized strings. Content, styles, themes and actual tracked revisions remain. `stable_repeat:true` only establishes that these two normalized reads matched; it does not prove equivalence to effective COM formatting or stability across every edit/save cycle. `xml-v1:` hashes cannot authorize any mutation, including with an unsafe flag.

This option requires full detail and one `--table`; it cannot be combined with cell pagination or `--max`. Processing is limited to 8 Mi characters per snapshot and requires Windows MSXML 6.0, with DTDs and external resolution disabled. A parse error, oversized snapshot, unsupported snapshot shape, unstable repeat or incomplete full inspection returns nonzero exit, `ok:false`, `fingerprint:null`, `inspection_complete:false` and explicit errors. For normal editing, omit the diagnostic flag and obtain a fresh complete full inspection. See [maintenance](maintenance.md) before using this comparison to investigate performance.


## Back Up

Follow the workflow's [backup policy](workflow.md#backup-before-editing) before editing. Choose a new path with the source document's extension; `<backup.docx>` below assumes a DOCX source.

```powershell
cscript //nologo $wc save-copy --path "<backup.docx>" --expect-path "<doc>" --yes
```

If Word cannot create a copy while the source has unsaved changes, the command fails rather than copying an older on-disk version. Apply the workflow's verification and failure rules before continuing.

## Selection Changes

Use `story_type`, `start`, `end`, and `text_hash` from the same fresh `selection-info` result:

```powershell
cscript //nologo $wc replace-selection --input revised.txt --track --expect-path "<doc>" --expect-story-type <story_type> --expect-start <start> --expect-end <end> --expect-selection-hash <hash> --yes

cscript //nologo $wc insert-comment --input comment.txt --expect-path "<doc>" --expect-story-type <story_type> --expect-start <start> --expect-end <end> --expect-selection-hash <hash> --yes
```

Snapshots now require `story_type` as well as the three older selection values; refresh old snapshots and command templates. A successful `--track` edit keeps Track Changes on; a failed edit attempts to restore its previous state and reports restoration failures.

Selection changes invalidate the snapshot. Inspect again before a second change. A collapsed selection is rejected unless insertion is deliberate and `--allow-insert` is also supplied.

Comments and tracked changes can be created or toggled, but a full review-history report is not supported.

## Tables

Use the target fingerprint from a fresh full `tables --table N` result:

```powershell
cscript //nologo $wc create-table --rows 3 --cols 3 --input table.tsv --at end --expect-path "<doc>" --yes

cscript //nologo $wc set-cell --table 1 --expect-table-fingerprint <hash> --row 2 --col 2 --input cell.txt --expect-path "<doc>" --yes

cscript //nologo $wc set-cell --table 1 --expect-table-fingerprint <hash> --cell 4 --input cell.txt --expect-path "<doc>" --yes

cscript //nologo $wc swap-cell-text --table 1 --expect-table-fingerprint <hash> --from-row 1 --from-col 1 --to-row 1 --to-col 3 --expect-path "<doc>" --yes

cscript //nologo $wc insert-row --table 1 --expect-table-fingerprint <hash> --before 2 --expect-path "<doc>" --yes
cscript //nologo $wc delete-row --table 1 --expect-table-fingerprint <hash> --row 2 --expect-path "<doc>" --yes
cscript //nologo $wc insert-column --table 1 --expect-table-fingerprint <hash> --at-end --expect-path "<doc>" --yes
cscript //nologo $wc delete-column --table 1 --expect-table-fingerprint <hash> --col 4 --expect-path "<doc>" --yes

cscript //nologo $wc set-cell-shading --table 1 --expect-table-fingerprint <hash> --row 2 --col 2 --color EAF2F8 --expect-path "<doc>" --yes

cscript //nologo $wc set-table-borders --table 1 --expect-table-fingerprint <hash> --edges outer --style single --color 4472C4 --width 1 --expect-path "<doc>" --yes
cscript //nologo $wc set-table-borders --table 1 --expect-table-fingerprint <hash> --edges inside-h,inside-v --style dotted --color A6A6A6 --width 0.5 --expect-path "<doc>" --yes
cscript //nologo $wc set-cell-borders --table 1 --expect-table-fingerprint <hash> --row 2 --col 2 --edges top,bottom --style double --color 1F4E78 --width 0.75 --expect-path "<doc>" --yes

cscript //nologo $wc normalize-table-borders --table 1 --expect-table-fingerprint <hash> --color D9DEE8 --line-width 4 --expect-path "<doc>" --yes

cscript //nologo $wc delete-table --table 1 --expect-table-fingerprint <hash> --expect-path "<doc>" --yes
```

For `create-table --at selection`, also provide the fresh selection guards. TSV input larger than the requested dimensions is rejected unless `--allow-truncate` is explicit. For merged or irregular tables, use `linear_cells[].index` with `set-cell --cell`; row and column insertion/deletion require a regular table. Structural operations require Track Changes off unless `--allow-track-changes` is explicitly approved.

`swap-cell-text` moves plain text while preserving paragraph marks, manual breaks, tabs and trailing empty paragraphs. It removes only Word's terminal cell marker and compares exact text after writing. Both cell bodies are inspected before either write; a nested table's cell markers cause refusal. This command does not transfer rich formatting or promise to preserve embedded objects, fields or bookmarks within replaced text; use a narrower selection when those matter. Query text remains cleaned for display and must not be copied back as a lossless cell snapshot.

Table fingerprints include raw Word text, including paragraph and cell-boundary characters, plus the existing guarded formatting. Refresh full table inspections after updating this version; fingerprints from the previous text normalization are stale. Display output and selection/equation hashes retain their existing normalization.

`set-cell` additionally returns `document`, `applied`, `verified`, `inspection_complete`, `readback`, and `errors`. A verified success compares the actual cell content with the requested text (normalizing CRLF/CR/LF and excluding only Word's terminal cell marker), then returns the existing full post-write fingerprint. `readback` includes `matches_requested`, actual character count and text hash without echoing the whole cell. For a subsequent operation on this verified target, the returned fingerprint can replace a redundant full query: the next mutation still recomputes and compares the live full fingerprint. Requery when a different scope or visual inspection is needed.

If a write completes but readback or formatting inspection fails, `set-cell` returns `applied:true`, `verified:false`, `fingerprint:null` and a nonzero exit. If the write itself throws, `applied:null` means its effect is unknown. Inspect the current document before retrying; neither case promises rollback or permits automatic replay.

`swap-cell-text`, `set-cell-shading`, `insert-row`, `delete-row`, `insert-column`, `delete-column`, `create-table`, `set-table-borders`, `set-cell-borders` and `normalize-table-borders` also return `document`, `applied`, `verified`, `inspection_complete` and `errors`, preserving their existing success fields. Any write, readback or final-inspection error returns a nonzero exit, `verified:false`, `inspection_complete:false` and `fingerprint:null`. A later inspection failure does not replace earlier write or rollback errors. Whole-table and equation deletion use the [deletion result contract](#deletion-results).

`swap-cell-text` and `set-cell-shading` return `readback.matches_requested`, `failure_count`, `failures`, `rolled_back` and `rollback_failures`; shading readback also contains the actual `color_value` and `texture`. If only the final inspection fails, `applied:true` records the completed, read-back write without rolling it back. Write/readback failures attempt restoration: `applied:false` and `rolled_back:true` require successful restoration calls and matching original text or shading values. A restoration error, mismatch or unreadable value yields `applied:null`. This verifies the restored cell text or background/texture values, not full rich formatting, document save state or embedded objects. Inspect before retrying; replaying a completed swap would reverse it.

For row/column insertion and deletion, `applied:true` means the Word method returned; `applied:null` means it threw and may have partially changed the table. `rows` and `cols` report independently read dimensions, using null for unreadable/invalid counts. `readback` contains `expected_rows`, `expected_cols` and `matches_requested` (true/false, or null when a dimension is unknown). Dimension agreement alone does not verify retained content or layout; the full post-write fingerprint and task-specific affected-scope checks remain necessary. Write, dimension and final-inspection errors are retained together, and dimension mismatches retain the existing `error` field. No automatic structural rollback or replay occurs. Refresh or recalculate affected indices after any structural change; a new table fingerprint does not make old cell/row/column indices current.

For `create-table`, `applied:true` means Word returned the created table, even if filling it later failed; `applied:null` means creation threw and its effect is unknown. `fill_complete` is true only when every requested in-bounds TSV write returned, false if filling did not finish, and null without TSV input. Verified success also checks actual dimensions, supplied cell text and the full fingerprint. A failed table-count read returns `table_count:null`. Partial tables are left for inspection; creation is not automatically replayed or rolled back.

For border commands, `applied:true` means all requested writes returned and their immediate readbacks passed, even if the final fingerprint failed. `applied:null` indicates uncertain or partial effects; `set-*-borders` can return `applied:false` when snapshot capture failed before any write. Check `failures` and `rollback_failures` as well as final-inspection `errors`. `rolled_back:true` records successful rollback calls (or no writes), not a transaction guarantee or permission to retry.

After a verified cell text, shading, row/column or border command, reuse the returned fingerprint for the next operation on the same verified table under the [workflow checks](workflow.md#core-workflow). This avoids an extra full query; the next mutation still computes and compares a fresh live fingerprint. A different target, uncertain structural index, incomplete result or needed visual review requires inspection. After `create-table`, identify its index separately; `table_count` is not the created table's index.

Border edges are `top`, `left`, `bottom`, `right`, `inside-h`, and `inside-v`; `outer` expands to the four outside edges and `all` selects every valid edge for the scope. Supported point widths are `0.25`, `0.5`, `0.75`, `1`, `1.5`, `2.25`, `3`, `4.5`, and `6`. Use `--style none` to remove selected edges. Use the verified result or a fresh full inspection for the next fingerprint. `layout_warnings` describe non-rectangular fallback. Fingerprint read failures produce `fingerprint:null` and `inspection_complete:false`; do not mutate from an incomplete inspection. Any `read_errors`, `failure_count`, rollback failures, or a nonzero exit mean the operation is not fully verified.

### Deletion results

`delete-table` enforces the structural Track Changes guard. An explicitly approved `--allow-track-changes` permits the call, but does not establish removal: Word may retain a table or OMath object while its deletion revision is pending. `delete-equation` can still record tracked edits. Neither command changes tracking policy, accepts revisions, repeats deletion or reconstructs deleted content.

Both commands retain `remaining_tables` or `remaining_equations` and add document/target identity, `track_revisions`, `applied`, `verified`, `inspection_complete`, `readback` and `errors`. A verified success requires the delete call to complete and the relevant collection count to decrease by exactly one. `readback.expected_remaining` records that expected count; `matches_requested` is true/false, or null if the remaining count is unreadable/invalid. This count check does not establish preservation of neighboring content or document layout; perform the task's affected-scope checks separately. A deleted target has no reusable guard, so `fingerprint` is always null and remaining targets must be inspected again.

`applied:true` records a returned delete call, even when a retained revision or failed readback prevents confirmed removal. For equations, `deleted_units` is Word's reported [Range.Delete result](https://learn.microsoft.com/en-us/office/vba/api/word.range.delete), not proof that the OMath object disappeared; a zero result sets `applied:false`. A thrown call or invalid return leaves `applied:null`. Write and count-read errors are retained together. Unconfirmed removal returns a nonzero exit with `verified:false` and `inspection_complete:false`; inspect the current target and revision state before any further edit.

With tracking off, a display equation is first set to [inline form](https://learn.microsoft.com/en-us/office/vba/api/word.omath.type), so deleting its range can remove the equation instead of leaving an empty OMath at the paragraph mark. `prepared_inline` is true after that setting is read back, false when unnecessary, or null when preparation was attempted but could not be confirmed. Failed preparation prevents the delete call and returns `applied:null`. This preparation may have changed formatting even if a later deletion fails or returns zero; inspect the equation before continuing. Tracked equations retain their original form and revision history.

## Equations

Validate conversion without touching Word:

```powershell
cscript //nologo $wc convert-equation --input equation.txt --format latex --output converted.json
```

Then insert or target an inspected equation:

```powershell
cscript //nologo $wc insert-equation --input equation.txt --format latex --at end --expect-path "<doc>" --yes

cscript //nologo $wc set-equation --index 1 --expect-equation-fingerprint <hash> --input equation.txt --format latex --expect-path "<doc>" --yes

cscript //nologo $wc delete-equation --index 1 --expect-equation-fingerprint <hash> --expect-path "<doc>" --yes
```

The conservative LaTeX converter supports common fractions, square roots, integrals, sums, products, superscripts, subscripts, comparisons and Greek letters. Unknown commands, complex environments such as `align`, `matrix`, `cases`, and custom macros are unsupported. Use `--format linear` only after checking Word linear syntax.

Verify the resulting OMath object after insertion or replacement. When layout matters, export a PDF for visual review.

For `delete-equation`, follow the [deletion result contract](#deletion-results), especially when Track Changes is enabled.

## Save, Export, and Close

```powershell
cscript //nologo $wc save-active --expect-path "<doc>" --yes
cscript //nologo $wc export-pdf --path "<preview.pdf>" --expect-path "<doc>" --yes
cscript //nologo $wc close-active --save --expect-path "<doc>" --yes
```

Use exactly one of `--save` or `--discard` with `close-active`. It closes only the verified active document and never quits Word. After save confirmation, `--save` still passes Word's `wdSaveChanges` when closing so an edit introduced by a close event is not discarded; `--discard` uses `wdDoNotSaveChanges`.

`save-active` and `close-active --save` wait up to 30 seconds for Word's background-save queue to empty, then require `Document.Saved == true` and an existing nonempty file. A save error, cancellation, unreadable state or pending-save timeout returns a nonzero exit; the command does not proceed to close. The 30-second limit covers queue polling, not a blocked COM call or visible save dialog. These checks confirm Word's reported state and file presence, not storage durability or a full content readback. Reopen and compare the affected scope when the task needs that assurance.

`save-copy`, `export-pdf` and `smoke` generate into an exclusively created `.word-control-*.tmp` directory beside the destination. They require nonempty output; PDF also requires a `%PDF-` header. Only then is an approved existing output moved to `previous.<extension>` and the new file moved into place. A publication error attempts to restore the old file. If restoration is blocked, the error reports the exact retained recovery file; preserve it for inspection. A successful publication with failed cleanup reports `cleanup_warning` and the retained directory. Do not blindly retry after a missing response.

`smoke --path "<new-test.docx>" --yes` generates a disposable DOCX in its own hidden Word instance. Before publication it confirms the save queue is empty, the document is saved at the expected staging path, and the output is nonempty; it then closes that document to release the file. A generation or save failure leaves the existing destination intact. The command closes its own Word instance, and its save calls exclude staging paths from Word's recent-file list. An existing destination still requires explicit `--overwrite` approval.

This is a recoverable sequence of same-volume moves, not an atomic or crash-durable transaction: interruption between moves may leave the final path absent and the old file in the staging directory. Prefer new versioned backup names. This sequence uses existing WSH/FSO capabilities and does not add a process or runtime dependency.

PDF destinations must end in `.pdf` and differ from the active source, including case and Windows 8.3 aliases. Backup, export and smoke paths must be literal file paths without wildcards, alternate streams or trailing dots/spaces; smoke requires `.docx`. `save-copy` does not change the source format; its disk-copy fallback remains limited to a saved document with unchanged identity and no pending save. It never silently saves unsaved changes to make a backup possible.

## Troubleshooting

See [Recovery](recovery.md) for stale targets, incomplete inspections and blocked Word automation.

## Cleanup

Follow [task-artifact cleanup](recovery.md#cleanup); preserve source documents, backups and requested review outputs.

## Maintenance

See [Maintenance and validation](maintenance.md) for the canonical source, isolated tests and verified synchronization.
