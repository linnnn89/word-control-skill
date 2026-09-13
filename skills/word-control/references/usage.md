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


## Back Up

Choose a new path; do not overwrite an existing backup by default.

```powershell
cscript //nologo $wc save-copy --path "<backup.docx>" --expect-path "<doc>" --yes
```

If Word cannot create a copy while the source has unsaved changes, the command fails rather than copying an older on-disk version. Do not silently save the source; obtain approval first.

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

For `create-table --at selection`, also provide the fresh selection guards. TSV input larger than the requested dimensions is rejected unless `--allow-truncate` is explicit. For merged or irregular tables, use `linear_cells[].index` with `set-cell --cell`; row and column insertion/deletion require a regular table. `swap-cell-text` moves text only, not formatting. Structural operations require Track Changes off unless `--allow-track-changes` is explicitly approved.

`set-cell` additionally returns `document`, `applied`, `verified`, `inspection_complete`, `readback`, and `errors`. A verified success compares the actual cell content with the requested text (normalizing CRLF/CR/LF and excluding only Word's terminal cell marker), then returns the existing full post-write fingerprint. `readback` includes `matches_requested`, actual character count and text hash without echoing the whole cell. For a subsequent operation on this verified target, the returned fingerprint can replace a redundant full query: the next mutation still recomputes and compares the live full fingerprint. Requery when a different scope or visual inspection is needed.

If a write completes but readback or formatting inspection fails, `set-cell` returns `applied:true`, `verified:false`, `fingerprint:null` and a nonzero exit. If the write itself throws, `applied:null` means its effect is unknown. Inspect the current document before retrying; neither case promises rollback or permits automatic replay.

`create-table`, `set-table-borders`, `set-cell-borders` and `normalize-table-borders` also return `document`, `applied`, `verified`, `inspection_complete` and `errors`, preserving their existing result fields. Any write, readback or final-inspection error returns a nonzero exit, `verified:false`, `inspection_complete:false` and `fingerprint:null`. A later inspection failure does not replace earlier write or rollback errors. Other mutation commands retain their existing contracts.

For `create-table`, `applied:true` means Word returned the created table, even if filling it later failed; `applied:null` means creation threw and its effect is unknown. `fill_complete` is true only when every requested in-bounds TSV write returned, false if filling did not finish, and null without TSV input. Verified success also checks actual dimensions, supplied cell text and the full fingerprint. A failed table-count read returns `table_count:null`. Partial tables are left for inspection; creation is not automatically replayed or rolled back.

For border commands, `applied:true` means all requested writes returned and their immediate readbacks passed, even if the final fingerprint failed. `applied:null` indicates uncertain or partial effects; `set-*-borders` can return `applied:false` when snapshot capture failed before any write. Check `failures` and `rollback_failures` as well as final-inspection `errors`. `rolled_back:true` records successful rollback calls (or no writes), not a transaction guarantee or permission to retry.

Border edges are `top`, `left`, `bottom`, `right`, `inside-h`, and `inside-v`; `outer` expands to the four outside edges and `all` selects every valid edge for the scope. Supported point widths are `0.25`, `0.5`, `0.75`, `1`, `1.5`, `2.25`, `3`, `4.5`, and `6`. Use `--style none` to remove selected edges. Reinspect after each change and use the new fingerprint. `layout_warnings` describe non-rectangular fallback. Fingerprint read failures produce `fingerprint:null` and `inspection_complete:false`; do not mutate from an incomplete inspection. Any `read_errors`, `failure_count`, rollback failures, or a nonzero exit mean the operation is not fully verified.

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

## Save, Export, and Close

```powershell
cscript //nologo $wc save-active --expect-path "<doc>" --yes
cscript //nologo $wc export-pdf --path "<preview.pdf>" --expect-path "<doc>" --yes
cscript //nologo $wc close-active --save --expect-path "<doc>" --yes
```

Use exactly one of `--save` or `--discard` with `close-active`. It closes only the verified active document and never quits Word.

## Troubleshooting

See [Recovery](recovery.md) for stale targets, incomplete inspections and blocked Word automation.

## Cleanup

Follow [task-artifact cleanup](recovery.md#cleanup); preserve source documents, backups and requested review outputs.

## Maintenance

See [Maintenance and validation](maintenance.md) for the canonical source, isolated tests and verified synchronization.
