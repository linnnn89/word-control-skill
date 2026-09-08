# Word Control Command Guide

Set the entry point once per PowerShell session:

```powershell
$wc = "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js"
```

## Inspect First

```powershell
cscript //nologo $wc status --output status.json
cscript //nologo $wc selection-info --output selection-info.json
cscript //nologo $wc tables --max 25 --output tables.json
cscript //nologo $wc equations --output equations.json
```

Read-only commands are `help`, `status`, `selection`, `selection-info`, `document-text`, `paragraphs`, `tables`, `equations`, and `convert-equation`.

Use the active document path returned by `status` as `<doc>`. For an unsaved document, use its exact name with `--expect-name` instead.

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

## Tables

Use the target fingerprint from a fresh `tables` result:

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

Unknown LaTeX commands are rejected. Use `--format linear` only after checking Word linear syntax.

## Save, Export, and Close

```powershell
cscript //nologo $wc save-active --expect-path "<doc>" --yes
cscript //nologo $wc export-pdf --path "<preview.pdf>" --expect-path "<doc>" --yes
cscript //nologo $wc close-active --save --expect-path "<doc>" --yes
```

Use exactly one of `--save` or `--discard` with `close-active`. It closes only the verified active document and never quits Word.

## Troubleshooting

- `no running Word instance found`: open the document in desktop Word.
- `active document path mismatch`: bring the intended document to the front and rerun `status`.
- `selection ... changed`: rerun `selection-info`; do not reuse old coordinates or hashes.
- `target fingerprint changed`: rerun `tables` or `equations`; indexes alone are not stable identifiers.
- `selection is collapsed`: select text, or explicitly approve cursor insertion.
- Word dialog blocks automation: resolve the visible dialog manually, then rerun inspection.
- Non-ASCII console text is corrupted: use UTF-8 input files and `--output` files.

## Cleanup

Delete only task-generated status, selection, table, equation, conversion, and input scratch files after verification. Keep the original DOCX, explicit backups, requested PDF/PNG previews, and anything needed for rollback or audit.

## Maintenance

The `.codex` installation is the source of truth. Back up all installations, edit and fully test `.codex`, then copy its files to `.agents` and `.grok` and compare SHA-256 hashes. The copied command examples intentionally continue to target the canonical `.codex` entry point. Do not synchronize after failed or skipped integration tests.
