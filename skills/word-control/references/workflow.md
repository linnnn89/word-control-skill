# Editing workflow

Work directly in the verified active Word document. Read the relevant command-guide section for exact options.

New-document generation and whole-document reconstruction are outside this skill's scope.

For discovery, use `find-text` in the relevant story, bounded paragraph pages or table summaries. For table values, use `tables --detail text`, with `--table`, `--cell-from` and `--cell-max` when only a cell range matters. Use `equations --index N` for one formula. Before editing, inspect only the target in full. Full inspection of every table still reads every cell and its guarded formatting; use it when the task actually needs that scope.

`document-text` reads Word's main text story, and `--at end` uses its end. Endnotes can render after that point: inspect notes separately when they are relevant, and verify the visible insertion location instead of inferring the last rendered page from the range offset.

## Core Workflow

1. Run `status` and confirm the active document name, full path, save state, read-only state, and protection state.
2. Establish a verified backup before the first edit, following [Backup before editing](#backup-before-editing).
3. Inspect the exact target immediately before mutation:
   - Text or comments: run `selection-info`, require `selection.range_edit_supported:true`, and retain `story_type`, `start`, `end`, and `text_hash`.
   - Tables: run `tables --table N` in full detail and retain the target table's `fingerprint`.
   - Equations: run `equations --index N` and retain the target equation's `fingerprint`.
4. Put non-ASCII input in a task-local UTF-8 file. Use an isolated temporary folder, not the document's source folder unless necessary.
5. Execute one narrow mutation with `--yes`, the expected document path or name, and the fresh selection or object fingerprint.
6. Verify the result. Cell text/shading, row/column insertion/deletion and border commands support result reuse under their [documented result contracts](usage.md#tables): `ok`, `applied`, `verified` and `inspection_complete` must be true, document/target must match, all returned error/failure arrays must be empty and the fingerprint must be non-null. Cell text/shading and row/column commands also require `readback.matches_requested:true`. Reuse that fingerprint without an extra full query; the next mutation still recomputes the live full fingerprint. After structural changes, recalculate or reinspect affected row/column/cell indices before selecting the next target. Reinspect for additional scope or visual evidence, and never reuse a failed or incomplete result. After `create-table`, identify the new table separately; `table_count` alone is not its index.
7. Save, close, or export only when the user requested it. Never quit the user's Word application.
8. Remove task-only JSON, TXT, TSV, and probe artifacts. Preserve source documents, backups, and requested review outputs.

For acceptance, compare before/after state across the affected scope: the target, shared cell boundaries or neighboring content, and relevant document structure. A verified mutation result covers the command's own checks; it does not establish unchanged formatting, bookmarks or layout everywhere else. Reuse its guard while performing any additional checks required by the task, and investigate unexpected differences before saving or continuing dependent edits.

For save/backup/export failures, also compare the source file, existing destination, document content and save/open state. Save commands require an empty background-save queue, a saved document and a nonempty file; this is not a substitute for content readback. Preserve any reported recovery file or cleanup warning. Prefer a new backup name, and do not discard or close the document after a failed save.

## Backup before editing

Before the first content, formatting, comment, revision or structural change to each user document in an editing task, use [`save-copy`](usage.md#back-up) to create a separate, versioned backup of its current state. Use the source folder or a user-designated backup folder, preserve the source format, and choose an unused name such as `report.before-edit-YYYYMMDD-HHMMSS.docx`. Keep the active document on its original path; a Save As that changes the editing target is not this backup operation.

Require exit code 0, `ok:true`, `includes_current_document_state:true`, and a returned `path` matching the intended backup. Confirm that file exists and is nonempty before editing. Record the backup path and the document state it covers in the task context; do not infer a valid backup from a filename alone. Preserve and inspect any `cleanup_warning` as described in [recovery](recovery.md).

Reuse this baseline backup for successive known edits to the same document within the task; do not copy the whole document before every cell or text change. Start a new baseline for a new editing task or after changes outside the verified edit sequence. Each document needs its own backup. Read-only inspection and task-created disposable test documents do not need this step.

If the backup fails, is unverified or cannot include unsaved changes, stop dependent edits and report the reason. Do not silently save over the source or substitute an older disk copy. If the existing request explicitly authorizes saving the source, carry out that save, verify it and retry the backup; otherwise obtain the user's decision. An existing explicit waiver permits skipping the backup only for its agreed scope, without asking again. It does not waive document/target guards or affected-scope acceptance.

This is an instruction for the agent's editing workflow; mutation commands do not automatically create or enforce a backup.

## Non-Negotiable Guards

- Require `--expect-path` for saved documents. Use `--expect-name` only for an intentionally unsaved document.
- Require `selection.range_edit_supported:true` and all four selection values from one fresh `selection-info` result: `--expect-story-type`, `--expect-start`, `--expect-end`, and `--expect-selection-hash`. For tables/equations inserted at the selection, apply the same check. Unsupported selection kinds cannot be bypassed with an override; use the [selection guidance](usage.md#selection-changes) to choose a supported target.
- Require `--expect-table-fingerprint` or `--expect-equation-fingerprint` from a fresh inspection before indexed object changes.
- Treat `--allow-active`, `--allow-unverified-selection`, `--allow-unverified-target`, and `--allow-insert` as exceptional overrides. Use them only after manual target verification and only when their specific behavior is intended.
- Refuse an existing backup, PDF, or smoke-test output unless the user explicitly approved replacement and `--overwrite` is passed.
- Do not use `replace-paragraph`; it is disabled because Word paragraph ranges can duplicate or shift content. Select the exact range and use `replace-selection`.
- Do not infer that a collapsed selection is intended. Insertion at the cursor requires explicit `--allow-insert`.
- Do not close documents, dismiss dialogs, or alter Track Changes unless required by the user's request.

Table summary/text modes are read-only results: `fingerprint:null` and `inspection_complete:false` are intentional; summaries also use `layout:"unverified"`. Before a table mutation, use `tables --table N` in full detail and require a complete inspection. Never turn an omitted formatting check into an unsafe-override justification.

A successful `--track` replacement leaves Track Changes on. A failed replacement attempts to restore the previous state; report any restoration failure. Preserve partial-operation and readback failures instead of describing them as success.

## Medical Manuscript Formatting Gate

When the active document is a medical manuscript, default to text-only editing. Do not add styling while inserting or replacing the title, headings, subheadings, or body text.

- Keep inserted manuscript text in an ordinary non-heading paragraph. Do not create Word heading/outline levels, collapsible sections, automatic numbering, multilevel lists, or added paragraph-number punctuation.
- Do not add bold, italics, underlining, font/size/color changes, alignment, indentation, paragraph spacing, line spacing, borders, shading, or page breaks unless the user explicitly requests that exact change.
- Preserve the formatting already attached to an edited range or paragraph. Do not normalize, restyle, or strip the rest of the document during a narrow text edit.
- Formatting commands are allowed for a manuscript only when the user explicitly requests them or explicitly requests pre-submission formatting for a named journal whose current mandatory requirements have been verified.
- In submission mode, apply only mandatory requirements. If a requirement is uncertain, leave the text plain and ask for or verify the journal instruction rather than guessing.
