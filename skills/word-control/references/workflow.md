# Editing workflow

Work directly in the verified active Word document. Read the relevant command-guide section for exact options.

New-document generation and whole-document reconstruction are outside this skill's scope.

For discovery, use `find-text` in the relevant story, bounded paragraph pages or table summaries. For table values, use `tables --detail text`, with `--table`, `--cell-from` and `--cell-max` when only a cell range matters. Use `equations --index N` for one formula. Before editing, inspect only the target in full. Full inspection of every table still reads every cell and its guarded formatting; use it when the task actually needs that scope.

`document-text` reads Word's main text story, and `--at end` uses its end. Endnotes can render after that point: inspect notes separately when they are relevant, and verify the visible insertion location instead of inferring the last rendered page from the range offset.

## Core Workflow

1. Run `status` and confirm the active document name, full path, save state, read-only state, and protection state.
2. For an important document, create a new versioned backup with `save-copy` before editing. If the document has unsaved changes and backup fails, do not silently save or use a stale disk copy; ask whether to save the source or proceed without a backup.
3. Inspect the exact target immediately before mutation:
   - Text or comments: run `selection-info` and retain `story_type`, `start`, `end`, and `text_hash`.
   - Tables: run `tables --table N` in full detail and retain the target table's `fingerprint`.
   - Equations: run `equations --index N` and retain the target equation's `fingerprint`.
4. Put non-ASCII input in a task-local UTF-8 file. Use an isolated temporary folder, not the document's source folder unless necessary.
5. Execute one narrow mutation with `--yes`, the expected document path or name, and the fresh selection or object fingerprint.
6. Verify the result. A successful `set-cell` with `verified:true`, matching document/target, `readback.matches_requested:true` and a non-null fingerprint already includes actual text readback and full post-write inspection; reuse this fresh fingerprint for the next operation without an extra full query. The next mutation still recomputes the live fingerprint. Otherwise re-run the relevant target inspection. Reinspect for additional scope or visual evidence, and never reuse a failed or incomplete result.
7. Save, close, or export only when the user requested it. Never quit the user's Word application.
8. Remove task-only JSON, TXT, TSV, and probe artifacts. Preserve source documents, backups, and requested review outputs.

## Non-Negotiable Guards

- Require `--expect-path` for saved documents. Use `--expect-name` only for an intentionally unsaved document.
- Require all four selection values from one fresh `selection-info` result: `--expect-story-type`, `--expect-start`, `--expect-end`, and `--expect-selection-hash`.
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
