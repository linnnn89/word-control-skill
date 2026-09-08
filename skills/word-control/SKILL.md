---
name: word-control
description: Control the active Microsoft Word desktop document on Windows through local COM automation. Use for narrow in-place DOCX work such as reading the active document or selection, replacing selected text, adding comments or tracked changes, precisely editing table cells, rows, columns, borders, and shading, creating or editing Word equations, saving a protected copy, or exporting a PDF. Prefer this skill over generating a replacement DOCX when the user asks to edit an open Word file directly.
---

# Word Control

Work directly in desktop Microsoft Word while preserving the user's existing document and session.

## Core Workflow

1. Run `status` and confirm the active document name, full path, save state, read-only state, and protection state.
2. For an important document, create a new versioned backup with `save-copy` before editing. If the document has unsaved changes and backup fails, do not silently save or use a stale disk copy; ask whether to save the source or proceed without a backup.
3. Inspect the exact target immediately before mutation:
   - Text or comments: run `selection-info` and retain `story_type`, `start`, `end`, and `text_hash`.
   - Tables: run `tables` and retain the target table's `fingerprint`.
   - Equations: run `equations` and retain the target equation's `fingerprint`.
4. Put non-ASCII input in a task-local UTF-8 file. Use an isolated temporary folder, not the document's source folder unless necessary.
5. Execute one narrow mutation with `--yes`, the expected document path or name, and the fresh selection or object fingerprint.
6. Re-run the relevant read-only inspection. A table or equation fingerprint changes after a successful edit, so inspect again before another mutation.
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

## Medical Manuscript Formatting Gate

When the active document is a medical manuscript, default to text-only editing. Do not add styling while inserting or replacing the title, headings, subheadings, or body text.

- Keep inserted manuscript text in an ordinary non-heading paragraph. Do not create Word heading/outline levels, collapsible sections, automatic numbering, multilevel lists, or added paragraph-number punctuation.
- Do not add bold, italics, underlining, font/size/color changes, alignment, indentation, paragraph spacing, line spacing, borders, shading, or page breaks unless the user explicitly requests that exact change.
- Preserve the formatting already attached to an edited range or paragraph. Do not normalize, restyle, or strip the rest of the document during a narrow text edit.
- Formatting commands are allowed for a manuscript only when the user explicitly requests them or explicitly requests pre-submission formatting for a named journal whose current mandatory requirements have been verified.
- In submission mode, apply only mandatory requirements. If a requirement is uncertain, leave the text plain and ask for or verify the journal instruction rather than guessing.

## Entry Point

```powershell
$wc = "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js"
cscript //nologo $wc help
```

Use `--output` for Chinese text or long output because the console code page can corrupt display text. Output must use `.json`, `.txt`, `.tsv`, or `.log`, must differ from input/document/export paths, and requires explicit `--overwrite` to replace an existing scratch file.

## Operation Routing

- Read document state: `status`, `selection`, `selection-info`, `document-text`, `paragraphs`, `tables`, `equations`.
- Edit prose: select the exact range, inspect it, then use `replace-selection`. Add `--track` for tracked manuscript edits.
- Add an author query: select the exact non-empty range, inspect it, then use `insert-comment`.
- Edit tables: inspect first, then use `set-cell`, `swap-cell-text`, row/column commands, targeted border/shading commands, `normalize-table-borders`, or `delete-table`. Reinspect after every operation because content and formatting both affect the fingerprint.
- Edit equations: validate conversion with `convert-equation`, inspect existing equations when relevant, then use `insert-equation`, `set-equation`, or `delete-equation`.
- Protect or review output: use `save-copy` with a new path or `export-pdf` with a new path.
- Save or close: use `save-active` or `close-active` only when explicitly requested. `close-active` never terminates Word.

For exact command forms, guard examples, and troubleshooting, read [references/usage.md](references/usage.md).

## Equations

The conservative LaTeX converter supports common fractions, square roots, integrals, sums, products, superscripts, subscripts, comparison symbols, and Greek letters. It rejects unknown LaTeX commands instead of inserting misleading plain text. Complex environments such as `align`, `matrix`, `cases`, and custom macros are unsupported; convert and inspect them manually before insertion.

Word equation rendering must be verified by reopening or inspecting the resulting OMath object and, when layout matters, exporting a PDF for visual review.

## Tables

TSV dimensions must fit the requested table. Use `--allow-truncate` only when the user explicitly accepts dropped cells. Merged or irregular tables may not expose a rectangular row/column model; inspect their `linear_cells` and use `set-cell --cell`. Do not insert or delete whole rows or columns in an irregular table.

Use `swap-cell-text` only to exchange textual contents while preserving each cell's formatting. Use explicit row/column insertion and deletion commands for structural changes; they reject deletion of the final row or column. Keep Track Changes off for deterministic structural changes unless the user explicitly approves `--allow-track-changes`.

Use `set-cell-shading` for a solid RGB background. Use `set-cell-borders` for selected cell edges and `set-table-borders` for outer or internal table edges. Supported border styles are `none`, `single`, `dotted`, `dashed`, `dash-large`, `dash-dot`, `dash-dot-dot`, `double`, and `triple`; supported widths are the Word-native point values listed in the command guide. Formatting changes are included in the table fingerprint. `layout_warnings` document fallback behavior, while a null fingerprint, `inspection_complete:false`, any `read_errors`, `failures`, or rollback failures mean the operation is incomplete and must not be reported as successful.

## Verification

The `.codex` copy is the maintenance source. Validate it before copying the complete skill to `.agents` and `.grok`; do not maintain divergent instructions in those copies.

Run the isolated test after changing this skill (Windows desktop Word and Node.js are required for testing; ordinary commands still use only Windows Script Host):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\.codex\skills\word-control\scripts\test_word_control.ps1"
```

The test runs pure regressions, creates its own hidden Word instance and a unique temporary DOCX, and strictly parses every generated JSON output before removing its artifacts. Existing Word sessions cause command integration to be skipped; a skipped result does not authorize synchronization. It must not attach to or close the user's Word session.

## Limitations

- Windows desktop Word only; this is not a Word Online or cross-platform Office.js add-in.
- Whole-document rewrites and complex layout reconstruction are outside the intended scope.
- Comments and tracked changes can be created or toggled, but the script does not yet provide a full review-history report.
- For new documents or structural DOCX generation, use the document-generation skill and use Word Control only for final in-place review or narrow corrections.
