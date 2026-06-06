---
name: word-control
description: Use when the user wants Codex to work directly with Microsoft Word desktop on Windows, especially to read the active Word document or selection, replace selected text, insert comments, enable tracked changes, save a copy, export PDF, or avoid creating a new DOCX from scratch. This skill uses local Word COM automation through Windows Script Host and should prefer safe selection-based edits with explicit backups for unpublished or important documents.
---

# Word Control

## Mission

Control the local Microsoft Word desktop application directly when the user wants edits inside an open `.docx` instead of generating a new document file.

Use this skill for current-document workflows:

- Read the active Word document or selected text.
- Replace the current Word selection after revising text.
- Enable or disable Track Changes.
- Insert comments on the current selection.
- Create, inspect, edit, or delete simple Word tables.
- Normalize table borders when Word shows inconsistent black edges around merged cells.
- Create, inspect, edit, or delete Word equation objects from Word linear syntax or common LaTeX syntax.
- Save or close the active document when explicitly requested.
- Save a copy before edits.
- Export the active document to PDF for visual QA.

## Safety Rules

- Before mutating a user document, make a backup with `save-copy` unless the user explicitly says not to.
- Prefer editing the current selection, not whole-document replacement.
- For manuscript editing, enable Track Changes before replacement unless the user asks for clean edits.
- If the selection is empty, treat replacement as insertion at the cursor and say so.
- Do not close Word or kill Word processes unless they were created by a smoke test and have no visible document title.
- Do not use this for cloud-only Word Online documents unless they are open in desktop Word and COM can access them.
- At task close, clean Word Control scratch artifacts created only for the task, such as probe/status `*.json`, temporary input/output `*.txt`, equation/table scratch files, and other transient files. Preserve source documents, edited documents, explicit backups, and final review artifacts such as PDF/PNG previews unless the user explicitly asks to remove them.

## Script

Main script:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" help
```

Use UTF-8 input and output files for non-ASCII text. Prefer `--output` over console output when reading Chinese text or long manuscript passages.

## Common Commands

Status:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" status --output status.json
```

Read current selection:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" selection --output selection.txt
```

Replace current selection using tracked changes:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" replace-selection --input revised.txt --track --yes
```

Insert a comment on the current selection:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" insert-comment --input comment.txt --yes
```

Inspect tables:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" tables --output tables.json
```

Create a table at the current selection or cursor:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" create-table --rows 3 --cols 3 --input table.tsv --yes
```

Edit a cell:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" set-cell --table 1 --row 2 --col 2 --input cell.txt --yes
```

Delete a table:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" delete-table --table 1 --yes
```

Normalize a table's borders to the light gray grid style:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" normalize-table-borders --table 1 --color D9DEE8 --line-width 4 --yes
```

Inspect equations:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" equations --output equations.json
```

Insert a LaTeX-like equation:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" insert-equation --input equation.txt --format latex --at end --yes
```

Edit an existing equation:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" set-equation --index 1 --input equation.txt --format latex --yes
```

Delete an equation:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" delete-equation --index 1 --yes
```

Save the active document:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" save-active --yes
```

Close the active document:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" close-active --save --quit-if-empty --yes
```

Save a copy:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" save-copy --path backup.docx --yes
```

Export PDF:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" export-pdf --path preview.pdf --yes
```

## Editing Workflow

1. Ask the user to open the target `.docx` in desktop Word and select the passage to edit.
2. Run `status` and `selection`.
3. Save a backup with `save-copy`.
4. Use the relevant writing/editing skill to revise the selected text.
5. Write revised text to a UTF-8 temporary file.
6. Run `replace-selection --track --yes`.
7. Insert comments for unresolved evidence, wording, statistics, citations, or author decisions.
8. Export PDF if layout verification matters.
9. Clean task-local scratch files created by Word Control (`status.json`, `selection.txt`, table/equation probe files, temporary revision text, and similar intermediate `*.json`/`*.txt` artifacts). Do not delete user documents, backups, or final verification previews.

## Cleanup Standard

Before final response, list the Word Control scratch artifacts created during the task and remove those that are purely intermediate. Keep anything needed for audit, rollback, or user review:

- Keep original `.docx` files and edited final documents.
- Keep explicit backups made with `save-copy`.
- Keep final PDF/PNG previews when they document visual QA.
- Remove temporary status, selection, table, paragraph, equation, and revision files unless the user asked to keep them.
- If a file's role is ambiguous, keep it and mention it instead of deleting it silently.

## Paragraph Replacement Caution

Known local issue: `replace-paragraph --index n` may behave like an insertion-plus-blanking operation in some Word documents rather than a clean in-place replacement. It can leave the old paragraph later in the document, create extra blank paragraphs, and shift subsequent paragraph indexes. Do not use `replace-paragraph` to delete paragraphs by replacing them with a blank file.

Safer rules:

- Prefer `replace-selection` for targeted prose edits after selecting the exact range in Word.
- If `replace-paragraph` is used, rerun `paragraphs` immediately after each mutation and verify with `document-text --full` before making another indexed edit.
- For generated DOCX drafts, prefer regenerating a clean DOCX, then use Word Control for `open`, `save-copy`, `export-pdf`, and text/layout verification rather than chaining many indexed paragraph replacements.
- Always check for duplicated old text after paragraph operations, especially when correcting a heading/body pair or after clearing an earlier paragraph.

## Limitations

- The current implementation uses Windows Script Host JScript because this machine's Python `pywin32` install path was not reliable at setup time.
- It controls the desktop Word application on Windows; it is not a cross-platform Office.js add-in.
- It does not yet provide a Word task-pane UI.
- Whole-document rewrites are intentionally not the default because they are easier to damage.
- Equation support uses Word's OMath engine. Raw LaTeX is translated through a conservative converter for common constructs such as `\frac{}`, `\sqrt{}`, `\int`, `\sum`, superscripts, subscripts, and Greek letters. Complex LaTeX environments such as `align`, `matrix`, or custom macros are not yet supported.
