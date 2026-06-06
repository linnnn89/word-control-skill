# Word Control Usage

## Recommended Human Workflow

1. Open the `.docx` in Microsoft Word desktop.
2. Select the passage you want Codex to revise.
3. Ask Codex to read the selection.
4. Codex saves a backup copy.
5. Codex revises the passage using the relevant writing skill.
6. Codex replaces the selection with Track Changes enabled.
7. Codex adds comments where facts, citations, numbers, or author intent need confirmation.

## Command Reference

All commands are run through Windows Script Host:

```powershell
cscript //nologo path\to\word_control.js <command> [options]
```

Read-only commands:

- `help`
- `status`
- `selection`
- `document-text`
- `paragraphs`
- `tables`
- `equations`

Mutation commands require `--yes`:

- `enable-track-changes --yes`
- `disable-track-changes --yes`
- `replace-selection --input file --yes`
- `replace-paragraph --index n --input file --yes`
- `insert-comment --input file --yes`
- `create-table --rows n --cols n --input table.tsv --yes`
- `set-cell --table n --row n --col n --input file --yes`
- `delete-table --table n --yes`
- `normalize-table-borders --table n --color D9DEE8 --line-width 4 --yes`
- `insert-equation --input file --format latex --yes`
- `set-equation --index n --input file --format latex --yes`
- `delete-equation --index n --yes`
- `save-active --yes`
- `close-active --save --yes`
- `close-active --discard --yes`
- `save-copy --path file --yes`
- `export-pdf --path file --yes`

## Encoding

For Chinese or manuscript text, always pass text through UTF-8 files:

- Use `--input revised.txt` for replacement or comments.
- Use `--output selection.txt` for selected text.

Console output may display non-ASCII text incorrectly depending on the active Windows code page.

## Safe Editing Pattern

For important manuscripts:

```powershell
cscript //nologo word_control.js status --output status.json
cscript //nologo word_control.js selection --output selection.txt
cscript //nologo word_control.js save-copy --path manuscript.backup.docx --yes
cscript //nologo word_control.js replace-selection --input revised.txt --track --yes
```

## Troubleshooting

- If `no running Word instance found`, open Word desktop first.
- If `Word is running but no document is open`, open a `.docx`.
- If replacement inserts text at the cursor, the Word selection was collapsed.
- If Word Online is open only in a browser, this script cannot access it.
- If Word prompts block automation, save and close dialogs manually, then retry.

## Equation Notes

Word does not reliably build every raw LaTeX command through COM. The script converts common LaTeX syntax into Word linear equation syntax before calling Word's OMath `BuildUp`.

Supported common patterns:

- `x^2 + y^2 = z^2`
- `\frac{a+b}{c}`
- `\sqrt{a+b}`
- `\int_0^1 x^2 dx`
- `\sum_{i=1}^n x_i`
- Greek letters such as `\alpha`, `\beta`, `\mu`, `\Delta`

Not yet supported:

- `align`, `matrix`, `cases`
- custom macros
- deeply nested LaTeX with complex brace structure
