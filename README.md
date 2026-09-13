# Word Control Skill

## English

Word Control is a Codex skill for narrow, in-place editing of the active Microsoft Word desktop document on Windows. It uses local Word COM automation so an agent can inspect and edit the document already open in Word instead of regenerating a DOCX.

Typical tasks include:

- inspect the active document and selection, find literal text in the main story or notes, and read paragraph/cell pages or targeted tables/equations;
- replace selected text, add comments, and make tracked changes;
- edit table cells, swap cell contents, add or remove rows and columns, and change borders or shading;
- insert, replace, and remove Word equations from supported LaTeX or linear input;
- create a protected copy, save the active document, or export a PDF.

The skill is intentionally conservative. Mutations require explicit confirmation plus a fresh document path and, where applicable, a selection hash or table/equation fingerprint. Existing backups and exports are not overwritten by default, and closing a document never quits the user's Word application.

## 中文

Word Control 是一个面向 Windows 桌面版 Microsoft Word 的 Codex 技能。它通过本地 Word 自动化，对当前已经打开的文档进行小范围、原位修改，避免为了简单编辑而重新生成整个 DOCX 文件。

主要能力包括：

- 检查当前文档和选区，定位正文、脚注或尾注中的文字，按需读取段落、单元格分页及指定表格或公式；
- 替换选中文字、添加批注和使用修订模式修改；
- 精细修改表格单元格，交换内容，增删行列，更改边框和底色；
- 根据受支持的 LaTeX 或线性输入插入、替换和删除 Word 内置公式；
- 创建保护性副本、保存当前文档或导出 PDF。

本技能默认采用保守的安全策略。写操作除了明确确认外，还需要核对当前文档路径；选区、表格和公式操作还会校验刚刚取得的哈希或指纹。已有备份和导出文件默认不会被覆盖，关闭文档也不会退出用户正在使用的 Word。

## Requirements

- Windows
- Microsoft Word desktop
- Windows Script Host (`cscript.exe`)
- PowerShell and Node.js for the test suite (ordinary Word commands do not require Node.js)

This is not an Office.js add-in and does not control Word Online.

## Repository Layout

```text
skills/word-control/
|-- SKILL.md
|-- agents/openai.yaml
|-- references/
|   |-- workflow.md
|   |-- usage.md
|   |-- recovery.md
|   `-- maintenance.md
`-- scripts/
    |-- word_control.js
    |-- test_word_control.ps1
    |-- test_word_control_integration.js
    `-- test_word_control_pure.cjs
```

## Installation

Use `.agents/skills/word-control` as the only global installation. Keep backups outside skill discovery directories and avoid installing another global copy.

Copy `skills/word-control` into the global skills directory:

```powershell
Copy-Item -Recurse -Force ".\skills\word-control" "$env:USERPROFILE\.agents\skills\word-control"
```

Restart or refresh Codex so it discovers the installed skill.

## Basic Use

```powershell
$wc = "$env:USERPROFILE\.agents\skills\word-control\scripts\word_control.js"
cscript //nologo $wc status --output status.json
cscript //nologo $wc selection-info --output selection-info.json
cscript //nologo $wc tables --detail summary --output tables-summary.json
cscript //nologo $wc tables --detail text --output tables-text.json
cscript //nologo $wc tables --table 2 --output table-2.json
cscript //nologo $wc paragraphs --from 81 --max 40 --output paragraphs-81.json
cscript //nologo $wc find-text --input query.txt --max 5 --output matches.json
cscript //nologo $wc tables --table 2 --detail text --cell-from 5 --cell-max 10 --output cells.json
cscript //nologo $wc equations --index 1 --output equation-1.json
```

Choose a table index from the current summary; `2` above is an example. `--table N` inspects only that table and cannot be combined with `--max`. Full detail remains the default. Summary mode reads dimensions and cell counts without reading cell text or formatting; it deliberately returns a null fingerprint, `inspection_complete:false` and an unverified layout. Obtain a full inspection of the target before editing.

Use `--detail text` to read table values without the expensive format fingerprint pass. It preserves regular and merged-cell text output, with `fingerprint:null` and `inspection_complete:false`. Summary and text modes cannot authorize mutations; use full detail for that step.

Cell pagination requires a single table and text detail. It returns absolute linear cell indices, the table's whole cell count and `next_cell`; it does not assume a rectangular layout. `find-text` searches case-sensitive literal input from a UTF-8 file without a trailing newline. Choose `--story main`, `footnotes` or `endnotes`, and use bounded `--max`/`--context` plus `next_from` to continue. Results identify the searched story; they do not claim header/footer/text-box coverage or provide mutation guards. Search preserves the selection and the Find settings it changes. See the command guide for limits and partial-failure semantics.

`set-cell` now verifies actual cell text and returns `applied`, `verified`, `readback`, document identity and a full post-write fingerprint. A verified result can supply the next operation's expected fingerprint without an extra full query; the next write still compares it against a newly computed live fingerprint. A failed write/verification returns a nonzero exit and no reusable fingerprint, distinguishing completed-but-unverified writes from unknown effects. Inspect before retrying.

Table creation and border commands also preserve operation state when a later check fails. `create-table` reports whether a table was created and whether TSV filling completed; verified success checks its dimensions, supplied text and full fingerprint. Border commands retain write and rollback failures alongside final-inspection errors. Failed results have `verified:false` and `fingerprint:null`, even if part of the operation succeeded. Inspect before retrying; these results do not promise automatic rollback or replay. Other mutation commands keep their existing result contracts.

Paragraph indices are 1-based. `--from N --max count` returns the requested page with absolute indices and `next_from` (null at the end). Starting beyond the end returns no paragraphs. Document edits can shift indices. Successful calls without the new options retain their defaults and output shapes. Failed paragraph/equation reads are marked with null text and explicit read errors; an incomplete response must not be treated as complete document evidence. Equation text and its fingerprint use one snapshot.

Before changing content, read [the editing workflow](skills/word-control/references/workflow.md), use the active document path returned by `status`, and obtain fresh selection or full object guards. The short [Skill entry](skills/word-control/SKILL.md) routes to the [command guide](skills/word-control/references/usage.md), [recovery guide](skills/word-control/references/recovery.md) and [maintenance guide](skills/word-control/references/maintenance.md) as needed.

## Guard compatibility

Selection mutations now require `--expect-story-type` from a fresh `selection-info` result, alongside start, end, and text hash. Saved documents require `--expect-path`; filename-only verification is reserved for unsaved documents.

Scratch `--output` files must use `.json`, `.txt`, `.tsv`, or `.log`, must differ from input/document/export paths, and need explicit `--overwrite` to replace an existing file. Incomplete table inspection returns a null fingerprint and must not be used for mutation.

## Validation

The test runs pure regression checks, creates a separate hidden Word instance and temporary DOCX, exercises text, table, equation, save-copy, and PDF operations, and validates every JSON output with a strict parser. Scoped-query tests cover paragraph pages, regular and merged tables, summary guard omission, and unchanged document/selection state. It removes its own artifacts. If Word is already running, command integration is skipped; a skipped result is not a complete validation.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\skills\word-control\scripts\test_word_control.ps1"
```

Do not run document automation tests against an irreplaceable source file. Use a copy and a separate temporary output directory.

For a local paper, pass `-FixturePath` to the same test command. The test copies the source, verifies its SHA-256 remains unchanged, exercises guarded edits on appended test content, exports a PDF, and reopens the saved copy to verify the original body text and object counts. It restores the Word options it changes. The repository ignores `/测试用WORD/`, including its contents; keep private fixtures there. Fixture test artifacts are removed on completion, so preserve separate review outputs when visual comparison is needed.

## Limitations

- Intended for narrow edits, not whole-document reconstruction.
- Complex LaTeX environments and custom macros require manual conversion and visual verification.
- Merged or irregular tables have restricted structural editing support.
- Visible Word dialogs can block automation and must be resolved manually.

## License

No open-source license is currently granted. All rights reserved by the repository owner.
