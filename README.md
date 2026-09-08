# Word Control Skill

## English

Word Control is a Codex skill for narrow, in-place editing of the active Microsoft Word desktop document on Windows. It uses local Word COM automation so an agent can inspect and edit the document already open in Word instead of regenerating a DOCX.

Typical tasks include:

- inspect the active document, selection, paragraphs, tables, and equations;
- replace selected text, add comments, and make tracked changes;
- edit table cells, swap cell contents, add or remove rows and columns, and change borders or shading;
- insert, replace, and remove Word equations from supported LaTeX or linear input;
- create a protected copy, save the active document, or export a PDF.

The skill is intentionally conservative. Mutations require explicit confirmation plus a fresh document path and, where applicable, a selection hash or table/equation fingerprint. Existing backups and exports are not overwritten by default, and closing a document never quits the user's Word application.

## 中文

Word Control 是一个面向 Windows 桌面版 Microsoft Word 的 Codex 技能。它通过本地 Word 自动化，对当前已经打开的文档进行小范围、原位修改，避免为了简单编辑而重新生成整个 DOCX 文件。

主要能力包括：

- 检查当前文档、选区、段落、表格和公式；
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
|-- references/usage.md
`-- scripts/
    |-- word_control.js
    |-- test_word_control.ps1
    |-- test_word_control_integration.js
    `-- test_word_control_pure.cjs
```

## Installation

Copy `skills/word-control` into the Codex skills directory:

```powershell
Copy-Item -Recurse -Force ".\skills\word-control" "$env:USERPROFILE\.codex\skills\word-control"
```

Restart or refresh Codex so it discovers the installed skill.

## Basic Use

```powershell
$wc = "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js"
cscript //nologo $wc status --output status.json
cscript //nologo $wc selection-info --output selection-info.json
cscript //nologo $wc tables --output tables.json
cscript //nologo $wc equations --output equations.json
```

Before changing content, use the active document path returned by `status` and the fresh selection or object fingerprint returned by the relevant inspection command. See [the command guide](skills/word-control/references/usage.md) for guarded editing examples and cleanup requirements.

## Guard compatibility

Selection mutations now require `--expect-story-type` from a fresh `selection-info` result, alongside start, end, and text hash. Saved documents require `--expect-path`; filename-only verification is reserved for unsaved documents.

Scratch `--output` files must use `.json`, `.txt`, `.tsv`, or `.log`, must differ from input/document/export paths, and need explicit `--overwrite` to replace an existing file. Incomplete table inspection returns a null fingerprint and must not be used for mutation.

## Validation

The test runs pure regression checks, creates a separate hidden Word instance and temporary DOCX, exercises text, table, equation, save-copy, and PDF operations, and validates every JSON output with a strict parser. It removes its own artifacts. If Word is already running, command integration is skipped; a skipped result is not a complete validation.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\skills\word-control\scripts\test_word_control.ps1"
```

Do not run document automation tests against an irreplaceable source file. Use a copy and a separate temporary output directory.

## Limitations

- Intended for narrow edits, not whole-document reconstruction.
- Complex LaTeX environments and custom macros require manual conversion and visual verification.
- Merged or irregular tables have restricted structural editing support.
- Visible Word dialogs can block automation and must be resolved manually.

## License

No open-source license is currently granted. All rights reserved by the repository owner.
