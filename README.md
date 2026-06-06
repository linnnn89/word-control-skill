# Word Control Skill

Direct Microsoft Word desktop control for Codex on Windows.

This public repository contains a Codex skill that lets Codex work with an open Microsoft Word document through the local Windows desktop Word application. It is designed for small, practical, in-place document edits instead of repeatedly generating new `.docx` files from scratch.

本公开仓库提供一个 Codex skill，用于在 Windows 桌面版 Microsoft Word 中直接操作当前打开的 `.docx` 文档。它适合小范围、实用、就地修改，而不是每次都重新生成一个新的 Word 文件。

## What It Does

- Read the active Word document, current selection, paragraphs, tables, and equations.
- Replace the current selection with revised text.
- Enable or disable Track Changes.
- Insert comments on the selected text.
- Create, inspect, edit, normalize, and delete Word tables.
- Insert, edit, inspect, and delete Word equation objects from Word linear syntax or common LaTeX syntax.
- Save the active document, save a backup copy, close the active document, and export PDF for visual review.
- Clean task-only scratch files such as temporary `*.json` and `*.txt` probe outputs after the Word edit is complete.

## 功能概览

- 读取当前 Word 文档、当前选区、段落、表格和公式。
- 将润色或修改后的文本替换回当前选区。
- 开启或关闭修订模式。
- 给当前选区插入批注。
- 创建、查看、编辑、统一边框和删除 Word 内置表格。
- 从 Word 线性公式或常见 LaTeX 写法创建、编辑、查看和删除 Word 内置公式对象。
- 保存当前文档、保存备份副本、关闭当前文档，并导出 PDF 进行视觉检查。
- 完成 Word 修改后清理仅用于本次任务的临时 `*.json`、`*.txt` 等中间产物。

## Why This Exists

For local Word editing, the fastest workflow is often:

1. Open the target `.docx` in desktop Word.
2. Select the passage, table, or location to edit.
3. Ask Codex to use `$word-control`.
4. Let Codex save a backup, apply a narrow change, and export a PDF preview.
5. Let Codex remove task-only scratch files while preserving the document, backup, and final review previews.

This avoids unnecessary token use, temporary folders, and repeated new-document generation when the task is a small edit to an existing Word file.

## 设计目的

对于本机 Word 文件的小修改，更高效的工作流通常是：

1. 在桌面版 Word 中打开目标 `.docx`。
2. 选中需要修改的文字、表格或插入位置。
3. 让 Codex 使用 `$word-control`。
4. 由 Codex 先备份，再做局部修改，最后导出 PDF 预览。
5. 由 Codex 清理仅用于本次任务的临时中间产物，同时保留文档、备份和最终预览文件。

这样可以减少 token 消耗、临时文件夹堆积，以及反复生成新 docx 的麻烦。

## Requirements

- Windows.
- Microsoft Word desktop installed.
- Windows Script Host enabled (`cscript.exe`).
- Codex local skills directory.

The current implementation uses Windows Script Host JScript and Word COM automation. It does not require `pywin32`, Node.js, npm, or an Office.js task pane.

## 系统要求

- Windows。
- 已安装桌面版 Microsoft Word。
- Windows Script Host 可用，即 `cscript.exe` 可运行。
- 本地 Codex skills 目录。

当前实现使用 Windows Script Host JScript 调用 Word COM 自动化，不依赖 `pywin32`、Node.js、npm 或 Office.js 侧边栏插件。

## Repository Structure

```text
skills/
  word-control/
    SKILL.md
    agents/
      openai.yaml
    references/
      usage.md
    scripts/
      word_control.js
```

## Installation

Copy the skill folder into your Codex skills directory:

```powershell
Copy-Item -Recurse `
  .\skills\word-control `
  "$env:USERPROFILE\.codex\skills\word-control"
```

Then restart Codex so the skill list can refresh.

## 安装方式

将 skill 文件夹复制到本机 Codex skills 目录：

```powershell
Copy-Item -Recurse `
  .\skills\word-control `
  "$env:USERPROFILE\.codex\skills\word-control"
```

然后重启 Codex，让 skill 列表刷新。

## Usage Examples

Read the current selection:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" selection --output selection.txt
```

Replace the current selection with Track Changes enabled:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" replace-selection --input revised.txt --track --yes
```

Normalize a table's borders to a light gray grid:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" normalize-table-borders --table 1 --color D9DEE8 --line-width 4 --yes
```

Insert a LaTeX-like equation:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" insert-equation --input equation.txt --format latex --at end --yes
```

Export the active document to PDF:

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" export-pdf --path preview.pdf --yes
```

## 使用示例

读取当前选区：

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" selection --output selection.txt
```

开启修订并替换当前选区：

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" replace-selection --input revised.txt --track --yes
```

统一表格边框为浅灰网格线：

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" normalize-table-borders --table 1 --color D9DEE8 --line-width 4 --yes
```

插入类 LaTeX 公式：

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" insert-equation --input equation.txt --format latex --at end --yes
```

导出当前文档为 PDF：

```powershell
cscript //nologo "$env:USERPROFILE\.codex\skills\word-control\scripts\word_control.js" export-pdf --path preview.pdf --yes
```

## Safety Model

Mutation commands require `--yes`.

For important documents, the recommended pattern is:

1. Open the document in Word.
2. Run `status`.
3. Run `save-copy` to create a backup.
4. Apply a narrow edit.
5. Export PDF and visually check the result.

The skill is intentionally biased toward selection-based and localized edits. Whole-document rewrites are not the default.

## 安全设计

所有会修改文档的命令都要求显式传入 `--yes`。

重要文档建议流程：

1. 在 Word 中打开文档。
2. 运行 `status` 确认当前文档。
3. 运行 `save-copy` 创建备份。
4. 执行局部修改。
5. 导出 PDF 并目视检查结果。

这个 skill 默认偏向选区级和局部修改，不鼓励直接整篇重写。

## Equation Support

Word does not reliably build every raw LaTeX command through COM. This skill uses a conservative converter for common LaTeX syntax before calling Word's OMath engine.

Supported common patterns include:

```latex
x^2 + y^2 = z^2
\frac{a+b}{c}
\sqrt{a+b}
\int_0^1 x^2 dx
\sum_{i=1}^n x_i
\alpha, \beta, \mu, \Delta
\leq, \geq, \neq, \pm, \times, \cdot
```

Not yet supported:

```text
align
matrix
cases
custom macros
deeply nested complex LaTeX
```

## 公式支持

Word COM 并不能稳定直接渲染所有原始 LaTeX 命令。因此本 skill 会先把常见 LaTeX 写法保守转换为 Word 线性公式，再调用 Word 的 OMath 引擎生成内置公式对象。

目前较稳的写法包括：

```latex
x^2 + y^2 = z^2
\frac{a+b}{c}
\sqrt{a+b}
\int_0^1 x^2 dx
\sum_{i=1}^n x_i
\alpha, \beta, \mu, \Delta
\leq, \geq, \neq, \pm, \times, \cdot
```

暂不支持：

```text
align
matrix
cases
自定义宏
复杂深层嵌套 LaTeX
```

## License

No license is currently granted. The repository is public for visibility and sharing, but all rights are reserved unless a license file is added later.

## 许可证

当前未授予开源许可证。仓库是公开可见的，但除非后续添加许可证文件，否则默认保留所有权利。
