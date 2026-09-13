# Word Control Skill

Word Control is a Codex skill for narrow, in-place editing of the active Microsoft Word desktop document on Windows, using local Word COM automation.

它直接检查和修改 Word 中已经打开的文档，支持文字、批注、修订、表格、公式、备份和 PDF 导出。首次修改用户文档前，AI 应按操作手册建立可核对的备份；每次写入仍需核对文档身份和最新的选区或对象指纹，并检查修改影响范围内的前后差异。

## 运行条件

- Windows、桌面版 Microsoft Word、Windows Script Host（`cscript.exe`）。
- 以下安装命令使用 PowerShell；Git 用于克隆仓库，也可以下载 ZIP。
- 只有运行完整测试套件才需要 Node.js；日常 Word 操作不依赖 Node.js。

这是本地 Word COM 技能，不支持 Word Online。

## 安装 Skill

本项目使用以下位置作为唯一的用户全局安装：

```text
%USERPROFILE%\.agents\skills\word-control\SKILL.md
```

只安装仓库中的 **`skills/word-control` 整个目录**，其中包含入口、手册、脚本和界面提示。仓库本身可以放在任意源码目录。不要额外在 `.codex/skills`、`.grok/skills` 或项目的 `.agents/skills` 中安装同名副本或链接；旧版备份也应放在技能发现目录之外。Codex 会读取用户目录下的 `.agents/skills`，同名技能可能同时出现在选择器中。[官方技能文档](https://learn.chatgpt.com/docs/build-skills#where-codex-loads-local-skills)

### 1. 获取仓库

在准备存放源码的目录打开 PowerShell，执行：

```powershell
git clone https://github.com/linnnn89/word-control-skill.git
if ($LASTEXITCODE -ne 0) { throw 'Clone failed; check the destination and Git output.' }
Set-Location -LiteralPath '.\word-control-skill'
```

也可以在 [GitHub 仓库](https://github.com/linnnn89/word-control-skill) 点击 **Code → Download ZIP**，解压后在含有 `README.md` 和 `skills` 的仓库根目录打开 PowerShell。已有源码时直接使用该目录；升级前先核对本地修改和准备安装的版本。

### 2. 首次安装

在仓库根目录执行。下面的命令遇到已有安装会停止，请改按后面的升级步骤处理。

```powershell
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath '.\skills\word-control').Path
$skillsRoot = Join-Path $env:USERPROFILE '.agents\skills'
$destination = Join-Path $skillsRoot 'word-control'
if (-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md') -PathType Leaf)) {
    throw 'The source must be the complete skills/word-control directory.'
}
if (Test-Path -LiteralPath $destination) {
    throw 'word-control is already installed. Follow the upgrade steps first.'
}
New-Item -ItemType Directory -Path $skillsRoot -Force | Out-Null
Copy-Item -LiteralPath $source -Destination $destination -Recurse
```

### 3. 校验并使用

继续在同一个 PowerShell 窗口执行，逐文件核对 SHA-256，再运行不连接 Word 的帮助命令：

```powershell
$sourceFiles = @(Get-ChildItem -LiteralPath $source -File -Recurse)
$installedFiles = @(Get-ChildItem -LiteralPath $destination -File -Recurse)
if ($sourceFiles.Count -ne $installedFiles.Count) { throw 'Installed file count differs.' }
foreach ($file in $sourceFiles) {
    $relative = $file.FullName.Substring($source.Length + 1)
    $installedFile = Join-Path $destination $relative
    if ((Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash -ne
        (Get-FileHash -LiteralPath $installedFile -Algorithm SHA256).Hash) {
        throw "Installed file differs: $relative"
    }
}
cscript //nologo (Join-Path $destination 'scripts\word_control.js') help
if ($LASTEXITCODE -ne 0) { throw 'Skill help failed.' }
Write-Output "Installed and verified: $destination"
```

确认 `SKILL.md` 直接位于 `word-control` 下，没有多嵌套一层同名文件夹。以上检查验证安装文件和命令入口，真实 Word 编辑的验收见[维护手册](skills/word-control/references/maintenance.md)。

Codex 会自动发现新安装或更新的技能；在后续消息中输入 `$word-control` 使用。如果仍未出现，重启 Codex 后再试。[官方安装说明](https://learn.chatgpt.com/docs/build-skills#install-curated-skills-for-local-use)

例如，在桌面 Word 中打开需要处理的文档，然后对 Codex 说：

> 使用 $word-control，先确认当前文档并读取操作手册，按手册备份后，把我选中的文字修改为……，核对受影响范围内的前后差异。

### 已有安装：升级和恢复

先比较现有安装和准备安装的版本。有本地定制时，先保留并合并这些修改；维护本项目的规范副本时，遵循[维护手册](skills/word-control/references/maintenance.md)，不要用较旧的下载版本覆盖本地开发成果。

确定要替换后，在没有任务正在使用该技能时，把旧安装整体移到技能发现目录之外。以下命令保留旧文件，并拒绝复用已有备份路径：

```powershell
$ErrorActionPreference = 'Stop'
$installedSkill = Join-Path $env:USERPROFILE '.agents\skills\word-control'
$backupRoot = Join-Path $env:USERPROFILE '.agents\skill-backups'
$backup = Join-Path $backupRoot ('word-control-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
if (-not (Test-Path -LiteralPath $installedSkill -PathType Container)) { throw 'No installed skill to back up.' }
if (Test-Path -LiteralPath $backup) { throw 'Backup path already exists.' }
New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
Move-Item -LiteralPath $installedSkill -Destination $backup
Write-Output "Previous installation preserved at: $backup"
```

然后在准备安装的仓库根目录重复“首次安装”和“校验并使用”。若新安装失败，保留备份，将失败的新目录移出技能发现目录，再把旧目录移回原安装路径。校验通过前不要删除旧版，也不要把备份命名为另一个位于 `skills` 下的目录。

### 给 AI 的安装指令

可以直接把下面这段话连同本 README 交给具备本地文件和命令权限的 AI：

> 请从 https://github.com/linnnn89/word-control-skill 获取源码，按 README 将 skills/word-control 完整安装到当前 Windows 用户的 %USERPROFILE%\.agents\skills\word-control，仅保留这一处全局安装。先检查已有安装和本地修改；升级时把旧目录保存在技能发现目录之外，再安装并逐文件校验 SHA-256。不要覆盖未处理的本地定制或自动删除其他已有副本。运行 cscript 的 help 验证入口，报告实际安装路径、校验结果及尚存的重复安装；安装后读取 SKILL.md，并按任务只读相关操作手册。

如果 AI 使用已有的 `skill-installer`，仓库参数为 `linnnn89/word-control-skill`，子目录参数为 `skills/word-control`，并须显式把 `--dest` 指向 `%USERPROFILE%\.agents\skills` 这个**父目录**。在 PowerShell 中用 `"$env:USERPROFILE\.agents\skills"` 展开路径；不要沿用安装器的其他默认目录。已有安装仍按上述升级流程处理。

## 入口与操作手册

安装后从简短的 [SKILL.md](skills/word-control/SKILL.md) 进入，按任务加载所需手册。README 负责安装和上手，具体命令与编辑规则以手册为准。

| 任务 | 读取位置 |
| --- | --- |
| 修改前备份、编辑顺序、验收与医学稿件格式边界 | [操作流程](skills/word-control/references/workflow.md) |
| 查询范围、命令参数、表格/公式操作、保存与导出 | [命令手册](skills/word-control/references/usage.md) |
| 目标已变化、备份或保存失败、部分结果与清理 | [故障恢复](skills/word-control/references/recovery.md) |
| 修改技能、测试、同步与性能实验 | [维护手册](skills/word-control/references/maintenance.md) |

备份要求由 AI 按工作流执行，命令本身没有自动备份门禁。连续修改同一文档可复用本次任务的有效基线备份；失败处理、明确豁免和核对要求见[备份规则](skills/word-control/references/workflow.md#backup-before-editing)。

维护测试可使用 `/测试用WORD/` 中的本地论文；该目录及内容已加入 Git ignore。通过 `-FixturePath` 指定后，测试会使用副本并核对源文件哈希，详见维护手册。

## 能力边界

- 适合小范围编辑；新建文档或整体重建不在本技能范围内。
- 复杂 LaTeX 环境和自定义宏需要人工转换及视觉核对。
- 合并或不规则表格的结构编辑受限。
- Word 对话框可能阻塞自动化，需要用户处理。
- 保存和导出成功仍需按影响范围核对内容；文件恢复机制不保证断电或进程终止时的原子提交。

## License

No open-source license is currently granted. All rights reserved by the repository owner.
