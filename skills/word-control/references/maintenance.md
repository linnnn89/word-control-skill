# Maintenance and validation

Use `.agents/skills/word-control` as the sole global installation and maintenance source for this setup. Keep backups outside skill discovery directories. Edit and validate this installation, then synchronize the repository copy and compare every file by SHA-256. Do not recreate additional global copies or discovery links under other skill directories. Keep machine-local maintenance history in the backup, outside the distributed manuals.

Ordinary commands use Windows Script Host (`cscript`), not Node.js. Tests require Windows desktop Word, PowerShell and Node.js.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\.agents\skills\word-control\scripts\test_word_control.ps1"
```

The test creates a unique temporary DOCX in a hidden Word instance, runs pure regressions and command integration, and strictly parses every generated JSON output. It checks paragraph/cell pagination, literal search in main text and notes with Find-setting restoration, targeted equations, summary/text modes without mutation guards, partial-read errors, verified cell/table/border results, failure reporting after partial writes or rollback, guarded text/table/equation edits, backups, PDF export and cleanup. It removes its own temporary artifacts.

An existing Word session causes command integration to be skipped. `skipped-existing-word-session` is not complete validation and does not authorize synchronization. Never attach tests to or close the user's Word session. Run real command integration when Word is not already running, and confirm task-owned Word processes exit.

`-FixturePath` optionally exercises a copy of a specifically chosen document and checks that the original file hash stays unchanged. Synthetic integration alone does not establish complex-document layout fidelity; inspect a relevant rendered output when a change affects layout.

Validate the Skill frontmatter and its relative links after changing the manuals. Keep the editing workflow, command help and examples consistent with actual runtime behavior. Preserve the medical-manuscript formatting gate and fresh-target guards when reducing instructions.
