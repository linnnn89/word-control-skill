# Maintenance and validation

Use `.agents/skills/word-control` as the sole global installation and maintenance source for this setup. Keep backups outside skill discovery directories. Edit and validate this installation, then synchronize the repository copy and compare every file by SHA-256. Do not recreate additional global copies or discovery links under other skill directories. Keep machine-local maintenance history in the backup, outside the distributed manuals.

Ordinary commands use Windows Script Host (`cscript`), not Node.js. Tests require Windows desktop Word, PowerShell and Node.js.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\.agents\skills\word-control\scripts\test_word_control.ps1"
```

The test creates a unique temporary DOCX in a hidden Word instance, runs pure regressions and command integration, and strictly parses every generated JSON output. It checks paragraph/cell pagination, literal search in main text and notes with Find-setting restoration, targeted equations, summary/text modes without mutation guards, partial-read errors, verified cell/table/border results, failure reporting after partial writes or rollback, guarded text/table/equation edits, backups, PDF export and cleanup. XML diagnostics are tested with real MSXML normalization and Word text, shading, border, row, column, merge and tracked-revision changes; experimental hashes are rejected for writes. It removes its own temporary artifacts.

An existing Word session causes command integration to be skipped. `skipped-existing-word-session` is not complete validation and does not authorize synchronization. Never attach tests to or close the user's Word session. Run real command integration when Word is not already running, and confirm task-owned Word processes exit.

`-FixturePath` optionally exercises a copy of a specifically chosen document and checks that the original file hash stays unchanged. Synthetic integration alone does not establish complex-document layout fidelity; inspect a relevant rendered output when a change affects layout.

Validate the Skill frontmatter and its relative links after changing the manuals. Keep the editing workflow, command help and examples consistent with actual runtime behavior. Preserve the medical-manuscript formatting gate and fresh-target guards when reducing instructions.

For performance comparisons, measure complete equivalent workflows, including process startup and required verification, on disposable copies. Verified border results already provide the next expected fingerprint; measure removal of an unnecessary follow-up query separately from changes to the write implementation.

`tables --table N --compare-xml` is an opt-in experiment, not a faster mutation guard. Its `elapsed_ms` covers two XML captures, parsing, normalization, comparison and hashing; the command still performs the ordinary full inspection first. Do not compare that diagnostic substep against an entire write workflow or claim end-to-end write acceleration. Editing-session IDs are excluded because repeated Word exports can vary them without a content edit. All other package content remains, which may still yield unrelated changes or miss differences in effective COM properties. Passing the current fixtures does not establish general equivalence. Keep XML candidates out of write authorization until a separately reviewed change establishes suitable coverage and fallback behavior.

The parser explicitly preserves whitespace, including whitespace-only text. MSXML's default can strip those nodes; see Microsoft's [preserveWhiteSpace documentation](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms761353(v=vs.85)).

Before accepting a runtime change, define its affected scope and allowed differences, then compare before/after state independently of the production fingerprint. Read-only XML checks compare all story text, object counts, bookmark ranges, target-cell text/font/shading/borders, selection and save state. Cell-border checks preserve non-border state and all boundaries outside the target and its shared neighbors. Document tests also retain original-file hashes and save/reopen checks. These checks establish the tested scope, not arbitrary document fidelity; add relevant visual comparison when layout behavior changes.

Known upstream risk signals include Office.js issues [#5014](https://github.com/OfficeDev/office-js/issues/5014) (cross-paragraph bookmark omission), [#6059](https://github.com/OfficeDev/office-js/issues/6059) (image-related OOXML retrieval failure) and [#6803](https://github.com/OfficeDev/office-js/issues/6803) (shape-body reads changing formatting on Mac). They are not confirmed failures of this Windows COM path. Microsoft's [Range.WordOpenXML contract](https://learn.microsoft.com/en-us/office/vba/api/word.range.wordopenxml) only covers markup needed for the requested range. Do not infer complete bookmark, document or effective-style coverage from a stable range snapshot, use it for XML writeback, or treat the 32-bit diagnostic hash as collision-proof.
