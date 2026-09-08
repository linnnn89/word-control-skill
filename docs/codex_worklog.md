# Work log

## 2026-09-08 (Beijing time)

- Goal: publish the validated canonical Codex skill and merge the reliability fixes into main.
- Changes: escape all JSON control characters; validate original LaTeX tokens and balanced braces; require selection story-type guards and saved-document paths; reject incomplete fingerprints; preflight scratch outputs; restore tracking state after a failed edit. Equation edits retain the original range story. Include the canonical medical-manuscript formatting gate.
- Tests: seven pure regression groups and 27 strictly parsed JSON artifacts passed in the canonical installation. Full isolated Word command integration and advanced table operations passed, including backup, PDF export, and close checks. An earlier attempt skipped integration because Word was running; the successful run followed after that session exited.
- Test setup: the Node helper uses `.cjs` to avoid inheriting host ES-module settings. No dependency installation or global configuration change was needed.
- Publication: runtime and test scripts are copied byte-for-byte from the validated canonical installation. Public usage documentation omits machine-local maintenance paths. README documents the new guard requirements and test prerequisites. Pure regressions and skill validation also passed in the repository checkout.
- Preparation correction: Python initially inherited the Windows default encoding while reading documentation; explicit UTF-8 resolved the read error.
- Scope: no batch-edit/search API, real-document fixture test, visual layout acceptance, or performance benchmark. Existing local backup handling is unchanged.
