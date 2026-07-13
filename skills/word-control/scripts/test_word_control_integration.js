// Isolated command-level integration test. Run only when no user Word session exists.

var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");

function argsList() {
  var out = [];
  for (var i = 0; i < WScript.Arguments.length; i++) out.push(String(WScript.Arguments(i)));
  return out;
}

var ARGS = argsList();

function opt(name) {
  for (var i = 0; i < ARGS.length - 1; i++) {
    if (ARGS[i] === name) return ARGS[i + 1];
  }
  throw new Error("missing " + name);
}

function optionalOpt(name, fallbackValue) {
  for (var i = 0; i < ARGS.length - 1; i++) {
    if (ARGS[i] === name) return ARGS[i + 1];
  }
  return fallbackValue;
}

function quoteArg(value) {
  value = String(value);
  if (value.indexOf('"') >= 0) throw new Error("test argument contains a quote");
  return '"' + value + '"';
}

function readUtf8(path) {
  var stream = new ActiveXObject("ADODB.Stream");
  stream.Type = 2;
  stream.Charset = "utf-8";
  stream.Open();
  stream.LoadFromFile(path);
  var text = stream.ReadText();
  stream.Close();
  return text;
}

function writeUtf8(path, text) {
  var stream = new ActiveXObject("ADODB.Stream");
  stream.Type = 2;
  stream.Charset = "utf-8";
  stream.Open();
  stream.WriteText(String(text));
  stream.SaveToFile(path, 2);
  stream.Close();
}

function parseJsonFile(path) {
  return eval("(" + readUtf8(path) + ")");
}

function assertTrue(condition, message) {
  if (!condition) throw new Error(message);
}

function execute(args) {
  var parts = ["cscript", "//nologo", quoteArg(bridge)];
  for (var i = 0; i < args.length; i++) parts.push(quoteArg(args[i]));
  var process = shell.Exec(parts.join(" "));
  while (process.Status === 0) WScript.Sleep(20);
  return {
    code: Number(process.ExitCode),
    stdout: process.StdOut.ReadAll(),
    stderr: process.StdErr.ReadAll()
  };
}

function requireSuccess(args) {
  var result = execute(args);
  if (result.code !== 0) {
    throw new Error("command failed: " + args.join(" ") + " | " + result.stderr);
  }
  return result;
}

function requireFailure(args) {
  var result = execute(args);
  if (result.code === 0) throw new Error("command unexpectedly succeeded: " + args.join(" "));
  return result;
}

function runJson(args, name) {
  var output = fso.BuildPath(testDir, name);
  args.push("--output", output);
  requireSuccess(args);
  return parseJsonFile(output);
}

function runText(args, name) {
  var output = fso.BuildPath(testDir, name);
  args.push("--output", output);
  requireSuccess(args);
  return readUtf8(output);
}

function cleanWordCell(cell) {
  return String(cell.Range.Text).replace(/\x07/g, "").replace(/\r+$/g, "");
}

function exerciseAdvancedTable(path, tableIndex, fingerprint, prefix) {
  var cellInput = fso.BuildPath(testDir, prefix + "-advanced-cell.txt");
  writeUtf8(cellInput, "CENTER");

  var setCell = runJson([
    "set-cell", "--table", String(tableIndex), "--expect-table-fingerprint", String(fingerprint),
    "--row", "2", "--col", "2", "--input", cellInput, "--expect-path", path, "--yes"
  ], prefix + "-advanced-set-cell.json");
  requireFailure([
    "swap-cell-text", "--table", String(tableIndex), "--expect-table-fingerprint", String(fingerprint),
    "--from-row", "1", "--from-col", "1", "--to-row", "1", "--to-col", "3", "--expect-path", path, "--yes"
  ]);

  var swapped = runJson([
    "swap-cell-text", "--table", String(tableIndex), "--expect-table-fingerprint", String(setCell.fingerprint),
    "--from-row", "1", "--from-col", "1", "--to-row", "1", "--to-col", "3", "--expect-path", path, "--yes"
  ], prefix + "-advanced-swap.json");
  var table = doc.Tables(tableIndex);
  assertTrue(cleanWordCell(table.Cell(1, 1)) === "C1" && cleanWordCell(table.Cell(1, 3)) === "A1", "cell text swap readback failed");

  var insertedRow = runJson([
    "insert-row", "--table", String(tableIndex), "--expect-table-fingerprint", String(swapped.fingerprint),
    "--before", "2", "--expect-path", path, "--yes"
  ], prefix + "-advanced-insert-row.json");
  assertTrue(insertedRow.rows === 4 && insertedRow.cols === 3, "row insertion dimension mismatch");
  var deletedRow = runJson([
    "delete-row", "--table", String(tableIndex), "--expect-table-fingerprint", String(insertedRow.fingerprint),
    "--row", "2", "--expect-path", path, "--yes"
  ], prefix + "-advanced-delete-row.json");
  assertTrue(deletedRow.rows === 3 && deletedRow.cols === 3, "row deletion dimension mismatch");

  var insertedColumn = runJson([
    "insert-column", "--table", String(tableIndex), "--expect-table-fingerprint", String(deletedRow.fingerprint),
    "--before", "2", "--expect-path", path, "--yes"
  ], prefix + "-advanced-insert-column.json");
  assertTrue(insertedColumn.rows === 3 && insertedColumn.cols === 4, "column insertion dimension mismatch");
  var deletedColumn = runJson([
    "delete-column", "--table", String(tableIndex), "--expect-table-fingerprint", String(insertedColumn.fingerprint),
    "--col", "2", "--expect-path", path, "--yes"
  ], prefix + "-advanced-delete-column.json");
  assertTrue(deletedColumn.rows === 3 && deletedColumn.cols === 3, "column deletion dimension mismatch");

  var shading = runJson([
    "set-cell-shading", "--table", String(tableIndex), "--expect-table-fingerprint", String(deletedColumn.fingerprint),
    "--row", "2", "--col", "2", "--color", "EAF2F8", "--expect-path", path, "--yes"
  ], prefix + "-advanced-shading.json");
  assertTrue(shading.color_value === 16315114, "cell shading color value mismatch");

  var outerBorders = runJson([
    "set-table-borders", "--table", String(tableIndex), "--expect-table-fingerprint", String(shading.fingerprint),
    "--edges", "outer", "--style", "single", "--color", "4472C4", "--width", "1", "--expect-path", path, "--yes"
  ], prefix + "-advanced-outer-borders.json");
  var innerBorders = runJson([
    "set-table-borders", "--table", String(tableIndex), "--expect-table-fingerprint", String(outerBorders.fingerprint),
    "--edges", "inside-h,inside-v", "--style", "dotted", "--color", "A6A6A6", "--width", "0.5", "--expect-path", path, "--yes"
  ], prefix + "-advanced-inner-borders.json");
  table = doc.Tables(tableIndex);
  assertTrue(Number(table.Borders(-1).LineStyle) === 1, "outer table border style readback failed");
  assertTrue(Number(table.Borders(-5).LineStyle) === 2 && Number(table.Borders(-6).LineStyle) === 2, "inner table border style readback failed");
  var cellBorders = runJson([
    "set-cell-borders", "--table", String(tableIndex), "--expect-table-fingerprint", String(innerBorders.fingerprint),
    "--row", "2", "--col", "2", "--edges", "top,bottom", "--style", "double", "--color", "1F4E78", "--width", "0.75", "--expect-path", path, "--yes"
  ], prefix + "-advanced-cell-borders.json");
  var removedTop = runJson([
    "set-cell-borders", "--table", String(tableIndex), "--expect-table-fingerprint", String(cellBorders.fingerprint),
    "--row", "2", "--col", "2", "--edges", "top", "--style", "none", "--expect-path", path, "--yes"
  ], prefix + "-advanced-remove-top-border.json");

  table = doc.Tables(tableIndex);
  assertTrue(Number(table.Cell(2, 2).Shading.BackgroundPatternColor) === 16315114, "cell shading COM readback failed");
  assertTrue(Number(table.Borders(-1).LineStyle) === 1, "outer table border style readback failed");
  assertTrue(Number(table.Cell(2, 2).Borders(-1).LineStyle) === 0, "cell top border removal readback failed");
  assertTrue(Number(table.Cell(2, 2).Borders(-3).LineStyle) === 7, "cell bottom double border readback failed");
  return String(removedTop.fingerprint);
}

var bridge = fso.GetAbsolutePathName(opt("--bridge"));
var testDir = fso.GetAbsolutePathName(opt("--dir"));
var fixtureCopy = optionalOpt("--fixture-copy", "");
if (fixtureCopy) fixtureCopy = fso.GetAbsolutePathName(fixtureCopy);
if (!fso.FileExists(bridge)) throw new Error("bridge not found: " + bridge);
if (!fso.FolderExists(testDir)) throw new Error("test directory not found: " + testDir);
if (fixtureCopy && !fso.FileExists(fixtureCopy)) throw new Error("fixture copy not found: " + fixtureCopy);

var word = null;
var doc = null;
var failure = null;
var successPayload = '{"ok":true,"guarded_selection":true,"guarded_tables":true,"advanced_tables":true,"guarded_equations":true,"backup":true,"pdf":true,"close":true}';

function runFixture(path) {
  var sentinel = "WORD_CONTROL_FIXTURE_SENTINEL";
  var revisedSentinel = "WORD_CONTROL_FIXTURE_REVISED";
  var sentinelInput = fso.BuildPath(testDir, "fixture-sentinel.txt");
  var revisedInput = fso.BuildPath(testDir, "fixture-revised.txt");
  var commentInput = fso.BuildPath(testDir, "fixture-comment.txt");
  var tableInput = fso.BuildPath(testDir, "fixture-table.tsv");
  var equationInput = fso.BuildPath(testDir, "fixture-equation.txt");
  var updatedEquationInput = fso.BuildPath(testDir, "fixture-equation-updated.txt");
  writeUtf8(sentinelInput, sentinel);
  writeUtf8(revisedInput, revisedSentinel);
  writeUtf8(commentInput, "Word Control fixture comment");
  writeUtf8(tableInput, "A1\tB1\tC1\nA2\tB2\tC2\nA3\tB3\tC3");
  writeUtf8(equationInput, "x^2+y^2=z^2");
  writeUtf8(updatedEquationInput, "a^2+b^2=c^2");

  word = new ActiveXObject("Word.Application");
  word.Visible = false;
  word.DisplayAlerts = 0;
  word.ScreenUpdating = false;
  word.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
  assertTrue(Number(word.AutomationSecurity) === 3, "could not force-disable document macros");
  try { word.Options.UpdateLinksAtOpen = false; } catch (linkOptionError) {}
  try { word.Options.CheckGrammarAsYouType = false; } catch (grammarOptionError) {}
  try { word.Options.CheckSpellingAsYouType = false; } catch (spellingOptionError) {}

  doc = word.Documents.Open(path, false, false, false);
  doc.Activate();
  assertTrue(!Boolean(doc.ReadOnly), "fixture copy opened read-only");

  var paragraphCount = Number(doc.Paragraphs.Count);
  var tableCount = Number(doc.Tables.Count);
  var equationCount = Number(doc.OMaths.Count);
  var commentCount = Number(doc.Comments.Count);
  var revisionCount = Number(doc.Revisions.Count);
  var sectionCount = Number(doc.Sections.Count);
  var inlineShapeCount = Number(doc.InlineShapes.Count);
  var shapeCount = Number(doc.Shapes.Count);
  var pageCount = 0;
  try { pageCount = Number(doc.ComputeStatistics(2)); } catch (pageError) {}

  var status = runJson(["status"], "fixture-status.json");
  assertTrue(String(status.active_document.path).toLowerCase() === path.toLowerCase(), "fixture status targeted the wrong document");
  var textSample = runText(["document-text", "--max", "20000"], "fixture-document-text.txt");
  assertTrue(textSample.length > 100, "fixture document text was unexpectedly short");
  var paragraphs = runJson(["paragraphs", "--max", "40"], "fixture-paragraphs.json");
  assertTrue(paragraphs.paragraph_count === paragraphCount && paragraphs.returned > 0, "fixture paragraph inspection mismatch");
  var tables = runJson(["tables", "--max", "25"], "fixture-tables.json");
  assertTrue(tables.table_count === tableCount && tables.returned <= 25, "fixture table inspection mismatch");
  var tableReadErrors = 0;
  var tableLayoutWarnings = 0;
  var irregularTables = 0;
  for (var t = 0; t < tables.tables.length; t++) {
    tableReadErrors += tables.tables[t].read_errors.length;
    tableLayoutWarnings += tables.tables[t].layout_warnings.length;
    if (tables.tables[t].layout === "irregular") irregularTables++;
  }
  var equations = runJson(["equations"], "fixture-equations.json");
  assertTrue(equations.equation_count === equationCount, "fixture equation inspection mismatch");

  if (Boolean(doc.TrackRevisions)) {
    runJson(["disable-track-changes", "--expect-path", path, "--yes"], "fixture-track-off-before.json");
  }

  var insertionPosition = Number(doc.Content.End) - 1;
  word.Selection.SetRange(insertionPosition, insertionPosition);
  var collapsed = runJson(["selection-info"], "fixture-collapsed-selection.json");
  var collapsedGuards = [
    "--expect-start", String(collapsed.selection.start),
    "--expect-end", String(collapsed.selection.end),
    "--expect-selection-hash", String(collapsed.selection.text_hash)
  ];
  requireFailure(["replace-selection", "--input", sentinelInput, "--expect-path", path].concat(collapsedGuards, ["--yes"]));
  var inserted = runJson([
    "replace-selection", "--input", sentinelInput, "--expect-path", path
  ].concat(collapsedGuards, ["--allow-insert", "--yes"]), "fixture-insert-sentinel.json");
  assertTrue(inserted.ok, "guarded cursor insertion failed on fixture copy");

  var sentinelRange = doc.Range(insertionPosition, insertionPosition + sentinel.length);
  sentinelRange.Select();
  var selectedSentinel = runJson(["selection-info"], "fixture-sentinel-selection.json");
  var sentinelGuards = [
    "--expect-start", String(selectedSentinel.selection.start),
    "--expect-end", String(selectedSentinel.selection.end),
    "--expect-selection-hash", String(selectedSentinel.selection.text_hash)
  ];
  var comment = runJson([
    "insert-comment", "--input", commentInput, "--expect-path", path
  ].concat(sentinelGuards, ["--yes"]), "fixture-comment.json");
  assertTrue(comment.ok && Number(doc.Comments.Count) === commentCount + 1, "fixture comment insertion failed");

  sentinelRange = doc.Range(insertionPosition, insertionPosition + sentinel.length);
  sentinelRange.Select();
  selectedSentinel = runJson(["selection-info"], "fixture-sentinel-selection-track.json");
  sentinelGuards = [
    "--expect-start", String(selectedSentinel.selection.start),
    "--expect-end", String(selectedSentinel.selection.end),
    "--expect-selection-hash", String(selectedSentinel.selection.text_hash)
  ];
  runJson(["enable-track-changes", "--expect-path", path, "--yes"], "fixture-track-on.json");
  var trackedReplacement = runJson([
    "replace-selection", "--input", revisedInput, "--expect-path", path
  ].concat(sentinelGuards, ["--track", "--yes"]), "fixture-tracked-replacement.json");
  assertTrue(trackedReplacement.ok && Boolean(doc.TrackRevisions), "fixture tracked replacement failed");
  assertTrue(Number(doc.Revisions.Count) > revisionCount, "fixture tracked replacement did not create a revision");
  runJson(["disable-track-changes", "--expect-path", path, "--yes"], "fixture-track-off.json");

  var createdTable = runJson([
    "create-table", "--rows", "3", "--cols", "3", "--input", tableInput, "--at", "end", "--expect-path", path, "--yes"
  ], "fixture-create-table.json");
  var testTableIndex = tableCount + 1;
  assertTrue(createdTable.ok && createdTable.table_count === testTableIndex, "fixture table creation count mismatch");
  var advancedTableFingerprint = exerciseAdvancedTable(path, testTableIndex, createdTable.fingerprint, "fixture");
  var deletedTable = runJson([
    "delete-table", "--table", String(testTableIndex), "--expect-table-fingerprint", advancedTableFingerprint,
    "--expect-path", path, "--yes"
  ], "fixture-delete-table.json");
  assertTrue(deletedTable.ok && deletedTable.remaining_tables === tableCount, "fixture table cleanup failed");

  var insertedEquation = runJson([
    "insert-equation", "--input", equationInput, "--format", "latex", "--at", "end",
    "--expect-path", path, "--yes"
  ], "fixture-insert-equation.json");
  var testEquationIndex = equationCount + 1;
  assertTrue(insertedEquation.ok && insertedEquation.equation_count === testEquationIndex, "fixture equation insertion count mismatch");
  var setEquation = runJson([
    "set-equation", "--index", String(testEquationIndex), "--expect-equation-fingerprint", String(insertedEquation.fingerprint),
    "--input", updatedEquationInput, "--format", "latex", "--expect-path", path, "--yes"
  ], "fixture-set-equation.json");
  var deletedEquation = runJson([
    "delete-equation", "--index", String(testEquationIndex), "--expect-equation-fingerprint", String(setEquation.fingerprint),
    "--expect-path", path, "--yes"
  ], "fixture-delete-equation.json");
  assertTrue(deletedEquation.ok && deletedEquation.remaining_equations === equationCount, "fixture equation cleanup failed");

  var unsavedBackupPath = fso.BuildPath(testDir, "fixture-unsaved-copy.doc");
  var unsavedBackupResult = execute([
    "save-copy", "--path", unsavedBackupPath, "--expect-path", path, "--yes"
  ]);
  var unsavedBackupBehavior = "";
  if (unsavedBackupResult.code === 0) {
    assertTrue(fso.FileExists(unsavedBackupPath), "successful unsaved backup did not create a file");
    unsavedBackupBehavior = "live-copy-created";
  } else {
    assertTrue(unsavedBackupResult.stderr.indexOf("unsaved changes") >= 0, "unsaved backup failed without stale-copy protection");
    unsavedBackupBehavior = "stale-copy-refused";
  }

  runJson(["save-active", "--expect-path", path, "--yes"], "fixture-save.json");
  var savedBackupPath = fso.BuildPath(testDir, "fixture-saved-backup.doc");
  var savedBackup = runJson([
    "save-copy", "--path", savedBackupPath, "--expect-path", path, "--yes"
  ], "fixture-saved-backup.json");
  assertTrue(savedBackup.ok && fso.FileExists(savedBackupPath), "fixture saved backup failed");

  var pdfPath = fso.BuildPath(testDir, "fixture-preview.pdf");
  var pdf = runJson([
    "export-pdf", "--path", pdfPath, "--expect-path", path, "--yes"
  ], "fixture-pdf.json");
  assertTrue(pdf.ok && fso.FileExists(pdfPath), "fixture PDF export failed");

  runJson(["close-active", "--discard", "--expect-path", path, "--yes"], "fixture-close.json");
  doc = null;

  return "{"
    + "\"ok\":true,"
    + "\"fixture\":true,"
    + "\"paragraphs\":" + paragraphCount + ","
    + "\"tables\":" + tableCount + ","
    + "\"equations\":" + equationCount + ","
    + "\"comments\":" + commentCount + ","
    + "\"revisions\":" + revisionCount + ","
    + "\"sections\":" + sectionCount + ","
    + "\"inline_shapes\":" + inlineShapeCount + ","
    + "\"shapes\":" + shapeCount + ","
    + "\"pages\":" + pageCount + ","
    + "\"tables_returned\":" + tables.returned + ","
    + "\"irregular_tables\":" + irregularTables + ","
    + "\"table_layout_warnings\":" + tableLayoutWarnings + ","
    + "\"table_read_errors\":" + tableReadErrors + ","
    + "\"advanced_table_operations\":true,"
    + "\"unsaved_backup_behavior\":" + '"' + unsavedBackupBehavior + '"'
    + "}";
}

try {
  if (fixtureCopy) {
    successPayload = runFixture(fixtureCopy);
  } else {
  var docPath = fso.BuildPath(testDir, "command-integration.docx");
  var wrongPath = fso.BuildPath(testDir, "wrong-document.docx");
  var replacementInput = fso.BuildPath(testDir, "replacement.txt");
  var tableInput = fso.BuildPath(testDir, "table.tsv");
  var equationInput = fso.BuildPath(testDir, "simple-equation.txt");
  var updatedEquationInput = fso.BuildPath(testDir, "updated-equation.txt");

  writeUtf8(replacementInput, "edited");
  writeUtf8(tableInput, "A1\tB1\tC1\nA2\tB2\tC2\nA3\tB3\tC3");
  writeUtf8(equationInput, "x^2+y^2=z^2");
  writeUtf8(updatedEquationInput, "y^3=z^3");

  word = new ActiveXObject("Word.Application");
  word.Visible = false;
  word.DisplayAlerts = 0;
  doc = word.Documents.Add();
  doc.Content.Text = "seed text\r";
  try { doc.SaveAs2(docPath); } catch (saveError) { doc.SaveAs(docPath); }

  word.Selection.SetRange(0, 4);
  var selection = runJson(["selection-info"], "selection-info.json");
  assertTrue(String(selection.document.path).toLowerCase() === docPath.toLowerCase(), "selection inspection targeted the wrong document");
  assertTrue(String(selection.selection.text_hash).length === 8, "selection hash was not emitted");

  var selectionArgs = [
    "--expect-start", String(selection.selection.start),
    "--expect-end", String(selection.selection.end),
    "--expect-selection-hash", String(selection.selection.text_hash)
  ];
  requireFailure(["replace-selection", "--input", replacementInput, "--expect-path", wrongPath].concat(selectionArgs, ["--yes"]));

  var replacement = runJson([
    "replace-selection", "--input", replacementInput, "--expect-path", docPath
  ].concat(selectionArgs, ["--yes"]), "replace.json");
  assertTrue(replacement.ok && String(doc.Content.Text).indexOf("edited") === 0, "guarded selection replacement failed");
  requireFailure(["insert-comment", "--input", replacementInput, "--expect-path", docPath].concat(selectionArgs, ["--yes"]));

  var createdTable = runJson([
    "create-table", "--rows", "3", "--cols", "3", "--input", tableInput, "--at", "end", "--expect-path", docPath, "--yes"
  ], "create-table.json");
  assertTrue(createdTable.ok && String(createdTable.fingerprint).length === 8, "guarded table creation failed");

  var tables = runJson(["tables"], "tables.json");
  assertTrue(tables.table_count === 1, "table inspection count mismatch");
  assertTrue(tables.tables[0].read_errors.length === 0, "table inspection reported read errors");
  var oldTableFingerprint = String(tables.tables[0].fingerprint);

  var advancedTableFingerprint = exerciseAdvancedTable(docPath, 1, oldTableFingerprint, "synthetic");
  requireFailure([
    "delete-table", "--table", "1", "--expect-table-fingerprint", oldTableFingerprint,
    "--expect-path", docPath, "--yes"
  ]);

  var deletedTable = runJson([
    "delete-table", "--table", "1", "--expect-table-fingerprint", advancedTableFingerprint,
    "--expect-path", docPath, "--yes"
  ], "delete-table.json");
  assertTrue(deletedTable.ok && deletedTable.remaining_tables === 0, "guarded table deletion failed");

  var insertedEquation = runJson([
    "insert-equation", "--input", equationInput, "--format", "latex", "--at", "end",
    "--expect-path", docPath, "--yes"
  ], "insert-equation.json");
  assertTrue(insertedEquation.ok && insertedEquation.equation_count === 1, "guarded equation insertion failed");

  var equations = runJson(["equations"], "equations.json");
  assertTrue(equations.equation_count === 1, "equation inspection count mismatch");
  var oldEquationFingerprint = String(equations.equations[0].fingerprint);

  var setEquation = runJson([
    "set-equation", "--index", "1", "--expect-equation-fingerprint", oldEquationFingerprint,
    "--input", updatedEquationInput, "--format", "latex", "--expect-path", docPath, "--yes"
  ], "set-equation.json");
  assertTrue(setEquation.ok && setEquation.equation_count === 1, "guarded equation replacement changed the equation count");
  assertTrue(String(setEquation.fingerprint) !== oldEquationFingerprint, "equation fingerprint did not change");
  requireFailure([
    "delete-equation", "--index", "1", "--expect-equation-fingerprint", oldEquationFingerprint,
    "--expect-path", docPath, "--yes"
  ]);

  var deletedEquation = runJson([
    "delete-equation", "--index", "1", "--expect-equation-fingerprint", String(setEquation.fingerprint),
    "--expect-path", docPath, "--yes"
  ], "delete-equation.json");
  assertTrue(deletedEquation.ok && deletedEquation.remaining_equations === 0, "guarded equation deletion failed");

  runJson(["save-active", "--expect-path", docPath, "--yes"], "save.json");
  var backupPath = fso.BuildPath(testDir, "command-integration.backup.docx");
  var backup = runJson([
    "save-copy", "--path", backupPath, "--expect-path", docPath, "--yes"
  ], "backup.json");
  assertTrue(backup.ok && fso.FileExists(backupPath), "guarded backup creation failed");

  var pdfPath = fso.BuildPath(testDir, "command-integration.pdf");
  var pdf = runJson([
    "export-pdf", "--path", pdfPath, "--expect-path", docPath, "--yes"
  ], "pdf.json");
  assertTrue(pdf.ok && fso.FileExists(pdfPath), "guarded PDF export failed");

  runJson(["close-active", "--discard", "--expect-path", docPath, "--yes"], "close.json");
  doc = null;
  }
} catch (e) {
  failure = e;
} finally {
  if (doc !== null) {
    try { doc.Close(false); } catch (closeError) {}
    doc = null;
  }
  if (word !== null) {
    try { word.Quit(0); } catch (quitError) {}
    word = null;
  }
  try { CollectGarbage(); } catch (gcError) {}
}

if (failure !== null) {
  WScript.StdErr.WriteLine("word-control integration error: " + (failure.message || String(failure)));
  WScript.Quit(1);
}

WScript.Echo(successPayload);
