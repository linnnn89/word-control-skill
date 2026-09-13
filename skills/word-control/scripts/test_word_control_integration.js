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
  var deadline = new Date().getTime() + 120000;
  while (process.Status === 0) {
    if (new Date().getTime() > deadline) {
      process.Terminate();
      throw new Error("command timed out; bridge process terminated (Word ownership remains with test)");
    }
    WScript.Sleep(20);
  }
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

function captureCellBorders(table) {
  var result = [], cells = table.Range.Cells;
  for (var i = 1; i <= cells.Count; i++) {
    var borders = cells(i).Borders;
    for (var edge = 1; edge <= 4; edge++) {
      var border = borders(-edge);
      result.push(String(border.LineStyle) + ":" + border.Color + ":" + border.LineWidth);
    }
  }
  return result;
}

function captureInspectionFootprint(document, table, skipBorders, skipShadingCell) {
  // Independent COM readback: no production fingerprint, XML normalizer or short hash.
  var values = [document.FullName, document.TrackRevisions], collections = ["Tables", "OMaths", "Comments", "Revisions",
    "Sections", "InlineShapes", "Shapes", "Fields", "Bookmarks", "Footnotes", "Endnotes"];
  for (var i = 0; i < collections.length; i++) values.push(document[collections[i]].Count);
  for (var stories = new Enumerator(document.StoryRanges); !stories.atEnd(); stories.moveNext()) {
    var story = stories.item();
    while (story) {
      values.push(story.StoryType, story.Start, story.End, String(story.Text));
      story = story.NextStoryRange;
    }
  }
  for (var b = 1; b <= document.Bookmarks.Count; b++) {
    var bookmark = document.Bookmarks(b), bookmarked = bookmark.Range;
    values.push(bookmark.Name, bookmarked.StoryType, bookmarked.Start, bookmarked.End);
  }
  var cells = table.Range.Cells;
  values.push(cells.Count);
  for (var c = 1; c <= cells.Count; c++) {
    var cell = cells(c), range = cell.Range, font = range.Font;
    values.push(range.Start, range.End, String(range.Text), font.Name, font.Size, font.Bold, font.Italic, font.Color,
      range.ParagraphFormat.Alignment, cell.Shading.ForegroundPatternColor);
    if (c !== skipShadingCell) values.push(cell.Shading.BackgroundPatternColor, cell.Shading.Texture);
  }
  if (!skipBorders) values = values.concat(captureCellBorders(table));
  for (var v = 0; v < values.length; v++) { var text = String(values[v]); values[v] = text.length + ":" + text; }
  return values.join("|");
}

function exerciseCellTextFidelity() {
  var sample = word.Documents.Add(), table = null;
  try {
    sample.Content.Text = "before\rafter\r";
    sample.Sections(1).Headers(1).Range.Text = "untouched header";
    sample.Bookmarks.Add("outside_swap", sample.Range(0, 6));
    table = sample.Tables.Add(sample.Range(7, 7), 1, 3);
    var originals = ["A\r\r", "B\x0bC\t\r", "untouched"];
    for (var i = 1; i <= 3; i++) {
      var range = table.Cell(1, i).Range; range.End--; range.Text = originals[i - 1];
    }
    var samplePath = fso.BuildPath(testDir, "cell-text-fidelity.docx");
    sample.SaveAs2(samplePath);
    // Fresh Word documents expose empty header/footer stories that can disappear on first reopen.
    // Establish a persisted baseline so the later comparison isolates the cell edits.
    sample.Close(false); sample = word.Documents.Open(samplePath); sample.Activate(); table = sample.Tables(1);
    var first = runJson(["tables", "--table", "1"], "fidelity-initial.json").tables[0];
    var last = table.Cell(1, 3).Range; last.End--; last.Text = originals[2] + "\r";
    var changed = runJson(["tables", "--table", "1"], "fidelity-paragraph-change.json").tables[0];
    assertTrue(first.fingerprint !== changed.fingerprint, "Trailing cell paragraph did not change the guard");
    var changedFootprint = captureInspectionFootprint(sample, table, false), savedBeforeRejection = Boolean(sample.Saved);
    requireFailure(["swap-cell-text", "--table", "1", "--from-cell", "1", "--to-cell", "2", "--expect-path", samplePath,
      "--expect-table-fingerprint", String(first.fingerprint), "--yes"]);
    assertTrue(captureInspectionFootprint(sample, table, false) === changedFootprint && Boolean(sample.Saved) === savedBeforeRejection,
      "Rejected stale guard changed the document");
    last = table.Cell(1, 3).Range; last.End--; last.Text = originals[2];
    var fresh = runJson(["tables", "--table", "1"], "fidelity-restored-baseline.json").tables[0];
    var a = String(table.Cell(1, 1).Range.Text), b = String(table.Cell(1, 2).Range.Text), c = String(table.Cell(1, 3).Range.Text);
    assertTrue(a === originals[0] + "\r\x07" && b === originals[1] + "\r\x07", "Native fixture did not retain exact break characters");
    var tableText = String(table.Range.Text), body = String(sample.Content.Text);
    assertTrue(tableText.indexOf(a + b + c) === 0 && body.indexOf(tableText) === body.lastIndexOf(tableText), "Unexpected native table serialization");
    var expectedTable = b + a + c + tableText.substring(a.length + b.length + c.length);
    var expectedBody = body.replace(tableText, expectedTable);
    var before = captureInspectionFootprint(sample, table, false), paragraphs = Number(sample.Paragraphs.Count);
    var swapped = runJson(["swap-cell-text", "--table", "1", "--from-cell", "1", "--to-cell", "2", "--expect-path", samplePath,
      "--expect-table-fingerprint", String(fresh.fingerprint), "--yes"], "fidelity-swapped.json");
    assertTrue(String(table.Cell(1, 1).Range.Text) === b && String(table.Cell(1, 2).Range.Text) === a, "Swap lost cell text or break characters");
    assertTrue(String(sample.Content.Text) === expectedBody && Number(sample.Paragraphs.Count) === paragraphs, "Swap changed text outside its two target cells or paragraph count");
    assertTrue(String(table.Cell(1, 3).Range.Text) === c && String(sample.Sections(1).Headers(1).Range.Text) === "untouched header\r",
      "Swap changed the neighboring cell or header");
    runJson(["save-active", "--expect-path", samplePath, "--yes"], "fidelity-saved.json");
    sample.Close(false); sample = word.Documents.Open(samplePath); sample.Activate(); table = sample.Tables(1);
    assertTrue(String(table.Cell(1, 1).Range.Text) === b && String(table.Cell(1, 2).Range.Text) === a && String(sample.Content.Text) === expectedBody,
      "Saved/reopened swap lost exact cell text");
    var reopened = runJson(["tables", "--table", "1"], "fidelity-reopened.json").tables[0];
    runJson(["swap-cell-text", "--table", "1", "--from-cell", "1", "--to-cell", "2", "--expect-path", samplePath,
      "--expect-table-fingerprint", String(reopened.fingerprint), "--yes"], "fidelity-swapped-back.json");
    assertTrue(captureInspectionFootprint(sample, table, false) === before, "Swap round trip changed story text, objects, bookmarks, cell formatting or borders");
    var nestedRange = table.Cell(1, 1).Range; nestedRange.Collapse(1);
    sample.Tables.Add(nestedRange, 1, 1);
    var nested = runJson(["tables", "--table", "1"], "fidelity-nested.json").tables[0];
    var nestedBefore = captureInspectionFootprint(sample, table, false);
    requireFailure(["swap-cell-text", "--table", "1", "--from-cell", "1", "--to-cell", "3", "--expect-path", samplePath,
      "--expect-table-fingerprint", String(nested.fingerprint), "--yes"]);
    assertTrue(captureInspectionFootprint(sample, table, false) === nestedBefore, "Nested-cell refusal changed the document");
  } finally { sample.Close(false); doc.Activate(); }
}

function exerciseVerifiedCellResults() {
  var sample = word.Documents.Add(), originalBridge = bridge;
  try {
    sample.Content.Text = "before\rafter\r";
    sample.Sections(1).Headers(1).Range.Text = "untouched header";
    sample.Bookmarks.Add("outside_cells", sample.Range(0, 6));
    var table = sample.Tables.Add(sample.Range(7, 7), 2, 3);
    for (var c = 1; c <= 6; c++) { var range = table.Range.Cells(c).Range; range.End--; range.Text = "cell" + c; }
    table.Range.Cells(5).Shading.BackgroundPatternColor = 16777215;
    table.Range.Cells(5).Shading.Texture = 0;
    var samplePath = fso.BuildPath(testDir, "verified-cell-results.docx"); sample.SaveAs2(samplePath);
    sample.Close(false); sample = word.Documents.Open(samplePath); sample.Activate(); table = sample.Tables(1);
    var before = captureInspectionFootprint(sample, table, false);
    var shadeBefore = captureInspectionFootprint(sample, table, false, 5);
    var source = readUtf8(bridge), needle = "function tableFingerprint(table, readErrors) {";
    assertTrue(source.indexOf(needle) >= 0, "Final-inspection injection point missing");
    var faultBridge = fso.BuildPath(testDir, "cell-final-inspection-failure.js");
    writeUtf8(faultBridge, source.replace(needle, "var injectedPostChecks = 0;\n" + needle
      + "\n  if (++injectedPostChecks === 2) { if (readErrors) readErrors.push('injected incomplete final inspection'); return null; }"));
    var initial = runJson(["tables", "--table", "1"], "cell-result-initial.json").tables[0];
    var fingerprint = initial.fingerprint;
    for (var phase = 0; phase < 2; phase++) {
      var command = phase === 0 ? "set-cell-shading" : "swap-cell-text";
      var target = phase === 0 ? ["--cell", "5", "--color", "FF0000"] : ["--from-cell", "1", "--to-cell", "2"];
      var args = [command, "--table", "1", "--expect-path", samplePath, "--expect-table-fingerprint", fingerprint, "--yes"];
      var output = fso.BuildPath(testDir, command + "-failed-postcheck.json");
      var outcome;
      bridge = faultBridge;
      try { outcome = requireFailure(args.concat(target, ["--output", output])); } finally { bridge = originalBridge; }
      var failed = parseJsonFile(output);
      assertTrue(outcome.code === 3 && failed.ok === false && failed.applied === true && failed.verified === false
        && failed.inspection_complete === false && failed.fingerprint === null && failed.readback.matches_requested === true,
        "An incomplete final inspection lost its applied state or authorized further writes");
      assertTrue(failed.failure_count === 0 && failed.rolled_back === false && failed.errors.join(" ").indexOf("injected") >= 0,
        "A final-inspection failure was confused with a write failure or rollback");
      assertTrue(String(failed.document.path).toLowerCase() === samplePath.toLowerCase(), "Failure result lost document identity");
      if (phase === 0) {
        assertTrue(Number(table.Range.Cells(5).Shading.BackgroundPatternColor) === 255 && Number(table.Range.Cells(5).Shading.Texture) === 0,
          "Shading failure receipt did not describe the native state");
        assertTrue(captureInspectionFootprint(sample, table, false, 5) === shadeBefore, "Shading affected text, non-target formatting or document structure");
      } else {
        assertTrue(cleanWordCell(table.Range.Cells(1)) === "cell2" && cleanWordCell(table.Range.Cells(2)) === "cell1",
          "Swap failure receipt did not describe the native state");
      }
      var fresh = runJson(["tables", "--table", "1"], command + "-recover-inspection.json").tables[0];
      var restored = runJson([command, "--table", "1", "--expect-path", samplePath, "--expect-table-fingerprint", fresh.fingerprint, "--yes"]
        .concat(phase === 0 ? ["--cell", "5", "--color", "FFFFFF"] : target), command + "-restored.json");
      assertTrue(restored.ok && restored.applied && restored.verified && restored.inspection_complete && restored.readback.matches_requested
        && restored.failure_count === 0 && restored.errors.length === 0 && restored.fingerprint, "Successful cell result was not reusable");
      assertTrue(captureInspectionFootprint(sample, table, false) === before, "Cell recovery changed the independently captured scope");
      fingerprint = restored.fingerprint;
    }
    sample.Save(); sample.Close(false); sample = word.Documents.Open(samplePath); sample.Activate();
    assertTrue(captureInspectionFootprint(sample, sample.Tables(1), false) === before, "Cell-result scenario changed after save and reopen");
  } finally { bridge = originalBridge; sample.Close(false); doc.Activate(); }
}

function exerciseXmlNormalization() {
  // Exercise the shipped normalizer in real MSXML, without substituting a different XML parser.
  var source = readUtf8(bridge);
  var start = source.indexOf("function parseSnapshotXml(");
  var end = source.indexOf("function equationFingerprint(", start);
  assertTrue(start >= 0 && end > start, "XML helper definitions missing");
  eval(source.substring(start, end));
  var xml = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    + '<pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body><w:tbl><w:tr w:rsidTr="00000001"><w:tc><w:tcPr><w:shd w:fill="FFFFFF"/></w:tcPr><w:p w:rsidR="00000001"><w:r><w:t xml:space="preserve"> A </w:t></w:r><w:ins w:id="7"><w:r><w:t>B</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl></w:body></w:document></pkg:xmlData></pkg:part>'
    + '<pkg:part pkg:name="/word/settings.xml"><pkg:xmlData><w:settings><w:rsids><w:rsid w:val="00000001"/></w:rsids></w:settings></pkg:xmlData></pkg:part>'
    + '<pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles><w:style w:styleId="S"><w:rsid w:val="00000002"/><w:tblPr><w:tblBorders><w:top w:val="single" w:color="000000"/></w:tblBorders></w:tblPr></w:style></w:styles></pkg:xmlData></pkg:part>'
    + '<pkg:part pkg:name="/word/theme/theme1.xml"><pkg:xmlData><a:theme><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="336699"/></a:accent1></a:clrScheme></a:themeElements></a:theme></pkg:xmlData></pkg:part></pkg:package>';
  var normalized = normalizeTableXml(xml);
  assertTrue(normalized === normalizeTableXml(xml.replace(/00000001/g, "00000009")), "editing-session metadata changed normalization");
  assertTrue(normalizeTableXml(xml.replace(">B<", "> <")) !== normalizeTableXml(xml.replace(">B<", ">  <")), "normalizer lost whitespace-only text");
  var changes = [["> A <", "> A  <"], ['w:fill="FFFFFF"', 'w:fill="FF0000"'], ['w:color="000000"', 'w:color="112233"'],
    ['val="336699"', 'val="669933"'], ['w:id="7"', 'w:id="8"'], ['w:val="00000002"', 'w:val="00000003"']];
  for (var i = 0; i < changes.length; i++) {
    assertTrue(normalized !== normalizeTableXml(xml.replace(changes[i][0], changes[i][1])), "normalizer lost content, formatting, revision or retained metadata");
  }
  var rejected = false;
  try { normalizeTableXml('<!DOCTYPE pkg:package [<!ENTITY forbidden "EXPANDED">]>' + xml); }
  catch (e) { rejected = true; }
  assertTrue(rejected, "XML normalization accepted a DTD");
  rejected = false;
  try { normalizeTableXml('<not-a-table/>'); } catch (e2) { rejected = true; }
  assertTrue(rejected, "XML normalization accepted a snapshot without a table");
  rejected = false;
  try { parseSnapshotXml(new Array(8 * 1024 * 1024 + 2).join("x")); }
  catch (e3) { rejected = String(e3.message).indexOf("snapshot limit") >= 0; }
  assertTrue(rejected, "XML normalization did not enforce the snapshot limit");
}

function exerciseXmlMutationDetection() {
  var previous = doc, sample = null;
  try {
    sample = word.Documents.Add();
    var path = fso.BuildPath(testDir, "xml-comparison.docx");
    sample.Content.Text = "OUTSIDE BEFORE\rOUTSIDE AFTER\r";
    var position = String("OUTSIDE BEFORE\r").length;
    var table = sample.Tables.Add(sample.Range(position, position), 2, 2);
    var range = table.Cell(1, 1).Range; range.End--; range.Text = "A";
    sample.Bookmarks.Add("cross_table_boundary", sample.Range(table.Range.Start, sample.Content.End - 1));
    sample.Footnotes.Add(sample.Range(2, 2)).Range.Text = "NOTE OUTSIDE TABLE";
    sample.Comments.Add(sample.Range(0, 1), "COMMENT OUTSIDE TABLE");
    sample.SaveAs2(path, 16);
    var lastCandidate = null;
    function inspect(name) {
      var selectionStart = Number(word.Selection.Start), selectionEnd = Number(word.Selection.End), saved = Boolean(sample.Saved);
      var footprint = captureInspectionFootprint(sample, table, false);
      var result = runJson(["tables", "--table", "1", "--compare-xml", "--expect-path", path], "xml-" + name + ".json");
      var comparison = result.tables[0].xml_comparison;
      assertTrue(comparison.diagnostic_only === true && comparison.stable_repeat === true && comparison.errors.length === 0, "XML comparison did not complete: " + name);
      assertTrue(/^xml-v1:[0-9a-f]{8}$/.test(comparison.candidate_hash), "XML candidate hash has no version");
      assertTrue(Number(word.Selection.Start) === selectionStart && Number(word.Selection.End) === selectionEnd && Boolean(sample.Saved) === saved, "XML query changed selection or save state");
      assertTrue(captureInspectionFootprint(sample, table, false) === footprint, "XML query changed document content, objects, bookmarks or table formatting: " + name);
      if (lastCandidate !== null) assertTrue(comparison.candidate_hash !== lastCandidate, "XML comparison missed " + name);
      lastCandidate = comparison.candidate_hash;
      return result.tables[0];
    }
    var initial = inspect("initial");
    var input = fso.BuildPath(testDir, "xml-cell.txt"); writeUtf8(input, "changed");
    requireFailure(["set-cell", "--table", "1", "--cell", "1", "--input", input,
      "--expect-table-fingerprint", lastCandidate, "--expect-path", path, "--yes"]);
    assertTrue(cleanWordCell(table.Cell(1, 1)) === "A", "XML candidate authorized a write");
    runJson(["set-cell", "--table", "1", "--cell", "1", "--input", input,
      "--expect-table-fingerprint", initial.fingerprint, "--expect-path", path, "--yes"], "xml-set-cell.json");
    inspect("text");
    table.Cell(1, 1).Shading.BackgroundPatternColor = 255; inspect("shading");
    table.Cell(1, 1).Borders(-1).LineStyle = 1; table.Cell(1, 1).Borders(-1).LineWidth = 8; inspect("border");
    table.Rows.Add(); inspect("row");
    table.Columns.Add(); inspect("column");
    table.Cell(1, 1).Merge(table.Cell(1, 2)); inspect("merge");
    sample.TrackRevisions = true;
    range = table.Range.Cells(1).Range; range.Collapse(1); range.InsertBefore("TRACKED ");
    sample.TrackRevisions = false;
    assertTrue(Number(sample.Revisions.Count) > 0, "tracked XML fixture has no revision");
    inspect("revision");
    sample.Save();
    var afterSave = runJson(["tables", "--table", "1", "--compare-xml", "--expect-path", path], "xml-after-save.json");
    assertTrue(afterSave.tables[0].xml_comparison.stable_repeat === true, "XML is unstable after save");
  } finally {
    if (sample) sample.Close(false);
    previous.Activate();
  }
}

function exerciseAdvancedTable(path, tableIndex, fingerprint, prefix) {
  var cellInput = fso.BuildPath(testDir, prefix + "-advanced-cell.txt");
  writeUtf8(cellInput, "CENTER");

  var setCell = runJson([
    "set-cell", "--table", String(tableIndex), "--expect-table-fingerprint", String(fingerprint),
    "--row", "2", "--col", "2", "--input", cellInput, "--expect-path", path, "--yes"
  ], prefix + "-advanced-set-cell.json");
  assertTrue(setCell.applied === true && setCell.verified === true && setCell.inspection_complete === true, "set-cell did not verify the write");
  assertTrue(setCell.readback.matches_requested === true && setCell.readback.chars === 6, "set-cell readback did not match CENTER");
  assertTrue(String(setCell.document.path).toLowerCase() === path.toLowerCase(), "set-cell result lost document identity");
  requireFailure([
    "swap-cell-text", "--table", String(tableIndex), "--expect-table-fingerprint", String(fingerprint),
    "--from-row", "1", "--from-col", "1", "--to-row", "1", "--to-col", "3", "--expect-path", path, "--yes"
  ]);

  var swapped = runJson([
    "swap-cell-text", "--table", String(tableIndex), "--expect-table-fingerprint", String(setCell.fingerprint),
    "--from-row", "1", "--from-col", "1", "--to-row", "1", "--to-col", "3", "--expect-path", path, "--yes"
  ], prefix + "-advanced-swap.json");
  var table = doc.Tables(tableIndex);
  assertTrue(swapped.applied && swapped.verified && swapped.inspection_complete && swapped.readback.matches_requested
    && String(swapped.document.path).toLowerCase() === path.toLowerCase(), "swap result was not verified");
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
  assertTrue(shading.applied && shading.verified && shading.inspection_complete && shading.readback.matches_requested
    && shading.readback.color_value === 16315114 && shading.readback.texture === 0
    && String(shading.document.path).toLowerCase() === path.toLowerCase(), "shading result was not verified");

  var outerBorders = runJson([
    "set-table-borders", "--table", String(tableIndex), "--expect-table-fingerprint", String(shading.fingerprint),
    "--edges", "outer", "--style", "single", "--color", "4472C4", "--width", "1", "--expect-path", path, "--yes"
  ], prefix + "-advanced-outer-borders.json");
  var innerBorders = runJson([
    "set-table-borders", "--table", String(tableIndex), "--expect-table-fingerprint", String(outerBorders.fingerprint),
    "--edges", "inside-h,inside-v", "--style", "dotted", "--color", "A6A6A6", "--width", "0.5", "--expect-path", path, "--yes"
  ], prefix + "-advanced-inner-borders.json");
  requireFailure([
    "set-cell-borders", "--table", String(tableIndex), "--expect-table-fingerprint", String(outerBorders.fingerprint),
    "--row", "2", "--col", "2", "--edges", "top", "--style", "none", "--expect-path", path, "--yes"
  ]);
  table = doc.Tables(tableIndex);
  assertTrue(Number(table.Borders(-1).LineStyle) === 1, "outer table border style readback failed");
  assertTrue(Number(table.Borders(-5).LineStyle) === 2 && Number(table.Borders(-6).LineStyle) === 2, "inner table border style readback failed");
  var borderScopeBefore = captureInspectionFootprint(doc, table, true), bordersBefore = captureCellBorders(table);
  var cellBorders = runJson([
    "set-cell-borders", "--table", String(tableIndex), "--expect-table-fingerprint", String(innerBorders.fingerprint),
    "--row", "2", "--col", "2", "--edges", "top,bottom", "--style", "double", "--color", "1F4E78", "--width", "0.75", "--expect-path", path, "--yes"
  ], prefix + "-advanced-cell-borders.json");
  assertTrue(captureInspectionFootprint(doc, table, true) === borderScopeBefore, "cell borders changed text, font, shading, structure or surrounding stories");
  var bordersAfter = captureCellBorders(table);
  for (var borderIndex = 0; borderIndex < bordersBefore.length; borderIndex++) {
    // The target's top/bottom edges share boundaries with the cells above/below it.
    var allowed = borderIndex === 16 || borderIndex === 18 || borderIndex === 6 || borderIndex === 28;
    if (!allowed) assertTrue(bordersAfter[borderIndex] === bordersBefore[borderIndex], "cell borders changed an unrelated boundary: " + borderIndex);
  }
  var removedTop = runJson([
    "set-cell-borders", "--table", String(tableIndex), "--expect-table-fingerprint", String(cellBorders.fingerprint),
    "--row", "2", "--col", "2", "--edges", "top", "--style", "none", "--expect-path", path, "--yes"
  ], prefix + "-advanced-remove-top-border.json");

  table = doc.Tables(tableIndex);
  assertTrue(Number(table.Cell(2, 2).Shading.BackgroundPatternColor) === 16315114, "cell shading COM readback failed");
  assertTrue(Number(table.Borders(-1).LineStyle) === 1, "outer table border style readback failed");
  assertTrue(Number(table.Cell(2, 2).Borders(-1).LineStyle) === 0, "cell top border removal readback failed");
  assertTrue(Number(table.Cell(2, 2).Borders(-3).LineStyle) === 7, "cell bottom double border readback failed");
  var normalized = runJson([
    "normalize-table-borders", "--table", String(tableIndex), "--expect-table-fingerprint", String(removedTop.fingerprint),
    "--color", "D9DEE8", "--line-width", "4", "--expect-path", path, "--yes"
  ], prefix + "-advanced-normalize-borders.json");
  var results = [outerBorders, innerBorders, cellBorders, removedTop, normalized];
  for (var i = 0; i < results.length; i++) {
    assertTrue(results[i].applied === true && results[i].verified === true && results[i].inspection_complete === true, "border result was not verified");
    assertTrue(results[i].errors.length === 0 && results[i].failure_count === 0, "successful border result contains errors");
    assertTrue(String(results[i].document.path).toLowerCase() === path.toLowerCase(), "border result lost document identity");
  }
  assertTrue(Number(table.Cell(2, 2).Borders(-1).LineStyle) === 1 && Number(table.Cell(2, 2).Borders(-1).LineWidth) === 4, "normalized cell border readback failed");
  return String(normalized.fingerprint);
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
var savedWordOptions = {};
var successPayload = '{"ok":true,"guarded_selection":true,"guarded_tables":true,"advanced_tables":true,"guarded_equations":true,"scoped_inspection":true,"cell_text_fidelity":true,"verified_cell_results":true,"backup":true,"pdf":true,"close":true}';

function exerciseScopedInspection() {
  var queryDoc = word.Documents.Add();
  var originalFind = {}, queryFinder = null;
  try {
    queryDoc.Content.Text = "first\r\u7b2c\u4e8c\u6bb5\rthird\rfourth\rfifth\rsixth\r\uff4e\uff45\uff45\uff44\uff4c\uff45 needle ^p needle Needle\r";
    queryDoc.Footnotes.Add(queryDoc.Range(1, 1)).Range.Text = "needle footnote";
    queryDoc.Endnotes.Add(queryDoc.Range(3, 3)).Range.Text = "needle endnote";
    queryDoc.Tables.Add(queryDoc.Range(queryDoc.Content.End - 1, queryDoc.Content.End - 1), 2, 2);
    queryDoc.Content.InsertAfter("\r");
    var queryTable = queryDoc.Tables.Add(queryDoc.Range(queryDoc.Content.End - 1, queryDoc.Content.End - 1), 3, 3);
    for (var c = 1; c <= 9; c++) {
      var cellRange = queryTable.Range.Cells(c).Range;
      cellRange.End--;
      cellRange.Text = "DETAIL_" + c + " " + new Array(30).join("sample text ");
    }
    queryTable.Cell(1, 1).Merge(queryTable.Cell(1, 2));
    var mathStart = Number(queryDoc.Content.End) - 1;
    queryDoc.Range(mathStart, mathStart).Text = "x+1";
    queryDoc.OMaths.Add(queryDoc.Range(mathStart, mathStart + 3)).OMaths(1).BuildUp();
    var queryPath = fso.BuildPath(testDir, "scoped-inspection.docx");
    try { queryDoc.SaveAs2(queryPath); } catch (saveError) { queryDoc.SaveAs(queryPath); }
    queryDoc.Activate();
    queryDoc.Range(0, 5).Select();
    var textBefore = String(queryDoc.Content.Text);
    var selectionStart = Number(word.Selection.Start);
    var selectionEnd = Number(word.Selection.End);
    var savedBefore = Boolean(queryDoc.Saved);

    queryFinder = queryDoc.Content.Find;
    var findNames = ["Text", "MatchCase", "MatchWholeWord", "MatchWildcards", "Forward", "Wrap", "Format"];
    for (var setting = 0; setting < findNames.length; setting++) originalFind[findNames[setting]] = queryFinder[findNames[setting]];
    queryFinder.Text = "user query"; queryFinder.MatchCase = false; queryFinder.Wrap = 1;
    queryFinder.MatchWholeWord = true; queryFinder.MatchWildcards = true; queryFinder.Forward = false;
    var configuredFind = {};
    for (var setting = 0; setting < findNames.length; setting++) configuredFind[findNames[setting]] = queryFinder[findNames[setting]];
    var findInput = fso.BuildPath(testDir, "find-query.txt");
    writeUtf8(findInput, "needle");
    var firstHit = runJson(["find-text", "--input", findInput, "--max", "1", "--context", "2"], "find-first.json");
    assertTrue(firstHit.returned === 1 && firstHit.has_more === true && firstHit.story_type === 1, "find did not bound the first page");
    assertTrue(firstHit.matches[0].start === textBefore.indexOf("needle"), "find returned the wrong Word position");
    var secondHit = runJson(["find-text", "--input", findInput, "--from", String(firstHit.next_from)], "find-next.json");
    assertTrue(secondHit.returned === 1 && secondHit.has_more === false && secondHit.next_from === null, "find pagination repeated a match or ignored case");
    assertTrue(secondHit.matches[0].start > firstHit.matches[0].end, "find did not advance");
    var noteHit = runJson(["find-text", "--input", findInput, "--story", "endnotes"], "find-endnotes.json");
    var footHit = runJson(["find-text", "--input", findInput, "--story", "footnotes"], "find-footnotes.json");
    assertTrue(noteHit.returned === 1 && noteHit.story_type === 3 && footHit.returned === 1 && footHit.story_type === 2, "find lost note coverage");
    writeUtf8(findInput, "^p");
    var literalHit = runJson(["find-text", "--input", findInput], "find-literal.json");
    assertTrue(literalHit.returned === 1 && literalHit.matches[0].text === "^p", "find interpreted literal text as a Word special code");
    for (var setting = 0; setting < findNames.length; setting++) assertTrue(queryFinder[findNames[setting]] === configuredFind[findNames[setting]], "find changed user option: " + findNames[setting]);

    var page = runJson(["paragraphs", "--from", "2", "--max", "2"], "scoped-paragraphs.json");
    assertTrue(page.returned === 2 && page.from === 2 && page.next_from === 4, "paragraph pagination metadata mismatch");
    assertTrue(page.paragraphs[0].index === 2 && page.paragraphs[0].text === "\u7b2c\u4e8c\u6bb5\n", "paragraph page lost absolute index or Unicode text");
    assertTrue(page.paragraphs[1].text === "third\n", "paragraph page returned the wrong range");
    var empty = runJson(["paragraphs", "--from", String(Number(queryDoc.Paragraphs.Count) + 1)], "scoped-paragraphs-end.json");
    assertTrue(empty.returned === 0 && empty.next_from === null, "paragraph pagination did not stop at the end");

    var all = runJson(["tables"], "scoped-tables-all.json");
    var target = runJson(["tables", "--table", "2"], "scoped-table-full.json");
    assertTrue(all.table_count === 2 && all.returned === 2 && target.returned === 1, "single-table query changed collection counts");
    assertTrue(target.tables[0].index === 2 && target.tables[0].fingerprint === all.tables[1].fingerprint, "single-table query returned a different target");
    assertTrue(target.tables[0].layout === "irregular" && target.tables[0].linear_cells.length === 8, "targeted query lost merged-table support");
    assertTrue(target.tables[0].inspection_complete && target.tables[0].read_errors.length === 0, "targeted full inspection was incomplete");
    var cellPage = runJson(["tables", "--table", "2", "--detail", "text", "--cell-from", "3", "--cell-max", "2"], "scoped-cell-page.json").tables[0];
    assertTrue(cellPage.returned_cells === 2 && cellPage.next_cell === 5 && cellPage.cell_count === 8, "cell pagination lost merged-cell indices");
    assertTrue(cellPage.fingerprint === null && cellPage.layout === "unverified", "cell page exposed an unchecked guard or layout");
    for (var item = 0; item < 2; item++) assertTrue(cellPage.linear_cells[item].index === item + 3 && cellPage.linear_cells[item].text === target.tables[0].linear_cells[item + 2].text, "cell page returned wrong content");
    var allEquations = runJson(["equations"], "scoped-equations.json");
    var oneEquation = runJson(["equations", "--index", "1"], "scoped-equation.json");
    assertTrue(oneEquation.returned === 1 && oneEquation.equations[0].fingerprint === allEquations.equations[0].fingerprint, "targeted equation differed from full inspection");
    var textOnly = runJson(["tables", "--detail", "text"], "scoped-tables-text.json");
    assertTrue(textOnly.returned === all.returned && textOnly.detail === "text", "text inspection scope mismatch");
    for (var t = 0; t < all.tables.length; t++) {
      var expected = all.tables[t];
      var actual = textOnly.tables[t];
      assertTrue(actual.layout === expected.layout && actual.cell_count === expected.cell_count, "text inspection changed table layout");
      assertTrue(actual.fingerprint === null && actual.inspection_complete === false, "text inspection exposed a mutation guard");
      assertTrue(actual.read_errors.length === 0, "text inspection failed");
      for (var row = 0; row < expected.cells.length; row++) {
        for (var col = 0; col < expected.cells[row].length; col++) {
          assertTrue(actual.cells[row][col] === expected.cells[row][col], "text inspection changed rectangular cell text");
        }
      }
      assertTrue(actual.linear_cells.length === expected.linear_cells.length, "text inspection lost merged cells");
      for (var cell = 0; cell < expected.linear_cells.length; cell++) {
        var a = actual.linear_cells[cell], b = expected.linear_cells[cell];
        assertTrue(a.index === b.index && a.row === b.row && a.col === b.col && a.text === b.text, "text inspection changed merged cell metadata");
      }
    }
    var summary = runJson(["tables", "--table", "2", "--detail", "summary"], "scoped-table-summary.json");
    assertTrue(summary.detail === "summary" && summary.returned === 1 && summary.tables[0].index === 2, "summary scope mismatch");
    assertTrue(summary.tables[0].cell_count === 8 && summary.tables[0].layout === "unverified", "summary guessed a verified layout");
    assertTrue(summary.tables[0].fingerprint === null && summary.tables[0].inspection_complete === false, "summary exposed a mutation guard");
    assertTrue(summary.tables[0].cells.length === 0 && summary.tables[0].linear_cells.length === 0, "summary returned cell contents");
    assertTrue(fso.GetFile(fso.BuildPath(testDir, "scoped-table-summary.json")).Size < fso.GetFile(fso.BuildPath(testDir, "scoped-table-full.json")).Size, "summary did not reduce response size");
    requireFailure(["tables", "--table", "2", "--max", "1"]);
    requireFailure(["tables", "--table", "3"]);
    assertTrue(String(queryDoc.Content.Text) === textBefore && Boolean(queryDoc.Saved) === savedBefore, "read-only inspection changed the document");
    assertTrue(Number(word.Selection.Start) === selectionStart && Number(word.Selection.End) === selectionEnd, "read-only inspection changed the selection");
    assertTrue(String(word.ActiveDocument.FullName).toLowerCase() === queryPath.toLowerCase(), "read-only inspection switched the active document");
    var cellInput = fso.BuildPath(testDir, "verified-multiline.txt");
    writeUtf8(cellInput, "\u4e2d\u6587\r\nsecond\n");
    var multiline = runJson(["set-cell", "--table", "1", "--row", "1", "--col", "1", "--input", cellInput,
      "--expect-path", queryPath, "--expect-table-fingerprint", all.tables[0].fingerprint, "--yes"], "verified-multiline.json");
    assertTrue(multiline.verified && multiline.readback.matches_requested, "multiline receipt rejected Word newline normalization");
    assertTrue(String(queryDoc.Tables(1).Cell(1, 1).Range.Text) === "\u4e2d\u6587\rsecond\r\r\x07", "multiline cell lost content or trailing paragraph");
    writeUtf8(cellInput, "");
    var emptyCell = runJson(["set-cell", "--table", "1", "--row", "1", "--col", "1", "--input", cellInput,
      "--expect-path", queryPath, "--expect-table-fingerprint", multiline.fingerprint, "--yes"], "verified-empty-cell.json");
    assertTrue(emptyCell.verified && emptyCell.readback.chars === 0, "empty cell was not verified");
  } finally {
    if (queryFinder) for (var setting in originalFind) queryFinder[setting] = originalFind[setting];
    queryDoc.Close(false);
    queryDoc = null;
    doc.Activate();
  }
}

function runFixture(path) {
  var fixtureExtension = "." + fso.GetExtensionName(path);
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
  // These options persist across Word instances; restore them when the test exits.
  var optionNames = ["UpdateLinksAtOpen", "CheckGrammarAsYouType", "CheckSpellingAsYouType"];
  for (var o = 0; o < optionNames.length; o++) {
    var optionName = optionNames[o];
    try {
      savedWordOptions[optionName] = word.Options[optionName];
      word.Options[optionName] = false;
    } catch (optionError) {}
  }

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
  var originalText = String(doc.Content.Text);
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
    "--expect-story-type", String(collapsed.selection.story_type),
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
    "--expect-story-type", String(selectedSentinel.selection.story_type),
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
    "--expect-story-type", String(selectedSentinel.selection.story_type),
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
  assertTrue(createdTable.applied === true && createdTable.fill_complete === true && createdTable.verified === true, "fixture table creation was not verified");
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

  var unsavedBackupPath = fso.BuildPath(testDir, "fixture-unsaved-copy" + fixtureExtension);
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
  var savedBackupPath = fso.BuildPath(testDir, "fixture-saved-backup" + fixtureExtension);
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

  // Reopen the saved copy: a command receipt alone cannot prove persistence.
  doc = word.Documents.Open(path, false, true, false);
  assertTrue(String(doc.Range(0, insertionPosition).Text) === originalText.substring(0, originalText.length - 1), "fixture original body text changed after save/reopen");
  assertTrue(Number(doc.Tables.Count) === tableCount && Number(doc.OMaths.Count) === equationCount, "fixture original object counts changed after save/reopen");
  assertTrue(Number(doc.Sections.Count) === sectionCount, "fixture sections changed after save/reopen");
  assertTrue(Number(doc.InlineShapes.Count) === inlineShapeCount && Number(doc.Shapes.Count) === shapeCount, "fixture images changed after save/reopen");
  assertTrue(Number(doc.Comments.Count) === commentCount + 1 && Number(doc.Revisions.Count) > revisionCount, "fixture comments/revisions were not persisted");
  doc.Close(false);
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
    + "\"original_content_preserved\":true,"
    + "\"saved_readback\":true,"
    + "\"unsaved_backup_behavior\":" + '"' + unsavedBackupBehavior + '"'
    + "}";
}

try {
  if (fixtureCopy) {
    successPayload = runFixture(fixtureCopy);
  } else {
  var docPath = fso.BuildPath(testDir, "command-integration.docx");
  exerciseXmlNormalization();
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
  doc.Content.Text = "seed text\rsoft\x0bline\rpage\fbreak\r";
  try { doc.SaveAs2(docPath); } catch (saveError) { doc.SaveAs(docPath); }
  exerciseXmlMutationDetection();
  exerciseCellTextFidelity();
  exerciseVerifiedCellResults();

  word.Selection.SetRange(0, 4);
  runJson(["paragraphs"], "control-character-paragraphs.json");
  var selection = runJson(["selection-info"], "selection-info.json");
  assertTrue(String(selection.document.path).toLowerCase() === docPath.toLowerCase(), "selection inspection targeted the wrong document");
  assertTrue(String(selection.selection.text_hash).length === 8, "selection hash was not emitted");

  var selectionArgs = [
    "--expect-story-type", String(selection.selection.story_type),
    "--expect-start", String(selection.selection.start),
    "--expect-end", String(selection.selection.end),
    "--expect-selection-hash", String(selection.selection.text_hash)
  ];
  requireFailure(["replace-selection", "--input", replacementInput, "--expect-path", wrongPath].concat(selectionArgs, ["--yes"]));
  requireFailure(["replace-selection", "--input", replacementInput, "--expect-name", String(doc.Name)].concat(selectionArgs, ["--yes"]));
  requireFailure(["replace-selection", "--input", replacementInput, "--expect-path", docPath, "--output", docPath, "--overwrite"].concat(selectionArgs, ["--yes"]));
  assertTrue(String(doc.Content.Text).indexOf("seed") === 0, "output preflight changed the document");

  var header = doc.Sections(1).Headers(1).Range;
  header.Text = "seed";
  header.End = header.Start + 4;
  header.Select();
  var headerSelection = runJson(["selection-info"], "header-selection.json");
  assertTrue(headerSelection.selection.story_type !== selection.selection.story_type, "header did not expose a different story");
  requireFailure(["replace-selection", "--input", replacementInput, "--expect-path", docPath].concat(selectionArgs, ["--yes"]));
  assertTrue(String(doc.Sections(1).Headers(1).Range.Text).indexOf("seed") === 0, "stale main-story guards changed the header");
  doc.Range(0, 4).Select();


  var replacement = runJson([
    "replace-selection", "--input", replacementInput, "--expect-path", docPath
  ].concat(selectionArgs, ["--yes"]), "replace.json");
  assertTrue(replacement.ok && String(doc.Content.Text).indexOf("edited") === 0, "guarded selection replacement failed");
  requireFailure(["insert-comment", "--input", replacementInput, "--expect-path", docPath].concat(selectionArgs, ["--yes"]));

  var createdTable = runJson([
    "create-table", "--rows", "3", "--cols", "3", "--input", tableInput, "--at", "end", "--expect-path", docPath, "--yes"
  ], "create-table.json");
  assertTrue(createdTable.ok && String(createdTable.fingerprint).length === 8, "guarded table creation failed");
  assertTrue(createdTable.applied === true && createdTable.fill_complete === true && createdTable.verified === true, "table creation was not verified");

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

  exerciseScopedInspection();

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
    for (var optionName in savedWordOptions) {
      try { word.Options[optionName] = savedWordOptions[optionName]; }
      catch (restoreError) { if (failure === null) failure = restoreError; }
    }
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
