// Windows Script Host JScript bridge for Microsoft Word desktop.
// Run with: cscript //nologo scripts\word_control.js <command> [options]

var fso = new ActiveXObject("Scripting.FileSystemObject");

function argList() {
  var out = [];
  for (var i = 0; i < WScript.Arguments.length; i++) out.push(String(WScript.Arguments(i)));
  return out;
}

var ARGS = argList();
var COMMAND = ARGS.length ? ARGS[0] : "help";

function hasFlag(flag) {
  for (var i = 1; i < ARGS.length; i++) if (ARGS[i] === flag) return true;
  return false;
}

function opt(name, fallbackValue) {
  for (var i = 1; i < ARGS.length - 1; i++) {
    if (ARGS[i] === name) return ARGS[i + 1];
  }
  return fallbackValue;
}

function die(message, code) {
  WScript.StdErr.WriteLine("word-control error: " + message);
  WScript.Quit(code || 1);
}

function requireYes() {
  if (!hasFlag("--yes")) die("mutation command requires --yes");
}

function absPath(path) {
  if (!path) return "";
  return fso.GetAbsolutePathName(path);
}

function readUtf8(path) {
  if (!path) die("missing input path");
  path = absPath(path);
  if (!fso.FileExists(path)) die("input file not found: " + path);
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
  path = absPath(path);
  var parent = fso.GetParentFolderName(path);
  if (parent && !fso.FolderExists(parent)) die("output folder not found: " + parent);
  var stream = new ActiveXObject("ADODB.Stream");
  stream.Type = 2;
  stream.Charset = "utf-8";
  stream.Open();
  stream.WriteText(String(text));
  stream.SaveToFile(path, 2);
  stream.Close();
}

function emit(text) {
  var output = opt("--output", "");
  if (output) writeUtf8(output, text);
  else WScript.Echo(text);
}

function jescape(value) {
  if (value === null || typeof value === "undefined") return "";
  return String(value)
    .replace(/\\/g, "\\\\")
    .replace(/"/g, "\\\"")
    .replace(/\r/g, "\\r")
    .replace(/\n/g, "\\n")
    .replace(/\t/g, "\\t");
}

function q(value) {
  return "\"" + jescape(value) + "\"";
}

function normalizeText(value) {
  if (value === null || typeof value === "undefined") return "";
  return String(value).replace(/\r/g, "\n");
}

function parsePositiveInt(value, name) {
  var n = parseInt(value, 10);
  if (!n || n < 1) die(name + " must be a positive integer");
  return n;
}

function parseHexColor(value) {
  value = String(value || "").replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(value)) die("color must be a 6-digit RGB hex value");
  var r = parseInt(value.substring(0, 2), 16);
  var g = parseInt(value.substring(2, 4), 16);
  var b = parseInt(value.substring(4, 6), 16);
  return r + (g * 256) + (b * 65536);
}

function setComBorder(border, color, lineWidth) {
  border.LineStyle = 1; // wdLineStyleSingle
  border.LineWidth = lineWidth;
  border.Color = color;
}

function setBorderCollection(borders, borderTypes, color, lineWidth) {
  for (var i = 0; i < borderTypes.length; i++) {
    try { setComBorder(borders(borderTypes[i]), color, lineWidth); } catch (e) {}
  }
}

function cleanCellText(value) {
  return normalizeText(value).replace(/\x07/g, "").replace(/\n+$/g, "");
}

function setCellText(cell, text) {
  var range = cell.Range;
  range.End = range.End - 1;
  range.Text = text;
}

function fillTableFromTsv(table, text) {
  var rows = text.replace(/\r/g, "").split("\n");
  var rowCount = Number(table.Rows.Count);
  var colCount = Number(table.Columns.Count);
  for (var r = 0; r < rows.length && r < rowCount; r++) {
    if (rows[r] === "" && r === rows.length - 1) continue;
    var cells = rows[r].split("\t");
    for (var c = 0; c < cells.length && c < colCount; c++) {
      setCellText(table.Cell(r + 1, c + 1), cells[c]);
    }
  }
}

function matchingBrace(text, openIndex) {
  var depth = 0;
  for (var i = openIndex; i < text.length; i++) {
    var ch = text.charAt(i);
    if (ch === "{") depth++;
    else if (ch === "}") {
      depth--;
      if (depth === 0) return i;
    }
  }
  return -1;
}

function needsGrouping(text) {
  return /[\s+\-*/=<>]/.test(text);
}

function compactGroup(text) {
  text = convertLatexToWordLinear(text);
  return needsGrouping(text) ? "(" + text + ")" : text;
}

function convertLatexCommandWithTwoGroups(text, command, joiner) {
  var needle = "\\" + command + "{";
  var pos = text.indexOf(needle);
  while (pos >= 0) {
    var firstOpen = pos + command.length + 1;
    var firstClose = matchingBrace(text, firstOpen);
    if (firstClose < 0 || text.charAt(firstClose + 1) !== "{") break;
    var secondOpen = firstClose + 1;
    var secondClose = matchingBrace(text, secondOpen);
    if (secondClose < 0) break;
    var a = text.substring(firstOpen + 1, firstClose);
    var b = text.substring(secondOpen + 1, secondClose);
    var replacement = compactGroup(a) + joiner + compactGroup(b);
    text = text.substring(0, pos) + replacement + text.substring(secondClose + 1);
    pos = text.indexOf(needle, pos + replacement.length);
  }
  return text;
}

function convertLatexCommandWithOneGroup(text, command, prefix, suffix) {
  var needle = "\\" + command + "{";
  var pos = text.indexOf(needle);
  while (pos >= 0) {
    var open = pos + command.length + 1;
    var close = matchingBrace(text, open);
    if (close < 0) break;
    var body = convertLatexToWordLinear(text.substring(open + 1, close));
    var replacement = prefix + body + suffix;
    text = text.substring(0, pos) + replacement + text.substring(close + 1);
    pos = text.indexOf(needle, pos + replacement.length);
  }
  return text;
}

function convertSupSubGroups(text) {
  text = text.replace(/\^\{([^{}]+)\}/g, function(_, body) {
    return "^" + (needsGrouping(body) ? "(" + convertLatexToWordLinear(body) + ")" : convertLatexToWordLinear(body));
  });
  text = text.replace(/_\{([^{}]+)\}/g, function(_, body) {
    return "_" + (needsGrouping(body) ? "(" + convertLatexToWordLinear(body) + ")" : convertLatexToWordLinear(body));
  });
  return text;
}

function convertLatexToWordLinear(text) {
  if (text === null || typeof text === "undefined") return "";
  text = String(text).replace(/\r/g, "").replace(/\n/g, " ");
  text = text.replace(/\$\$/g, "").replace(/\$/g, "");
  text = text.replace(/\\left/g, "").replace(/\\right/g, "");
  text = convertLatexCommandWithTwoGroups(text, "frac", "/");
  text = convertLatexCommandWithOneGroup(text, "sqrt", "\u221A(", ")");
  text = convertSupSubGroups(text);
  var replacements = [
    [/\\int/g, "\u222B"],
    [/\\sum/g, "\u2211"],
    [/\\prod/g, "\u220F"],
    [/\\cdot/g, "\u00B7"],
    [/\\times/g, "\u00D7"],
    [/\\leq?/g, "\u2264"],
    [/\\geq?/g, "\u2265"],
    [/\\neq/g, "\u2260"],
    [/\\approx/g, "\u2248"],
    [/\\pm/g, "\u00B1"],
    [/\\alpha/g, "\u03B1"],
    [/\\beta/g, "\u03B2"],
    [/\\gamma/g, "\u03B3"],
    [/\\delta/g, "\u03B4"],
    [/\\epsilon/g, "\u03B5"],
    [/\\theta/g, "\u03B8"],
    [/\\lambda/g, "\u03BB"],
    [/\\mu/g, "\u03BC"],
    [/\\pi/g, "\u03C0"],
    [/\\rho/g, "\u03C1"],
    [/\\sigma/g, "\u03C3"],
    [/\\tau/g, "\u03C4"],
    [/\\phi/g, "\u03C6"],
    [/\\omega/g, "\u03C9"],
    [/\\Delta/g, "\u0394"],
    [/\\Omega/g, "\u03A9"]
  ];
  for (var i = 0; i < replacements.length; i++) {
    text = text.replace(replacements[i][0], replacements[i][1]);
  }
  text = text.replace(/[{}]/g, "");
  return text.replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
}

function equationInputText() {
  var input = opt("--input", "");
  var text = readUtf8(input);
  var format = opt("--format", "latex");
  if (format === "latex") return convertLatexToWordLinear(text);
  if (format === "linear" || format === "word") return text.replace(/\r/g, "").replace(/\n/g, " ");
  die("--format must be latex, linear, or word");
}

function buildEquationInRange(doc, range, linearText) {
  var start = range.Start;
  range.Text = linearText;
  var mathRange = doc.Range(start, start + linearText.length);
  var eqRange = doc.OMaths.Add(mathRange);
  eqRange.OMaths(1).BuildUp();
  return eqRange;
}

function getWord() {
  try {
    return GetObject("", "Word.Application");
  } catch (e) {
    die("no running Word instance found; open a Word document first");
  }
}

function getActiveDocument(word) {
  if (word.Documents.Count < 1) die("Word is running but no document is open");
  return word.ActiveDocument;
}

function safeDocPath(doc) {
  try {
    return String(doc.FullName);
  } catch (e) {
    return "";
  }
}

function activeDocJson(word, doc) {
  var path = safeDocPath(doc);
  var saved = "";
  var track = "";
  try { saved = String(doc.Saved); } catch (e1) {}
  try { track = String(doc.TrackRevisions); } catch (e2) {}
  return "{"
    + "\"name\":" + q(doc.Name) + ","
    + "\"path\":" + q(path) + ","
    + "\"saved\":" + q(saved) + ","
    + "\"track_revisions\":" + q(track)
    + "}";
}

function commandHelp() {
  emit([
    "word-control commands:",
    "  status [--output file]",
    "  selection [--output file]",
    "  document-text [--max chars|--full] [--output file]",
    "  paragraphs [--max count] [--output file]",
    "  tables [--output file]",
    "  equations [--output file]",
    "  enable-track-changes --yes",
    "  disable-track-changes --yes",
    "  replace-selection --input file [--track] --yes",
    "  replace-paragraph --index n --input file [--track] --yes",
    "  insert-comment --input file --yes",
    "  create-table --rows n --cols n [--input table.tsv] [--at selection|end] --yes",
    "  set-cell --table n --row n --col n --input file --yes",
    "  delete-table --table n --yes",
    "  normalize-table-borders --table n [--color D9DEE8] [--line-width 4] --yes",
    "  insert-equation --input file [--format latex|linear] [--at selection|end] --yes",
    "  set-equation --index n --input file [--format latex|linear] --yes",
    "  delete-equation --index n --yes",
    "  save-active --yes",
    "  close-active --save|--discard [--quit-if-empty] --yes",
    "  save-copy --path file --yes",
    "  export-pdf --path file --yes",
    "  open --path file",
    "  smoke --path file"
  ].join("\n"));
}

function commandStatus() {
  var word = getWord();
  var docs = Number(word.Documents.Count);
  var selectionText = "";
  try { selectionText = String(word.Selection.Text); } catch (e1) {}
  var docPart = "null";
  if (docs > 0) docPart = activeDocJson(word, word.ActiveDocument);
  emit("{"
    + "\"ok\":true,"
    + "\"word_version\":" + q(word.Version) + ","
    + "\"documents_count\":" + docs + ","
    + "\"active_document\":" + docPart + ","
    + "\"selection_text_length\":" + normalizeText(selectionText).length
    + "}");
}

function commandSelection() {
  var word = getWord();
  getActiveDocument(word);
  var text = "";
  try { text = normalizeText(word.Selection.Text); } catch (e) { die("cannot read current selection: " + e.message); }
  emit(text);
}

function commandDocumentText() {
  var word = getWord();
  var doc = getActiveDocument(word);
  var text = normalizeText(doc.Content.Text);
  if (!hasFlag("--full")) {
    var max = parseInt(opt("--max", "20000"), 10);
    if (max > 0 && text.length > max) text = text.substring(0, max) + "\n[truncated at " + max + " chars]";
  }
  emit(text);
}

function commandParagraphs() {
  var word = getWord();
  var doc = getActiveDocument(word);
  var max = parseInt(opt("--max", "80"), 10);
  var count = Number(doc.Paragraphs.Count);
  if (max < 1 || max > count) max = count;
  var parts = [];
  for (var i = 1; i <= max; i++) {
    var text = "";
    try { text = normalizeText(doc.Paragraphs(i).Range.Text); } catch (e) {}
    parts.push("{\"index\":" + i + ",\"text\":" + q(text) + "}");
  }
  emit("{\"ok\":true,\"paragraph_count\":" + count + ",\"returned\":" + max + ",\"paragraphs\":[" + parts.join(",") + "]}");
}

function commandTables() {
  var word = getWord();
  var doc = getActiveDocument(word);
  var count = Number(doc.Tables.Count);
  var parts = [];
  for (var i = 1; i <= count; i++) {
    var table = doc.Tables(i);
    var rowCount = Number(table.Rows.Count);
    var colCount = Number(table.Columns.Count);
    var cells = [];
    for (var r = 1; r <= rowCount; r++) {
      var rowCells = [];
      for (var c = 1; c <= colCount; c++) {
        var text = "";
        try { text = cleanCellText(table.Cell(r, c).Range.Text); } catch (e) {}
        rowCells.push(q(text));
      }
      cells.push("[" + rowCells.join(",") + "]");
    }
    parts.push("{\"index\":" + i + ",\"rows\":" + rowCount + ",\"cols\":" + colCount + ",\"cells\":[" + cells.join(",") + "]}");
  }
  emit("{\"ok\":true,\"table_count\":" + count + ",\"tables\":[" + parts.join(",") + "]}");
}

function commandEquations() {
  var word = getWord();
  var doc = getActiveDocument(word);
  var count = Number(doc.OMaths.Count);
  var parts = [];
  for (var i = 1; i <= count; i++) {
    var text = "";
    try { text = normalizeText(doc.OMaths(i).Range.Text); } catch (e) {}
    parts.push("{\"index\":" + i + ",\"text\":" + q(text) + "}");
  }
  emit("{\"ok\":true,\"equation_count\":" + count + ",\"equations\":[" + parts.join(",") + "]}");
}

function commandTrack(on) {
  requireYes();
  var word = getWord();
  var doc = getActiveDocument(word);
  doc.TrackRevisions = on;
  emit("{\"ok\":true,\"track_revisions\":" + q(String(doc.TrackRevisions)) + "}");
}

function commandReplaceSelection() {
  requireYes();
  var input = opt("--input", "");
  var text = readUtf8(input);
  var word = getWord();
  var doc = getActiveDocument(word);
  if (hasFlag("--track")) doc.TrackRevisions = true;
  try {
    word.Selection.Range.Text = text;
  } catch (e) {
    die("failed to replace selection: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"replace-selection\",\"chars\":" + text.length + ",\"track_revisions\":" + q(String(doc.TrackRevisions)) + "}");
}

function commandReplaceParagraph() {
  requireYes();
  var idx = parseInt(opt("--index", "0"), 10);
  if (!idx || idx < 1) die("replace-paragraph requires --index n");
  var input = opt("--input", "");
  var text = readUtf8(input);
  var word = getWord();
  var doc = getActiveDocument(word);
  var count = Number(doc.Paragraphs.Count);
  if (idx > count) die("paragraph index out of range; document has " + count + " paragraphs");
  if (hasFlag("--track")) doc.TrackRevisions = true;
  if (!/\r$/.test(text)) text = text + "\r";
  try {
    doc.Paragraphs(idx).Range.Text = text;
  } catch (e) {
    die("failed to replace paragraph: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"replace-paragraph\",\"index\":" + idx + ",\"chars\":" + text.length + ",\"track_revisions\":" + q(String(doc.TrackRevisions)) + "}");
}

function commandInsertComment() {
  requireYes();
  var input = opt("--input", "");
  var text = readUtf8(input);
  var word = getWord();
  var doc = getActiveDocument(word);
  try {
    doc.Comments.Add(word.Selection.Range, text);
  } catch (e) {
    die("failed to insert comment: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"insert-comment\",\"chars\":" + text.length + "}");
}

function commandCreateTable() {
  requireYes();
  var rows = parsePositiveInt(opt("--rows", "0"), "--rows");
  var cols = parsePositiveInt(opt("--cols", "0"), "--cols");
  var at = opt("--at", "selection");
  var input = opt("--input", "");
  var word = getWord();
  var doc = getActiveDocument(word);
  var range;
  if (at === "end") {
    range = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
  } else {
    range = word.Selection.Range;
  }
  var table;
  try {
    table = doc.Tables.Add(range, rows, cols);
  } catch (e) {
    die("failed to create table: " + e.message);
  }
  if (input) fillTableFromTsv(table, readUtf8(input));
  emit("{\"ok\":true,\"action\":\"create-table\",\"table_count\":" + Number(doc.Tables.Count) + ",\"rows\":" + rows + ",\"cols\":" + cols + "}");
}

function commandSetCell() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var row = parsePositiveInt(opt("--row", "0"), "--row");
  var col = parsePositiveInt(opt("--col", "0"), "--col");
  var input = opt("--input", "");
  var text = readUtf8(input);
  var word = getWord();
  var doc = getActiveDocument(word);
  var tableCount = Number(doc.Tables.Count);
  if (tableIndex > tableCount) die("table index out of range; document has " + tableCount + " tables");
  var table = doc.Tables(tableIndex);
  if (row > Number(table.Rows.Count)) die("row index out of range; table has " + Number(table.Rows.Count) + " rows");
  if (col > Number(table.Columns.Count)) die("column index out of range; table has " + Number(table.Columns.Count) + " columns");
  try {
    setCellText(table.Cell(row, col), text);
  } catch (e) {
    die("failed to set cell: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"set-cell\",\"table\":" + tableIndex + ",\"row\":" + row + ",\"col\":" + col + ",\"chars\":" + text.length + "}");
}

function commandDeleteTable() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var word = getWord();
  var doc = getActiveDocument(word);
  var tableCount = Number(doc.Tables.Count);
  if (tableIndex > tableCount) die("table index out of range; document has " + tableCount + " tables");
  try {
    doc.Tables(tableIndex).Delete();
  } catch (e) {
    die("failed to delete table: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"delete-table\",\"remaining_tables\":" + Number(doc.Tables.Count) + "}");
}

function commandNormalizeTableBorders() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "1"), "--table");
  var colorHex = opt("--color", "D9DEE8");
  var lineWidth = parseInt(opt("--line-width", "4"), 10);
  if (!lineWidth || lineWidth < 1) die("--line-width must be a positive integer Word line-width constant");
  var color = parseHexColor(colorHex);
  var word = getWord();
  var doc = getActiveDocument(word);
  var tableCount = Number(doc.Tables.Count);
  if (tableIndex > tableCount) die("table index out of range; document has " + tableCount + " tables");
  var table = doc.Tables(tableIndex);
  var tableBorderTypes = [-1, -2, -3, -4, -5, -6]; // top, left, bottom, right, insideH, insideV
  var cellBorderTypes = [-1, -2, -3, -4];
  var cellCount = 0;
  var cellErrors = 0;

  setBorderCollection(table.Borders, tableBorderTypes, color, lineWidth);
  try { table.Borders.Enable = true; } catch (e1) {}

  try {
    cellCount = Number(table.Range.Cells.Count);
    for (var i = 1; i <= cellCount; i++) {
      try {
        setBorderCollection(table.Range.Cells(i).Borders, cellBorderTypes, color, lineWidth);
      } catch (e2) {
        cellErrors++;
      }
    }
  } catch (e3) {
    cellErrors++;
  }

  emit("{\"ok\":true,\"action\":\"normalize-table-borders\",\"table\":" + tableIndex
    + ",\"tables\":" + tableCount
    + ",\"cells_seen\":" + cellCount
    + ",\"cell_errors\":" + cellErrors
    + ",\"color\":" + q(colorHex)
    + ",\"line_width\":" + lineWidth + "}");
}

function commandInsertEquation() {
  requireYes();
  var linearText = equationInputText();
  if (!linearText) die("equation input is empty after conversion");
  var at = opt("--at", "selection");
  var word = getWord();
  var doc = getActiveDocument(word);
  var range;
  if (at === "end") range = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
  else range = word.Selection.Range;
  try {
    buildEquationInRange(doc, range, linearText);
  } catch (e) {
    die("failed to insert equation: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"insert-equation\",\"linear\":" + q(linearText) + ",\"equation_count\":" + Number(doc.OMaths.Count) + "}");
}

function commandSetEquation() {
  requireYes();
  var idx = parsePositiveInt(opt("--index", "0"), "--index");
  var linearText = equationInputText();
  if (!linearText) die("equation input is empty after conversion");
  var word = getWord();
  var doc = getActiveDocument(word);
  var count = Number(doc.OMaths.Count);
  if (idx > count) die("equation index out of range; document has " + count + " equations");
  try {
    var oldRange = doc.OMaths(idx).Range;
    var range = doc.Range(oldRange.Start, oldRange.End);
    buildEquationInRange(doc, range, linearText);
  } catch (e) {
    die("failed to set equation: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"set-equation\",\"index\":" + idx + ",\"linear\":" + q(linearText) + ",\"equation_count\":" + Number(doc.OMaths.Count) + "}");
}

function commandDeleteEquation() {
  requireYes();
  var idx = parsePositiveInt(opt("--index", "0"), "--index");
  var word = getWord();
  var doc = getActiveDocument(word);
  var count = Number(doc.OMaths.Count);
  if (idx > count) die("equation index out of range; document has " + count + " equations");
  try {
    doc.OMaths(idx).Range.Delete();
  } catch (e) {
    die("failed to delete equation: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"delete-equation\",\"remaining_equations\":" + Number(doc.OMaths.Count) + "}");
}

function commandSaveActive() {
  requireYes();
  var word = getWord();
  var doc = getActiveDocument(word);
  try {
    doc.Save();
  } catch (e) {
    die("failed to save active document: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"save-active\",\"path\":" + q(safeDocPath(doc)) + "}");
}

function commandCloseActive() {
  requireYes();
  var save = hasFlag("--save");
  var discard = hasFlag("--discard");
  if (save === discard) die("close-active requires exactly one of --save or --discard");
  var word = getWord();
  var doc = getActiveDocument(word);
  var path = safeDocPath(doc);
  try {
    if (save) doc.Save();
    doc.Close(false);
    if (hasFlag("--quit-if-empty") && Number(word.Documents.Count) === 0) {
      word.Quit(0);
      word = null;
      try { CollectGarbage(); } catch (e2) {}
    }
  } catch (e) {
    die("failed to close active document: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"close-active\",\"saved\":" + q(String(save)) + ",\"path\":" + q(path) + "}");
}

function commandSaveCopy() {
  requireYes();
  var path = absPath(opt("--path", ""));
  if (!path) die("save-copy requires --path file");
  var word = getWord();
  var doc = getActiveDocument(word);
  try {
    doc.SaveCopyAs(path);
    emit("{\"ok\":true,\"action\":\"save-copy\",\"method\":\"SaveCopyAs\",\"path\":" + q(path) + "}");
    return;
  } catch (e) {
    var source = safeDocPath(doc);
    if (!source || !fso.FileExists(source)) die("failed to save copy and active document has no saved source path: " + e.message);
    try {
      fso.CopyFile(source, path, true);
      emit("{\"ok\":true,\"action\":\"save-copy\",\"method\":\"FileSystemObject.CopyFile\",\"path\":" + q(path) + ",\"source\":" + q(source) + "}");
      return;
    } catch (e2) {
      die("failed to save copy: " + e.message + "; fallback failed: " + e2.message);
    }
  }
}

function commandExportPdf() {
  requireYes();
  var path = absPath(opt("--path", ""));
  if (!path) die("export-pdf requires --path file");
  var word = getWord();
  var doc = getActiveDocument(word);
  try {
    doc.ExportAsFixedFormat(path, 17);
  } catch (e) {
    die("failed to export PDF: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"export-pdf\",\"path\":" + q(path) + "}");
}

function commandOpen() {
  var path = absPath(opt("--path", ""));
  if (!path || !fso.FileExists(path)) die("open requires existing --path file");
  var word;
  try {
    word = GetObject("", "Word.Application");
  } catch (e) {
    word = new ActiveXObject("Word.Application");
  }
  word.Visible = true;
  word.Documents.Open(path);
  emit("{\"ok\":true,\"action\":\"open\",\"path\":" + q(path) + "}");
}

function commandSmoke() {
  var path = absPath(opt("--path", ""));
  if (!path) die("smoke requires --path file");
  var word = new ActiveXObject("Word.Application");
  word.Visible = false;
  word.DisplayAlerts = 0;
  var doc = word.Documents.Add();
  doc.Content.Text = "Word control smoke test.\rCreated by word_control.js.\r";
  try {
    doc.SaveAs2(path);
  } catch (e1) {
    doc.SaveAs(path);
  }
  doc.Close(false);
  word.Quit(0);
  doc = null;
  word = null;
  try { CollectGarbage(); } catch (e2) {}
  emit("{\"ok\":true,\"action\":\"smoke\",\"path\":" + q(path) + "}");
}

try {
  if (COMMAND === "help" || COMMAND === "--help" || COMMAND === "-h") commandHelp();
  else if (COMMAND === "status") commandStatus();
  else if (COMMAND === "selection") commandSelection();
  else if (COMMAND === "document-text") commandDocumentText();
  else if (COMMAND === "paragraphs") commandParagraphs();
  else if (COMMAND === "tables") commandTables();
  else if (COMMAND === "equations") commandEquations();
  else if (COMMAND === "enable-track-changes") commandTrack(true);
  else if (COMMAND === "disable-track-changes") commandTrack(false);
  else if (COMMAND === "replace-selection") commandReplaceSelection();
  else if (COMMAND === "replace-paragraph") commandReplaceParagraph();
  else if (COMMAND === "insert-comment") commandInsertComment();
  else if (COMMAND === "create-table") commandCreateTable();
  else if (COMMAND === "set-cell") commandSetCell();
  else if (COMMAND === "delete-table") commandDeleteTable();
  else if (COMMAND === "normalize-table-borders") commandNormalizeTableBorders();
  else if (COMMAND === "insert-equation") commandInsertEquation();
  else if (COMMAND === "set-equation") commandSetEquation();
  else if (COMMAND === "delete-equation") commandDeleteEquation();
  else if (COMMAND === "save-active") commandSaveActive();
  else if (COMMAND === "close-active") commandCloseActive();
  else if (COMMAND === "save-copy") commandSaveCopy();
  else if (COMMAND === "export-pdf") commandExportPdf();
  else if (COMMAND === "open") commandOpen();
  else if (COMMAND === "smoke") commandSmoke();
  else die("unknown command: " + COMMAND, 2);
} catch (e) {
  die(e.message || String(e), 1);
}
