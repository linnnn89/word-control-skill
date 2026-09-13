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
    if (ARGS[i] === name) {
      if (/^--/.test(ARGS[i + 1])) die("missing value for " + name);
      return ARGS[i + 1];
    }
  }
  if (ARGS.length > 1 && ARGS[ARGS.length - 1] === name) die("missing value for " + name);
  return fallbackValue;
}

function die(message, code) {
  WScript.StdErr.WriteLine("word-control error: " + message);
  WScript.Quit(code || 1);
}

function requireYes() {
  if (!hasFlag("--yes")) die("mutation command requires --yes");
}

function failJson(payload, code) {
  emit(payload);
  WScript.Quit(code || 1);
}

function absPath(path) {
  if (!path) return "";
  return fso.GetAbsolutePathName(path);
}

function canonicalPath(path) {
  if (!path) return "";
  return absPath(path).replace(/\//g, "\\").toLowerCase();
}

function prepareOutputPath(path, label) {
  path = absPath(path);
  if (!path) die(label + " requires an output path");
  var parent = fso.GetParentFolderName(path);
  if (parent && !fso.FolderExists(parent)) die("output folder not found: " + parent);
  if (fso.FileExists(path) && !hasFlag("--overwrite")) {
    die("refusing to overwrite existing output; choose a new path or pass --overwrite: " + path);
  }
  return path;
}

function prepareDocumentOutputPath(path, label) {
  // FSO MoveFile accepts wildcards; reject them, alternate streams and Win32 aliases.
  if (!path || /[\x00-\x1f*?"<>|]/.test(path) || /:/.test(path.replace(/^[A-Za-z]:/, ""))
      || /[. ](?:[\\\/]|$)/.test(path.replace(/(^|[\\\/])\.{1,2}(?=[\\\/])/g, "$1"))) {
    die(label + " requires a literal file path without wildcards, streams or trailing dots/spaces");
  }
  path = prepareOutputPath(path, label);
  if (fso.FolderExists(path)) die("output path is a folder: " + path);
  return path;
}

function createOutputStage(path) {
  var folder = fso.BuildPath(fso.GetParentFolderName(path), ".word-control-" + fso.GetTempName());
  var extension = fso.GetExtensionName(path);
  var suffix = extension ? "." + extension : "";
  var stage = { folder: folder, output: fso.BuildPath(folder, "new" + suffix),
    previous: fso.BuildPath(folder, "previous" + suffix), published: false };
  // Exclusive folder creation reserves our paths on the destination volume.
  fso.CreateFolder(folder);
  return stage;
}

function sameOutputFile(first, second) {
  if (canonicalPath(first) === canonicalPath(second)) return true;
  if (!first || !fso.FileExists(first) || !fso.FileExists(second)) return false;
  // Compare Windows short paths too, so an 8.3 spelling cannot rename the source.
  var firstShort = fso.GetFile(first).ShortPath, secondShort = fso.GetFile(second).ShortPath;
  if (!firstShort || !secondShort) throw new Error("cannot verify existing output file identity");
  return canonicalPath(firstShort) === canonicalPath(secondShort);
}

function verifyOutputFile(path, pdf) {
  if (!fso.FileExists(path) || Number(fso.GetFile(path).Size) <= 0) throw new Error("generated output is missing or empty: " + path);
  if (pdf) {
    var stream = fso.OpenTextFile(path, 1, false, 0);
    var header;
    try { header = stream.Read(5); } finally { stream.Close(); }
    if (header !== "%PDF-") throw new Error("generated output has no PDF header: " + path);
  }
}

function publishOutput(stage, path) {
  if (fso.FolderExists(path)) throw new Error("output path became a folder: " + path);
  var previous = fso.FileExists(path);
  if (previous) {
    if (!hasFlag("--overwrite")) throw new Error("refusing to overwrite an output created during generation: " + path);
    fso.MoveFile(path, stage.previous);
  }
  try {
    // MoveFile cannot overwrite a file that appeared after the preceding checks.
    fso.MoveFile(stage.output, path);
  } catch (publishError) {
    if (previous) {
      try { fso.MoveFile(stage.previous, path); }
      catch (restoreError) {
        throw new Error("publication failed: " + publishError.message + "; recovery failed: " + restoreError.message
          + "; previous output retained at: " + stage.previous);
      }
    }
    throw publishError;
  }
  stage.published = true;
}

function cleanupOutputStage(stage) {
  if (!stage) return "";
  try {
    if (fso.FileExists(stage.previous)) {
      if (!stage.published) return "previous output retained at: " + stage.previous;
      fso.DeleteFile(stage.previous, true);
    }
    if (fso.FileExists(stage.output)) fso.DeleteFile(stage.output, true);
    var folder = fso.GetFolder(stage.folder);
    if (folder.Files.Count || folder.SubFolders.Count) return "temporary files retained in: " + stage.folder;
    fso.DeleteFolder(stage.folder, false);
  } catch (e) { return "temporary/recovery files retained in: " + stage.folder + "; cleanup failed: " + e.message; }
  return "";
}

function preflightOutput() {
  var output = opt("--output", "");
  if (!output) return;
  var path = prepareOutputPath(output, "--output");
  if (!/\.(json|txt|tsv|log)$/i.test(path)) die("--output must be a scratch .json, .txt, .tsv, or .log file");
  var protectedOptions = ["--input", "--expect-path", "--path"];
  for (var i = 0; i < protectedOptions.length; i++) {
    var protectedPath = opt(protectedOptions[i], "");
    if (protectedPath && canonicalPath(path) === canonicalPath(protectedPath)) {
      die("--output must differ from " + protectedOptions[i]);
    }
  }
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
  stream.SaveToFile(path, hasFlag("--overwrite") ? 2 : 1);
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
    .replace(/\t/g, "\\t")
    .replace(/[\x00-\x1f]/g, function(ch) {
      var hex = ch.charCodeAt(0).toString(16);
      return "\\u" + ("0000" + hex).slice(-4);
    });
}

function q(value) {
  return "\"" + jescape(value) + "\"";
}

function stringArrayJson(values) {
  var out = [];
  for (var i = 0; i < values.length; i++) out.push(q(values[i]));
  return "[" + out.join(",") + "]";
}

function normalizeText(value) {
  if (value === null || typeof value === "undefined") return "";
  return String(value).replace(/\r/g, "\n");
}

function boolJson(value) {
  return value ? "true" : "false";
}

function textHash(value, preserveCharacters) {
  var text = preserveCharacters ? String(value) : normalizeText(value);
  var hash = 5381;
  for (var i = 0; i < text.length; i++) {
    hash = (((hash << 5) + hash) ^ text.charCodeAt(i)) >>> 0;
  }
  var hex = hash.toString(16);
  while (hex.length < 8) hex = "0" + hex;
  return hex;
}

function parseNonNegativeInt(value, name) {
  if (!/^\d+$/.test(String(value || ""))) die(name + " must be a non-negative integer");
  return parseInt(value, 10);
}

function parsePositiveInt(value, name) {
  if (!/^[1-9]\d*$/.test(String(value || ""))) die(name + " must be a positive integer");
  return parseInt(value, 10);
}

function parseHexColor(value) {
  value = String(value || "").replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(value)) die("color must be a 6-digit RGB hex value");
  var r = parseInt(value.substring(0, 2), 16);
  var g = parseInt(value.substring(2, 4), 16);
  var b = parseInt(value.substring(4, 6), 16);
  return r + (g * 256) + (b * 65536);
}

function parseBorderStyle(value) {
  var styles = {
    "none": 0,
    "single": 1,
    "dotted": 2,
    "dashed": 3,
    "dash-large": 4,
    "dash-dot": 5,
    "dash-dot-dot": 6,
    "double": 7,
    "triple": 8
  };
  value = String(value || "").toLowerCase();
  if (typeof styles[value] === "undefined") {
    die("border style must be none, single, dotted, dashed, dash-large, dash-dot, dash-dot-dot, double, or triple");
  }
  return styles[value];
}

function parseBorderWidth(value) {
  var widths = {
    "0.25": 2,
    "0.5": 4,
    "0.50": 4,
    "0.75": 6,
    "1": 8,
    "1.0": 8,
    "1.00": 8,
    "1.5": 12,
    "1.50": 12,
    "2.25": 18,
    "3": 24,
    "3.0": 24,
    "3.00": 24,
    "4.5": 36,
    "4.50": 36,
    "6": 48,
    "6.0": 48,
    "6.00": 48
  };
  value = String(value || "");
  if (typeof widths[value] === "undefined") {
    die("border width must be 0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, or 6 points");
  }
  return widths[value];
}

function setComBorder(border, color, lineWidth, lineStyle) {
  if (typeof lineStyle === "undefined") lineStyle = 1; // wdLineStyleSingle
  border.LineStyle = lineStyle;
  if (lineStyle !== 0) {
    border.LineWidth = lineWidth;
    border.Color = color;
  }
  if (Number(border.LineStyle) !== Number(lineStyle)) {
    throw new Error("border readback did not match requested style");
  }
  if (lineStyle !== 0 && (Number(border.LineWidth) !== Number(lineWidth) || Number(border.Color) !== Number(color))) {
    throw new Error("border readback did not match requested width or color");
  }
}

function setBorderCollection(borders, borderTypes, color, lineWidth, prefix, lineStyle) {
  var failures = [];
  for (var i = 0; i < borderTypes.length; i++) {
    try {
      setComBorder(borders(borderTypes[i]), color, lineWidth, lineStyle);
    } catch (e) {
      failures.push((prefix || "border") + ":" + borderTypes[i] + ":" + (e.message || String(e)));
    }
  }
  return failures;
}

function cleanCellText(value) {
  return normalizeText(value).replace(/\x07/g, "").replace(/\n+$/g, "");
}

function cellTextForSwap(cell) {
  var raw = String(cell.Range.Text);
  if (!/\r\x07$/.test(raw)) throw new Error("cannot swap cell text: cell terminator was not readable");
  // Remove only the cell terminator; paragraph and manual-break characters are content.
  var text = raw.substring(0, raw.length - 2);
  if (text.indexOf("\x07") >= 0) throw new Error("cannot swap cell text containing nested table cell markers");
  return text;
}

function setCellText(cell, text) {
  var range = cell.Range;
  range.End = range.End - 1;
  range.Text = text;
}

function getRegularTableDimensions(table, operation) {
  var rows;
  var cols;
  try { rows = Number(table.Rows.Count); } catch (e1) { die(operation + " requires a regular table without vertically merged cells: " + e1.message); }
  try { cols = Number(table.Columns.Count); } catch (e2) { die(operation + " requires a regular table without horizontally merged cells: " + e2.message); }
  return { rows: rows, cols: cols };
}

function resolveTableCell(table, cellOption, rowOption, colOption, label) {
  var cellText = opt(cellOption, "");
  var rowText = opt(rowOption, "");
  var colText = opt(colOption, "");
  if (cellText && (rowText || colText)) die(label + " must use either " + cellOption + " or " + rowOption + " with " + colOption);
  if (!cellText && (!rowText || !colText)) die(label + " requires " + cellOption + " or both " + rowOption + " and " + colOption);
  if (cellText) {
    var cellIndex = parsePositiveInt(cellText, cellOption);
    var linearCount = Number(table.Range.Cells.Count);
    if (cellIndex > linearCount) die(label + " cell index out of range; table has " + linearCount + " linear cells");
    return { cell: table.Range.Cells(cellIndex), cellIndex: cellIndex, row: 0, col: 0 };
  }
  var dimensions = getRegularTableDimensions(table, label);
  var row = parsePositiveInt(rowText, rowOption);
  var col = parsePositiveInt(colText, colOption);
  if (row > dimensions.rows) die(label + " row index out of range; table has " + dimensions.rows + " rows");
  if (col > dimensions.cols) die(label + " column index out of range; table has " + dimensions.cols + " columns");
  return { cell: table.Cell(row, col), cellIndex: 0, row: row, col: col };
}

function cellTargetJson(target) {
  return "{\"cell\":" + (target.cellIndex || "null") + ",\"row\":" + (target.row || "null") + ",\"col\":" + (target.col || "null") + "}";
}

function borderSignature(borders, borderTypes, errors, label) {
  var parts = [];
  for (var i = 0; i < borderTypes.length; i++) {
    try {
      var border = borders(borderTypes[i]);
      parts.push(borderTypes[i] + ":" + Number(border.LineStyle) + ":" + Number(border.LineWidth) + ":" + Number(border.Color));
    } catch (e) {
      parts.push(borderTypes[i] + ":?");
      errors.push(label + ":" + borderTypes[i] + ":" + (e.message || String(e)));
    }
  }
  return parts.join(",");
}

function parseBorderEdges(value, scope) {
  var tableMap = { "top": -1, "left": -2, "bottom": -3, "right": -4, "inside-h": -5, "inside-v": -6 };
  var cellMap = { "top": -1, "left": -2, "bottom": -3, "right": -4 };
  var map = scope === "table" ? tableMap : cellMap;
  var raw = String(value || "").toLowerCase().replace(/\s+/g, "").split(",");
  var expanded = [];
  for (var i = 0; i < raw.length; i++) {
    if (!raw[i]) continue;
    if (raw[i] === "all") {
      expanded = scope === "table" ? ["top", "left", "bottom", "right", "inside-h", "inside-v"] : ["top", "left", "bottom", "right"];
      break;
    }
    if (raw[i] === "outer") {
      expanded.push("top", "left", "bottom", "right");
    } else {
      expanded.push(raw[i]);
    }
  }
  if (!expanded.length) die("--edges must name at least one border edge");
  var names = [];
  var types = [];
  var seen = {};
  for (var j = 0; j < expanded.length; j++) {
    var name = expanded[j];
    if (typeof map[name] === "undefined") {
      die(scope + " border edge is invalid: " + name);
    }
    if (!seen[name]) {
      seen[name] = true;
      names.push(name);
      types.push(map[name]);
    }
  }
  return { names: names, types: types };
}

function applyBordersAtomically(borders, borderTypes, color, lineWidth, lineStyle) {
  var snapshots = [];
  var failures = [];
  var rollbackFailures = [];
  for (var i = 0; i < borderTypes.length; i++) {
    try {
      var border = borders(borderTypes[i]);
      snapshots.push({
        type: borderTypes[i],
        lineStyle: Number(border.LineStyle),
        lineWidth: Number(border.LineWidth),
        color: Number(border.Color)
      });
    } catch (e1) {
      failures.push("capture:" + borderTypes[i] + ":" + (e1.message || String(e1)));
      return { failures: failures, rollbackFailures: rollbackFailures, rolledBack: true, applied: false };
    }
  }
  for (var j = 0; j < borderTypes.length; j++) {
    try {
      setComBorder(borders(borderTypes[j]), color, lineWidth, lineStyle);
    } catch (e2) {
      failures.push("apply:" + borderTypes[j] + ":" + (e2.message || String(e2)));
      break;
    }
  }
  if (failures.length) {
    for (var k = 0; k < snapshots.length; k++) {
      try {
        var restore = borders(snapshots[k].type);
        restore.LineStyle = snapshots[k].lineStyle;
        if (snapshots[k].lineStyle !== 0) {
          restore.LineWidth = snapshots[k].lineWidth;
          restore.Color = snapshots[k].color;
        }
      } catch (e3) {
        rollbackFailures.push("restore:" + snapshots[k].type + ":" + (e3.message || String(e3)));
      }
    }
  }
  return {
    failures: failures,
    rollbackFailures: rollbackFailures,
    rolledBack: failures.length > 0 && rollbackFailures.length === 0,
    applied: failures.length ? null : true
  };
}

function parseTsv(text) {
  var rows = String(text || "").replace(/\r/g, "").split("\n");
  if (rows.length && rows[rows.length - 1] === "") rows.pop();
  var parsed = [];
  var maxCols = 0;
  for (var i = 0; i < rows.length; i++) {
    var cells = rows[i].split("\t");
    if (cells.length > maxCols) maxCols = cells.length;
    parsed.push(cells);
  }
  return { rows: parsed, rowCount: parsed.length, maxCols: maxCols };
}

function fillTableFromTsv(table, parsed) {
  var rowCount = Number(table.Rows.Count);
  var colCount = Number(table.Columns.Count);
  for (var r = 0; r < parsed.rows.length && r < rowCount; r++) {
    var cells = parsed.rows[r];
    for (var c = 0; c < cells.length && c < colCount; c++) {
      setCellText(table.Cell(r + 1, c + 1), cells[c]);
    }
  }
}

function tableFingerprint(table, readErrors) {
  var errors = [];
  var rows = "?";
  var cols = "?";
  var text = "";
  var formatting = [];
  // Irregular tables can legitimately lack a rectangular row/column model.
  try { rows = String(table.Rows.Count); } catch (e1) {}
  try { cols = String(table.Columns.Count); } catch (e2) {}
  // Display cleanup can hide trailing paragraphs and cell boundaries from a mutation guard.
  try { text = String(table.Range.Text); } catch (e3) { errors.push("table-text:" + e3.message); }
  try { formatting.push("tb:" + borderSignature(table.Borders, [-1, -2, -3, -4, -5, -6], errors, "table-borders")); }
  catch (e4) { errors.push("table-borders:" + e4.message); }
  try {
    var cells = table.Range.Cells;
    var cellCount = Number(cells.Count);
    for (var i = 1; i <= cellCount; i++) {
      var cell = cells(i);
      var shading = "?";
      var texture = "?";
      var cellShading = null;
      try { cellShading = cell.Shading; shading = String(Number(cellShading.BackgroundPatternColor)); } catch (e5) { errors.push("cell[" + i + "]-shading:" + e5.message); }
      try {
        if (cellShading === null) cellShading = cell.Shading;
        texture = String(Number(cellShading.Texture));
      } catch (e6) { errors.push("cell[" + i + "]-texture:" + e6.message); }
      formatting.push("c" + i + ":" + shading + ":" + texture + ":" + borderSignature(cell.Borders, [-1, -2, -3, -4], errors, "cell[" + i + "]-borders"));
    }
  } catch (e7) { errors.push("fingerprint-cells:" + e7.message); }
  if (errors.length) {
    if (!readErrors) throw new Error("cannot verify table fingerprint: " + errors.join("; "));
    for (var j = 0; j < errors.length; j++) readErrors.push(errors[j]);
    return null;
  }
  return textHash(rows + "x" + cols + "|" + text + "|" + formatting.join("|"), true);
}

// Post-check failures must not replace errors from a write or its rollback.
function postWriteTableFingerprint(table, errors) {
  try { return tableFingerprint(table, errors); }
  catch (e) { errors.push("fingerprint:" + (e.message || String(e))); return null; }
}

function parseSnapshotXml(xml) {
  if (xml.length > 8 * 1024 * 1024) throw new Error("XML comparison exceeds the 8 Mi-character snapshot limit");
  var dom = new ActiveXObject("Msxml2.DOMDocument.6.0");
  dom.async = false;
  dom.preserveWhiteSpace = true;
  dom.validateOnParse = false;
  dom.resolveExternals = false;
  dom.setProperty("ProhibitDTD", true);
  if (!dom.loadXML(xml)) throw new Error("cannot parse XML comparison: " + dom.parseError.reason);
  return dom;
}

function normalizeTableXml(xml) {
  var dom = parseSnapshotXml(xml);
  dom.setProperty("SelectionNamespaces", "xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'");
  if (dom.selectNodes("/pkg:package/pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData/w:document/w:body/w:tbl").length !== 1) {
    throw new Error("XML comparison requires exactly one top-level table in the snapshot");
  }
  // Only editing-session metadata is removed; styles, themes, revisions and content remain.
  var transform = parseSnapshotXml(
    '<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<xsl:output method="xml" omit-xml-declaration="yes"/>'
    + '<xsl:template match="@*|node()"><xsl:copy><xsl:apply-templates select="@*|node()"/></xsl:copy></xsl:template>'
    + '<xsl:template match="@w:rsidR|@w:rsidRDefault|@w:rsidRPr|@w:rsidP|@w:rsidDel|@w:rsidTr|@w:rsidSect|w:settings/w:rsids"/>'
    + '</xsl:stylesheet>');
  return String(dom.transformNode(transform));
}

function compareTableXml(table) {
  var started = new Date().getTime();
  var result = { candidate: null, stable: null, sourceChars: null, normalizedChars: null, elapsedMs: 0, errors: [] };
  try {
    var source = String(table.Range.WordOpenXML);
    result.sourceChars = source.length;
    var first = normalizeTableXml(source);
    result.normalizedChars = first.length;
    var second = normalizeTableXml(String(table.Range.WordOpenXML));
    result.stable = first === second;
    if (!result.stable) throw new Error("normalized XML changed between consecutive reads");
    result.candidate = "xml-v1:" + textHash(first);
  } catch (e) { result.errors.push(e.message || String(e)); }
  result.elapsedMs = new Date().getTime() - started;
  return result;
}

function xmlComparisonJson(result) {
  return "{\"version\":1,\"diagnostic_only\":true,\"candidate_hash\":" + (result.candidate === null ? "null" : q(result.candidate))
    + ",\"stable_repeat\":" + (result.stable === null ? "null" : boolJson(result.stable))
    + ",\"source_chars\":" + result.sourceChars + ",\"normalized_chars\":" + result.normalizedChars
    + ",\"elapsed_ms\":" + result.elapsedMs + ",\"errors\":" + stringArrayJson(result.errors) + "}";
}

function equationFingerprint(equation, readErrors) {
  try { return textHash(normalizeText(equation.Range.Text)); }
  catch (e) {
    if (!readErrors) throw new Error("cannot verify equation fingerprint: " + e.message);
    readErrors.push("equation-text:" + e.message);
    return null;
  }
}

function requireFingerprint(actual, optionName, unsafeFlag) {
  var expected = opt(optionName, "");
  if (/^xml-v1:/i.test(expected)) die("experimental XML comparison hashes cannot authorize mutations");
  if (actual === null || typeof actual === "undefined") die("target inspection is incomplete");
  if (!expected && !hasFlag(unsafeFlag)) {
    die("target mutation requires " + optionName + " from a fresh inspection; use " + unsafeFlag + " only after manual verification");
  }
  if (expected && String(expected).toLowerCase() !== String(actual).toLowerCase()) {
    die("target fingerprint changed; inspect the document again before mutating");
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
  if (format === "latex") {
    var supported = /^(frac|sqrt|left|right|int|sum|prod|cdot|times|le|leq|ge|geq|neq|approx|pm|alpha|beta|gamma|delta|epsilon|theta|lambda|mu|pi|rho|sigma|tau|phi|omega|Delta|Omega)$/;
    var commands = text.match(/\\[A-Za-z]+/g) || [];
    for (var i = 0; i < commands.length; i++) {
      if (!supported.test(commands[i].substring(1))) die("unsupported LaTeX command: " + commands[i]);
    }
    var depth = 0;
    for (var j = 0; j < text.length; j++) {
      if (text.charAt(j) === "{") depth++;
      if (text.charAt(j) === "}" && --depth < 0) die("unbalanced LaTeX braces");
    }
    if (depth !== 0) die("unbalanced LaTeX braces");
    var converted = convertLatexToWordLinear(text);
    var unknown = converted.match(/\\[A-Za-z]+/g);
    if (unknown && unknown.length) {
      die("unsupported LaTeX command(s): " + unknown.join(", ") + "; use --format linear only after verifying Word linear syntax");
    }
    return converted;
  }
  if (format === "linear" || format === "word") return text.replace(/\r/g, "").replace(/\n/g, " ");
  die("--format must be latex, linear, or word");
}

function buildEquationInRange(doc, range, linearText) {
  var start = range.Start;
  range.Text = linearText;
  var mathRange = range.Duplicate;
  mathRange.SetRange(start, start + linearText.length);
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

function requireExpectedDocument(doc) {
  var expectedPath = opt("--expect-path", "");
  var expectedName = opt("--expect-name", "");
  if (!expectedPath && !expectedName && !hasFlag("--allow-active")) {
    die("mutation requires --expect-path or --expect-name; use --allow-active only after manually verifying the active document");
  }
  if (expectedPath) {
    var actualPath = safeDocPath(doc);
    if (!actualPath || canonicalPath(actualPath) !== canonicalPath(expectedPath)) {
      die("active document path mismatch; expected " + absPath(expectedPath) + ", got " + (actualPath || "[unsaved document]"));
    }
  }
  if (expectedName && !expectedPath && String(doc.Path)) die("saved documents require --expect-path, not --expect-name");
  if (expectedName && String(doc.Name).toLowerCase() !== String(expectedName).toLowerCase()) {
    die("active document name mismatch; expected " + expectedName + ", got " + String(doc.Name));
  }
}

function getMutationDocument(word) {
  var doc = getActiveDocument(word);
  requireExpectedDocument(doc);
  return doc;
}

function selectionSnapshot(word) {
  var range = word.Selection.Range;
  var text = normalizeText(range.Text);
  return {
    range: range,
    storyType: Number(range.StoryType),
    start: Number(range.Start),
    end: Number(range.End),
    collapsed: Number(range.Start) === Number(range.End),
    text: text,
    hash: textHash(text)
  };
}

function selectionSnapshotJson(snapshot) {
  return "{"
    + "\"story_type\":" + snapshot.storyType + ","
    + "\"start\":" + snapshot.start + ","
    + "\"end\":" + snapshot.end + ","
    + "\"collapsed\":" + boolJson(snapshot.collapsed) + ","
    + "\"text_length\":" + snapshot.text.length + ","
    + "\"text_hash\":" + q(snapshot.hash)
    + "}";
}

function requireExpectedSelection(word, allowCollapsed) {
  var snapshot = selectionSnapshot(word);
  if (snapshot.collapsed && !allowCollapsed && !hasFlag("--allow-insert")) {
    die("selection is collapsed; pass --allow-insert only when insertion at the cursor is intended");
  }
  var expectedStory = opt("--expect-story-type", "");
  var expectedStart = opt("--expect-start", "");
  var expectedEnd = opt("--expect-end", "");
  var expectedHash = opt("--expect-selection-hash", "");
  if ((!expectedStory || !expectedStart || !expectedEnd || !expectedHash) && !hasFlag("--allow-unverified-selection")) {
    die("selection mutation requires --expect-story-type, --expect-start, --expect-end, and --expect-selection-hash from selection-info");
  }
  if (expectedStory && parsePositiveInt(expectedStory, "--expect-story-type") !== snapshot.storyType) die("selection story changed; rerun selection-info");
  if (expectedStart && parseNonNegativeInt(expectedStart, "--expect-start") !== snapshot.start) {
    die("selection start changed; rerun selection-info before mutating");
  }
  if (expectedEnd && parseNonNegativeInt(expectedEnd, "--expect-end") !== snapshot.end) {
    die("selection end changed; rerun selection-info before mutating");
  }
  if (expectedHash && String(expectedHash).toLowerCase() !== snapshot.hash) {
    die("selection text changed; rerun selection-info before mutating");
  }
  return snapshot;
}

function activeDocJson(word, doc) {
  var path = safeDocPath(doc);
  var saved = "";
  var track = "";
  var readOnly = "";
  var protectionType = "";
  try { saved = String(doc.Saved); } catch (e1) {}
  try { track = String(doc.TrackRevisions); } catch (e2) {}
  try { readOnly = String(doc.ReadOnly); } catch (e3) {}
  try { protectionType = String(doc.ProtectionType); } catch (e4) {}
  return "{"
    + "\"name\":" + q(doc.Name) + ","
    + "\"path\":" + q(path) + ","
    + "\"saved\":" + q(saved) + ","
    + "\"track_revisions\":" + q(track) + ","
    + "\"read_only\":" + q(readOnly) + ","
    + "\"protection_type\":" + q(protectionType)
    + "}";
}

function commandHelp() {
  emit([
    "word-control commands:",
    "  status [--output file]",
    "  selection [--output file]",
    "  selection-info [--output file]",
    "  document-text [--max chars|--full] [--output file]",
    "  find-text --input file [--story main|footnotes|endnotes] [--from position] [--max count] [--context chars] [--output file]",
    "  paragraphs [--from index] [--max count] [--output file]",
    "  tables [--table index|--max count] [--detail full|text|summary] [--output file]",
    "  tables --table index --detail text [--cell-from index] [--cell-max count] [--output file]",
    "  tables --table index --compare-xml [--expect-path file] [--output file] (diagnostic only)",
    "  equations [--index index] [--output file]",
    "  convert-equation --input file [--format latex|linear|word] [--output file]",
    "  enable-track-changes --expect-path file --yes",
    "  disable-track-changes --expect-path file --yes",
    "  replace-selection --input file --expect-path file --expect-story-type n --expect-start n --expect-end n --expect-selection-hash hash [--track] [--allow-insert] --yes",
    "  insert-comment --input file --expect-path file --expect-story-type n --expect-start n --expect-end n --expect-selection-hash hash --yes",
    "  create-table --rows n --cols n [--input table.tsv] [--at selection|end] --expect-path file [selection guards when --at selection] --yes",
    "  set-cell --table n --expect-table-fingerprint hash (--cell n | --row n --col n) --input file --expect-path file --yes",
    "  swap-cell-text --table n --expect-table-fingerprint hash (--from-cell n | --from-row n --from-col n) (--to-cell n | --to-row n --to-col n) --expect-path file --yes",
    "  insert-row --table n --expect-table-fingerprint hash (--before n | --at-end) --expect-path file --yes",
    "  delete-row --table n --expect-table-fingerprint hash --row n --expect-path file --yes",
    "  insert-column --table n --expect-table-fingerprint hash (--before n | --at-end) --expect-path file --yes",
    "  delete-column --table n --expect-table-fingerprint hash --col n --expect-path file --yes",
    "  set-cell-shading --table n --expect-table-fingerprint hash (--cell n | --row n --col n) --color RRGGBB --expect-path file --yes",
    "  set-cell-borders --table n --expect-table-fingerprint hash [cell target] --edges top,bottom --style single [--color RRGGBB] [--width 0.5] --expect-path file --yes",
    "  set-table-borders --table n --expect-table-fingerprint hash --edges outer,inside-h,inside-v --style single [--color RRGGBB] [--width 0.5] --expect-path file --yes",
    "  delete-table --table n --expect-table-fingerprint hash --expect-path file --yes",
    "  normalize-table-borders --table n --expect-table-fingerprint hash [--color D9DEE8] [--line-width 4] --expect-path file --yes",
    "  insert-equation --input file [--format latex|linear|word] [--at selection|end] --expect-path file [selection guards when --at selection] --yes",
    "  set-equation --index n --expect-equation-fingerprint hash --input file [--format latex|linear|word] --expect-path file --yes",
    "  delete-equation --index n --expect-equation-fingerprint hash --expect-path file --yes",
    "  save-active --expect-path file --yes",
    "  close-active --save|--discard --expect-path file --yes",
    "  save-copy --path file --expect-path file [--overwrite] --yes",
    "  export-pdf --path preview.pdf --expect-path file [--overwrite] --yes",
    "  open --path file",
    "  smoke --path new-file.docx [--overwrite] --yes",
    "",
    "Mutation guards: use --expect-name for unsaved documents; --allow-active, --allow-unverified-selection, and --allow-unverified-target are explicit unsafe overrides.",
    "replace-paragraph is disabled; select the exact paragraph and use replace-selection."
  ].join("\n"));
}

function commandStatus() {
  var word = getWord();
  var docs = Number(word.Documents.Count);
  var selectionText = "";
  try { selectionText = String(word.Selection.Text); } catch (e1) {}
  var docPart = "null";
  var selectionPart = "null";
  if (docs > 0) docPart = activeDocJson(word, word.ActiveDocument);
  if (docs > 0) {
    try { selectionPart = selectionSnapshotJson(selectionSnapshot(word)); } catch (e2) {}
  }
  emit("{"
    + "\"ok\":true,"
    + "\"word_version\":" + q(word.Version) + ","
    + "\"documents_count\":" + docs + ","
    + "\"active_document\":" + docPart + ","
    + "\"selection_text_length\":" + normalizeText(selectionText).length + ","
    + "\"selection\":" + selectionPart
    + "}");
}

function commandSelection() {
  var word = getWord();
  getActiveDocument(word);
  var text = "";
  try { text = normalizeText(word.Selection.Text); } catch (e) { die("cannot read current selection: " + e.message); }
  emit(text);
}

function commandSelectionInfo() {
  var word = getWord();
  var doc = getActiveDocument(word);
  var snapshot = selectionSnapshot(word);
  emit("{\"ok\":true,\"document\":" + activeDocJson(word, doc) + ",\"selection\":" + selectionSnapshotJson(snapshot) + "}");
}

function commandConvertEquation() {
  var linearText = equationInputText();
  if (!linearText) die("equation input is empty after conversion");
  emit("{\"ok\":true,\"linear\":" + q(linearText) + "}");
}

function commandDocumentText() {
  var word = getWord();
  var doc = getActiveDocument(word);
  var text = normalizeText(doc.Content.Text);
  if (!hasFlag("--full")) {
    var max = parsePositiveInt(opt("--max", "20000"), "--max");
    if (text.length > max) text = text.substring(0, max) + "\n[truncated at " + max + " chars]";
  }
  emit(text);
}

function commandParagraphs() {
  var word = getWord();
  var doc = getActiveDocument(word);
  var from = parsePositiveInt(opt("--from", "1"), "--from");
  if (!isFinite(from)) die("--from must be a finite positive integer");
  var max = parsePositiveInt(opt("--max", "80"), "--max");
  var paragraphs = doc.Paragraphs;
  var count = Number(paragraphs.Count);
  max = Math.min(max, Math.max(0, count - from + 1));
  var parts = [];
  var readErrors = [];
  for (var i = from; i < from + max; i++) {
    var textJson = "null";
    try { textJson = q(normalizeText(paragraphs(i).Range.Text)); }
    catch (e) { readErrors.push("paragraph[" + i + "]:" + (e.message || String(e))); }
    parts.push("{\"index\":" + i + ",\"text\":" + textJson + "}");
  }
  var page = hasFlag("--from") ? ",\"from\":" + from + ",\"next_from\":" + (max > 0 && from + max <= count ? from + max : "null") : "";
  var incomplete = readErrors.length ? ",\"inspection_complete\":false,\"read_errors\":" + stringArrayJson(readErrors) : "";
  emit("{\"ok\":true,\"paragraph_count\":" + count + ",\"returned\":" + max + page + incomplete + ",\"paragraphs\":[" + parts.join(",") + "]}");
}

function commandFindText() {
  var query = readUtf8(opt("--input", ""));
  if (!query.length || query.length > 200 || /[\x00-\x1f]/.test(query)) die("find-text requires 1-200 literal characters without control characters or a trailing newline");
  var story = opt("--story", "main");
  var storyType = story === "main" ? 1 : (story === "footnotes" ? 2 : (story === "endnotes" ? 3 : 0));
  if (!storyType) die("--story must be main, footnotes or endnotes");
  var from = parseNonNegativeInt(opt("--from", "0"), "--from");
  var max = parsePositiveInt(opt("--max", "20"), "--max");
  var contextChars = parseNonNegativeInt(opt("--context", "80"), "--context");
  if (!isFinite(from) || from > 2147483647 || max > 100 || contextChars > 500) die("find-text limits: --from <= 2147483647, --max <= 100, --context <= 500");
  var word = getWord();
  var doc = getActiveDocument(word);
  if (hasFlag("--expect-path") || hasFlag("--expect-name")) requireExpectedDocument(doc);
  var documentJson = "{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var available = storyType === 1 || Number(storyType === 2 ? doc.Footnotes.Count : doc.Endnotes.Count) > 0;
  var matches = [], errors = [], more = false, next = null, end = 0;
  if (available) {
    var base = (storyType === 1 ? doc.Content : doc.StoryRanges(storyType)).Duplicate;
    var start = Number(base.Start);
    end = Number(base.End);
    var cursor = Math.max(start, from);
    if (cursor < end) {
      var range = base.Duplicate;
      var finder = range.Find;
      var names = ["Text", "MatchCase", "MatchWholeWord", "MatchWildcards", "MatchSoundsLike", "MatchAllWordForms", "Forward", "Wrap", "Format", "MatchByte", "MatchFuzzy", "MatchPrefix", "MatchSuffix", "IgnoreSpace", "IgnorePunct"];
      var previous = [], changed = 0;
      try {
        // Read all settings before changing any; do not clear the user's formatting criteria.
        for (var p = 0; p < names.length; p++) previous.push(finder[names[p]]);
        for (var s = 0; s < names.length; s++) {
          changed = s + 1;
          finder[names[s]] = names[s] === "Text" ? query.replace(/\^/g, "^^") : (names[s] === "MatchCase" || names[s] === "MatchByte" || names[s] === "Forward" ? true : (names[s] === "Wrap" ? 0 : false));
        }
        while (cursor < end) {
          range.SetRange(cursor, end);
          if (!finder.Execute()) break;
          var hitStart = Number(range.Start), hitEnd = Number(range.End);
          if (hitStart < cursor || hitEnd <= hitStart || hitEnd > end || String(range.Text) !== query) throw new Error("find returned a nonliteral or out-of-scope match");
          if (matches.length === max) { more = true; break; }
          var excerpt = base.Duplicate;
          excerpt.SetRange(Math.max(start, hitStart - contextChars), Math.min(end, hitEnd + contextChars));
          var excerptText = normalizeText(excerpt.Text);
          var textLimit = query.length + contextChars * 2;
          matches.push("{\"start\":" + hitStart + ",\"end\":" + hitEnd + ",\"text\":" + q(query)
            + ",\"context\":" + q(excerptText.substring(0, textLimit)) + ",\"context_truncated\":" + boolJson(excerptText.length > textLimit) + "}");
          cursor = hitEnd;
          next = cursor;
        }
      } catch (e) { errors.push("find:" + (e.message || String(e))); }
      finally {
        for (var restore = changed - 1; restore >= 0; restore--) {
          try { finder[names[restore]] = previous[restore]; }
          catch (restoreError) { errors.push("restore-find-" + names[restore] + ":" + (restoreError.message || String(restoreError))); }
        }
      }
    }
  }
  var payload = "{\"ok\":" + boolJson(errors.length === 0) + ",\"document\":" + documentJson
    + ",\"story\":" + q(story) + ",\"story_type\":" + storyType + ",\"story_available\":" + boolJson(available)
    + ",\"from\":" + from + ",\"story_end\":" + end + ",\"returned\":" + matches.length
    + ",\"has_more\":" + (errors.length ? "null" : boolJson(more)) + ",\"next_from\":" + (more && !errors.length ? next : "null")
    + ",\"inspection_complete\":" + boolJson(errors.length === 0) + ",\"read_errors\":" + stringArrayJson(errors)
    + ",\"matches\":[" + matches.join(",") + "]}";
  if (errors.length) failJson(payload, 3);
  emit(payload);
}

function commandTables() {
  var tableText = opt("--table", "");
  var maxText = opt("--max", "");
  if (tableText && maxText) die("use either --table or --max, not both");
  var detail = opt("--detail", "full");
  if (detail !== "full" && detail !== "text" && detail !== "summary") die("--detail must be full, text or summary");
  var compareXml = hasFlag("--compare-xml");
  if (compareXml && (!tableText || detail !== "full")) die("XML comparison requires a single table with full detail");
  var cellPage = hasFlag("--cell-from") || hasFlag("--cell-max");
  if (cellPage && (!tableText || detail !== "text")) die("cell pagination requires --table and --detail text");
  var cellFrom = cellPage ? parsePositiveInt(opt("--cell-from", "1"), "--cell-from") : 1;
  var cellMax = cellPage ? parsePositiveInt(opt("--cell-max", "40"), "--cell-max") : 0;
  if (cellPage && (!isFinite(cellFrom) || cellMax > 200)) die("cell pagination requires a finite index and --cell-max <= 200");
  var first = tableText ? parsePositiveInt(tableText, "--table") : 1;
  var word = getWord();
  var doc = getActiveDocument(word);
  if ((cellPage || compareXml) && (hasFlag("--expect-path") || hasFlag("--expect-name"))) requireExpectedDocument(doc);
  var tables = doc.Tables;
  var count = Number(tables.Count);
  if (tableText && first > count) die("table index out of range; document has " + count + " tables");
  var returned = tableText ? 1 : (maxText ? parsePositiveInt(maxText, "--max") : count);
  if (returned > count) returned = count;
  var parts = [];
  var comparisonFailed = false;
  for (var i = first; i < first + returned; i++) {
    var table = tables(i);
    var rowCount = 0;
    var colCount = 0;
    var rowCountKnown = true;
    var colCountKnown = true;
    var layoutWarnings = [];
    var readErrors = [];
    try { rowCount = Number(table.Rows.Count); } catch (e1) { rowCountKnown = false; layoutWarnings.push("rows:" + (e1.message || String(e1))); }
    try { colCount = Number(table.Columns.Count); } catch (e2) { colCountKnown = false; layoutWarnings.push("columns:" + (e2.message || String(e2))); }
    var tableInfo = "{\"index\":" + i + ",\"rows\":" + (rowCountKnown ? rowCount : "null") + ",\"cols\":" + (colCountKnown ? colCount : "null");
    if (detail === "summary") {
      var summaryCellCount = "null";
      try { summaryCellCount = String(Number(table.Range.Cells.Count)); }
      catch (summaryError) { readErrors.push("cell-count:" + (summaryError.message || String(summaryError))); }
      // Counts alone cannot establish a rectangular layout or a safe write target.
      parts.push(tableInfo + ",\"layout\":\"unverified\",\"cell_count\":" + summaryCellCount
        + ",\"fingerprint\":null,\"inspection_complete\":false"
        + ",\"layout_warnings\":" + stringArrayJson(layoutWarnings)
        + ",\"read_errors\":" + stringArrayJson(readErrors)
        + ",\"cells\":[],\"linear_cells\":[]}");
      continue;
    }
    var cells = [];
    var linearCells = [];
    var cellCount = 0;
    var rectangular = !cellPage && rowCountKnown && colCountKnown && rowCount > 0 && colCount > 0;
    var rectangleFailed = false;
    if (rectangular) {
      for (var r = 1; r <= rowCount && !rectangleFailed; r++) {
        var rowCells = [];
        for (var c = 1; c <= colCount; c++) {
          try {
            rowCells.push(q(cleanCellText(table.Cell(r, c).Range.Text)));
          } catch (e3) {
            layoutWarnings.push("cell[" + r + "," + c + "]:" + (e3.message || String(e3)));
            rectangleFailed = true;
            break;
          }
        }
        if (!rectangleFailed) cells.push("[" + rowCells.join(",") + "]");
      }
      if (rectangleFailed) {
        rectangular = false;
        cells = [];
      }
    }
    if (rectangular) {
      cellCount = rowCount * colCount;
    } else {
      try {
        var nativeCells = table.Range.Cells;
        cellCount = Number(nativeCells.Count);
        var cellEnd = cellPage ? Math.min(cellCount, cellFrom + cellMax - 1) : cellCount;
        for (var k = cellFrom; k <= cellEnd; k++) {
          var cell = nativeCells(k);
          var rowJson = "null";
          var colJson = "null";
          var textJson = "null";
          try { rowJson = String(Number(cell.RowIndex)); } catch (e4) { layoutWarnings.push("linear-cell[" + k + "]-row:" + (e4.message || String(e4))); }
          try { colJson = String(Number(cell.ColumnIndex)); } catch (e5) { layoutWarnings.push("linear-cell[" + k + "]-column:" + (e5.message || String(e5))); }
          try { textJson = q(cleanCellText(cell.Range.Text)); } catch (e6) { readErrors.push("linear-cell[" + k + "]-text:" + (e6.message || String(e6))); }
          linearCells.push("{\"index\":" + k + ",\"row\":" + rowJson + ",\"col\":" + colJson + ",\"text\":" + textJson + "}");
        }
      } catch (e7) {
        readErrors.push("linear-cell-enumeration:" + (e7.message || String(e7)));
      }
    }
    // Text-only inspection never supplies a guard; mutations still need full formatting reads.
    var fingerprint = detail === "full" ? tableFingerprint(table, readErrors) : null;
    var comparison = null;
    if (compareXml) {
      comparison = readErrors.length ? { candidate: null, stable: null, sourceChars: null, normalizedChars: null,
        elapsedMs: 0, errors: ["legacy table inspection is incomplete; XML comparison skipped"] } : compareTableXml(table);
      if (comparison.errors.length) {
        comparisonFailed = true;
        fingerprint = null;
        for (var e = 0; e < comparison.errors.length; e++) readErrors.push("xml-comparison:" + comparison.errors[e]);
      }
    }
    parts.push(tableInfo
      + ",\"layout\":" + q(cellPage ? "unverified" : (rectangular ? "rectangular" : "irregular"))
      + ",\"cell_count\":" + cellCount
      + ",\"fingerprint\":" + (fingerprint === null ? "null" : q(fingerprint))
      + ",\"inspection_complete\":" + boolJson(detail === "full" && readErrors.length === 0)
      + ",\"layout_warnings\":" + stringArrayJson(layoutWarnings)
      + ",\"read_errors\":" + stringArrayJson(readErrors)
      + ",\"cells\":[" + cells.join(",") + "]"
      + ",\"linear_cells\":[" + linearCells.join(",") + "]"
      + (compareXml ? ",\"xml_comparison\":" + xmlComparisonJson(comparison) : "")
      + (cellPage ? ",\"cell_from\":" + cellFrom + ",\"returned_cells\":" + linearCells.length
        + ",\"next_cell\":" + (!readErrors.length && linearCells.length && cellFrom + linearCells.length <= cellCount ? cellFrom + linearCells.length : "null") : "") + "}");
  }
  var detailPart = hasFlag("--detail") ? ",\"detail\":" + q(detail) : "";
  var payload = "{\"ok\":" + boolJson(!comparisonFailed) + ",\"table_count\":" + count + ",\"returned\":" + returned + detailPart
    + (cellPage || compareXml ? ",\"document\":{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}" : "")
    + ",\"tables\":[" + parts.join(",") + "]}";
  if (comparisonFailed) failJson(payload, 3);
  emit(payload);
}

function commandEquations() {
  var indexText = opt("--index", "");
  var first = indexText ? parsePositiveInt(indexText, "--index") : 1;
  var word = getWord();
  var doc = getActiveDocument(word);
  var equations = doc.OMaths;
  var count = Number(equations.Count);
  if (indexText && first > count) die("equation index out of range; document has " + count + " equations");
  var end = indexText ? first : count;
  var parts = [];
  for (var i = first; i <= end; i++) {
    var text = null;
    var fingerprint = null;
    var readErrors = [];
    try {
      text = normalizeText(equations(i).Range.Text);
      fingerprint = textHash(text);
    } catch (e) { readErrors.push("equation[" + i + "]:" + (e.message || String(e))); }
    var incomplete = readErrors.length ? ",\"inspection_complete\":false,\"read_errors\":" + stringArrayJson(readErrors) : "";
    parts.push("{\"index\":" + i + ",\"text\":" + (text === null ? "null" : q(text))
      + ",\"fingerprint\":" + (fingerprint === null ? "null" : q(fingerprint)) + incomplete + "}");
  }
  emit("{\"ok\":true,\"equation_count\":" + count + (indexText ? ",\"returned\":1" : "") + ",\"equations\":[" + parts.join(",") + "]}");
}

function commandTrack(on) {
  requireYes();
  var word = getWord();
  var doc = getMutationDocument(word);
  doc.TrackRevisions = on;
  emit("{\"ok\":true,\"track_revisions\":" + q(String(doc.TrackRevisions)) + "}");
}

function commandReplaceSelection() {
  requireYes();
  var input = opt("--input", "");
  var text = readUtf8(input);
  var word = getWord();
  var doc = getMutationDocument(word);
  var selection = requireExpectedSelection(word, false);
  var trackBefore = Boolean(doc.TrackRevisions);
  try {
    if (hasFlag("--track")) doc.TrackRevisions = true;
    selection.range.Text = text;
  } catch (e) {
    var restoreError = "";
    try { doc.TrackRevisions = trackBefore; } catch (restore) { restoreError = "; track state restoration failed: " + restore.message; }
    die("failed to replace selection: " + e.message + restoreError);
  }
  emit("{\"ok\":true,\"action\":\"replace-selection\",\"chars\":" + text.length
    + ",\"selection_before\":" + selectionSnapshotJson(selection)
    + ",\"track_revisions\":" + q(String(doc.TrackRevisions)) + "}");
}

function commandReplaceParagraph() {
  die("replace-paragraph is disabled because Word paragraph ranges can duplicate or shift content; select the exact paragraph and use replace-selection");
}

function commandInsertComment() {
  requireYes();
  var input = opt("--input", "");
  var text = readUtf8(input);
  var word = getWord();
  var doc = getMutationDocument(word);
  var selection = requireExpectedSelection(word, false);
  try {
    doc.Comments.Add(selection.range, text);
  } catch (e) {
    die("failed to insert comment: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"insert-comment\",\"chars\":" + text.length
    + ",\"selection_before\":" + selectionSnapshotJson(selection) + "}");
}

function commandCreateTable() {
  requireYes();
  var rows = parsePositiveInt(opt("--rows", "0"), "--rows");
  var cols = parsePositiveInt(opt("--cols", "0"), "--cols");
  var at = opt("--at", "selection");
  if (at !== "selection" && at !== "end") die("--at must be selection or end");
  var input = opt("--input", "");
  var parsed = null;
  var truncated = false;
  if (input) {
    parsed = parseTsv(readUtf8(input));
    truncated = parsed.rowCount > rows || parsed.maxCols > cols;
    if (truncated && !hasFlag("--allow-truncate")) {
      die("TSV dimensions " + parsed.rowCount + "x" + parsed.maxCols + " exceed target table " + rows + "x" + cols + "; resize the table or pass --allow-truncate explicitly");
    }
  }
  var word = getWord();
  var doc = getMutationDocument(word);
  var range;
  if (at === "end") {
    range = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
  } else {
    range = requireExpectedSelection(word, true).range;
  }
  var documentJson = "{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var table = null, tableCount = null, applied = null, fillComplete = parsed ? false : null;
  var errors = [], fingerprint = null;
  try {
    table = doc.Tables.Add(range, rows, cols);
    applied = true;
  } catch (e) {
    errors.push("create:" + (e.message || String(e)));
  }
  if (table) {
    if (parsed) {
      try { fillTableFromTsv(table, parsed); fillComplete = true; }
      catch (fillError) { errors.push("fill:" + (fillError.message || String(fillError))); }
    }
    if (!parsed || fillComplete) {
      try {
        if (Number(table.Rows.Count) !== rows || Number(table.Columns.Count) !== cols) {
          throw new Error("table dimensions differ from requested dimensions");
        }
        if (parsed) {
          for (var r = 0; r < parsed.rows.length && r < rows; r++) {
            for (var c = 0; c < parsed.rows[r].length && c < cols; c++) {
              var raw = String(table.Cell(r + 1, c + 1).Range.Text);
              if (!/\r\x07$/.test(raw) || normalizeText(raw.substring(0, raw.length - 2)) !== normalizeText(parsed.rows[r][c])) {
                throw new Error("cell[" + (r + 1) + "," + (c + 1) + "] differs from requested text");
              }
            }
          }
        }
      } catch (readError) { errors.push("readback:" + (readError.message || String(readError))); }
    }
    fingerprint = postWriteTableFingerprint(table, errors);
  }
  try { tableCount = Number(doc.Tables.Count); }
  catch (countError) { errors.push("table-count:" + (countError.message || String(countError))); }
  var verified = applied === true && errors.length === 0 && fingerprint !== null;
  var payload = "{\"ok\":" + boolJson(verified) + ",\"action\":\"create-table\",\"table_count\":" + tableCount
    + ",\"rows\":" + rows + ",\"cols\":" + cols + ",\"truncated\":" + boolJson(truncated)
    + ",\"document\":" + documentJson + ",\"applied\":" + (applied === null ? "null" : boolJson(applied))
    + ",\"fill_complete\":" + (fillComplete === null ? "null" : boolJson(fillComplete))
    + ",\"verified\":" + boolJson(verified) + ",\"inspection_complete\":" + boolJson(verified)
    + ",\"errors\":" + stringArrayJson(errors) + ",\"fingerprint\":" + (verified ? q(fingerprint) : "null") + "}";
  if (!verified) failJson(payload, 3);
  emit(payload);
}

function commandSetCell() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var cellText = opt("--cell", "");
  var rowText = opt("--row", "");
  var colText = opt("--col", "");
  if (cellText && (rowText || colText)) die("use either --cell or --row with --col, not both");
  if (!cellText && (!rowText || !colText)) die("set-cell requires --cell or both --row and --col");
  var cellIndex = cellText ? parsePositiveInt(cellText, "--cell") : 0;
  var row = rowText ? parsePositiveInt(rowText, "--row") : 0;
  var col = colText ? parsePositiveInt(colText, "--col") : 0;
  var input = opt("--input", "");
  var text = readUtf8(input);
  var word = getWord();
  var doc = getMutationDocument(word);
  var tableCount = Number(doc.Tables.Count);
  if (tableIndex > tableCount) die("table index out of range; document has " + tableCount + " tables");
  var table = doc.Tables(tableIndex);
  requireFingerprint(tableFingerprint(table), "--expect-table-fingerprint", "--allow-unverified-target");
  var targetCell;
  if (cellIndex) {
    var linearCount = Number(table.Range.Cells.Count);
    if (cellIndex > linearCount) die("cell index out of range; table has " + linearCount + " linear cells");
    targetCell = table.Range.Cells(cellIndex);
  } else {
    var rowCount;
    var colCount;
    try { rowCount = Number(table.Rows.Count); } catch (e1) { die("table does not expose regular rows; inspect linear_cells and use --cell: " + e1.message); }
    try { colCount = Number(table.Columns.Count); } catch (e2) { die("table does not expose regular columns; inspect linear_cells and use --cell: " + e2.message); }
    if (row > rowCount) die("row index out of range; table has " + rowCount + " rows");
    if (col > colCount) die("column index out of range; table has " + colCount + " columns");
    targetCell = table.Cell(row, col);
  }
  var documentJson = "{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var applied = false, verified = false, fingerprint = null, readback = "null", errors = [];
  try { setCellText(targetCell, text); applied = true; }
  catch (e3) { applied = null; errors.push("write:" + (e3.message || String(e3))); }
  if (applied === true) {
    try {
      var raw = String(targetCell.Range.Text);
      if (!/\r\x07$/.test(raw)) throw new Error("cell terminator was not readable");
      var actual = normalizeText(raw.substring(0, raw.length - 2));
      var matches = actual === normalizeText(text.replace(/\r\n/g, "\n"));
      readback = "{\"matches_requested\":" + boolJson(matches) + ",\"chars\":" + actual.length + ",\"text_hash\":" + q(textHash(actual)) + "}";
      if (!matches) errors.push("readback:cell text differs from requested text");
    } catch (readError) { errors.push("readback:" + (readError.message || String(readError))); }
    try { fingerprint = tableFingerprint(table, errors); }
    catch (fingerprintError) { errors.push("fingerprint:" + (fingerprintError.message || String(fingerprintError))); }
    verified = errors.length === 0 && fingerprint !== null;
  }
  // An interrupted write may already have changed the document. Never return a retry guard.
  var payload = "{\"ok\":" + boolJson(verified) + ",\"action\":\"set-cell\",\"table\":" + tableIndex
    + ",\"cell\":" + (cellIndex || "null") + ",\"row\":" + (row || "null") + ",\"col\":" + (col || "null")
    + ",\"chars\":" + text.length + ",\"document\":" + documentJson + ",\"applied\":" + (applied === null ? "null" : boolJson(applied))
    + ",\"verified\":" + boolJson(verified) + ",\"inspection_complete\":" + boolJson(verified) + ",\"readback\":" + readback
    + ",\"errors\":" + stringArrayJson(errors) + ",\"fingerprint\":" + (verified ? q(fingerprint) : "null") + "}";
  if (!verified) failJson(payload, 3);
  emit(payload);
}

function getGuardedTable(doc, tableIndex) {
  var tableCount = Number(doc.Tables.Count);
  if (tableIndex > tableCount) die("table index out of range; document has " + tableCount + " tables");
  var table = doc.Tables(tableIndex);
  requireFingerprint(tableFingerprint(table), "--expect-table-fingerprint", "--allow-unverified-target");
  return table;
}

function requireTableTrackChangesOff(doc, operation) {
  if (Boolean(doc.TrackRevisions) && !hasFlag("--allow-track-changes")) {
    die(operation + " requires Track Changes to be off for deterministic table structure; use --allow-track-changes only after explicit approval");
  }
}

function commandSwapCellText() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var word = getWord();
  var doc = getMutationDocument(word);
  requireTableTrackChangesOff(doc, "swap-cell-text");
  var table = getGuardedTable(doc, tableIndex);
  var from = resolveTableCell(table, "--from-cell", "--from-row", "--from-col", "source cell");
  var to = resolveTableCell(table, "--to-cell", "--to-row", "--to-col", "destination cell");
  if (Number(from.cell.Range.Start) === Number(to.cell.Range.Start)) die("source and destination cells must differ");
  var fromText = cellTextForSwap(from.cell);
  var toText = cellTextForSwap(to.cell);
  var documentJson = "{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var applied = null, rolledBack = false, readback = "null", failures = [], rollbackFailures = [];
  try {
    setCellText(from.cell, toText);
    setCellText(to.cell, fromText);
    applied = true;
    var matches = cellTextForSwap(from.cell) === toText && cellTextForSwap(to.cell) === fromText;
    readback = "{\"matches_requested\":" + boolJson(matches) + "}";
    if (!matches) throw new Error("cell text readback did not match the requested swap");
  } catch (e1) {
    failures.push(e1.message || String(e1));
    try { setCellText(from.cell, fromText); } catch (e2) { rollbackFailures.push("source:" + (e2.message || String(e2))); }
    try { setCellText(to.cell, toText); } catch (e3) { rollbackFailures.push("destination:" + (e3.message || String(e3))); }
    try {
      if (cellTextForSwap(from.cell) !== fromText || cellTextForSwap(to.cell) !== toText) throw new Error("original cell text was not restored");
    } catch (e4) { rollbackFailures.push("readback:" + (e4.message || String(e4))); }
    rolledBack = rollbackFailures.length === 0;
    applied = rolledBack ? false : null;
  }
  var errors = [], fingerprint = postWriteTableFingerprint(table, errors);
  var verified = applied === true && failures.length === 0 && errors.length === 0 && fingerprint !== null;
  var payload = "{\"ok\":" + boolJson(verified) + ",\"action\":\"swap-cell-text\",\"table\":" + tableIndex
    + ",\"from\":" + cellTargetJson(from) + ",\"to\":" + cellTargetJson(to)
    + ",\"document\":" + documentJson + ",\"applied\":" + (applied === null ? "null" : boolJson(applied))
    + ",\"verified\":" + boolJson(verified) + ",\"inspection_complete\":" + boolJson(verified) + ",\"readback\":" + readback
    + ",\"failure_count\":" + failures.length + ",\"failures\":" + stringArrayJson(failures)
    + ",\"rolled_back\":" + boolJson(rolledBack) + ",\"rollback_failures\":" + stringArrayJson(rollbackFailures)
    + ",\"errors\":" + stringArrayJson(errors) + ",\"fingerprint\":" + (verified ? q(fingerprint) : "null") + "}";
  if (!verified) failJson(payload, 3);
  emit(payload);
}

function emitTableStructureResult(table, contextJson, expectedRows, expectedCols, applied, errors) {
  // Post-write reads must retain the operation state instead of exiting through a preflight helper.
  var rows = null, cols = null;
  try {
    var rowCount = Number(table.Rows.Count);
    if (!isFinite(rowCount) || rowCount < 1 || rowCount !== Math.floor(rowCount)) throw new Error("invalid row count");
    rows = rowCount;
  } catch (rowError) { errors.push("readback:rows:" + (rowError.message || String(rowError))); }
  try {
    var colCount = Number(table.Columns.Count);
    if (!isFinite(colCount) || colCount < 1 || colCount !== Math.floor(colCount)) throw new Error("invalid column count");
    cols = colCount;
  } catch (colError) { errors.push("readback:cols:" + (colError.message || String(colError))); }
  var matches = rows === null || cols === null ? null : rows === expectedRows && cols === expectedCols;
  if (matches === false) errors.push("readback:dimension readback mismatch");
  var fingerprint = postWriteTableFingerprint(table, errors);
  var verified = applied === true && matches === true && errors.length === 0 && fingerprint !== null;
  var payload = "{\"ok\":" + boolJson(verified) + "," + contextJson
    + ",\"rows\":" + (rows === null ? "null" : rows) + ",\"cols\":" + (cols === null ? "null" : cols)
    + ",\"applied\":" + (applied === null ? "null" : boolJson(applied)) + ",\"verified\":" + boolJson(verified)
    + ",\"inspection_complete\":" + boolJson(verified) + ",\"readback\":{\"matches_requested\":" + (matches === null ? "null" : boolJson(matches))
    + ",\"expected_rows\":" + expectedRows + ",\"expected_cols\":" + expectedCols + "}"
    + (matches === false ? ",\"error\":\"dimension readback mismatch\"" : "")
    + ",\"errors\":" + stringArrayJson(errors) + ",\"fingerprint\":" + (verified ? q(fingerprint) : "null") + "}";
  if (!verified) failJson(payload, 3);
  emit(payload);
}

function commandInsertRow() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var beforeText = opt("--before", "");
  var atEnd = hasFlag("--at-end");
  if (Boolean(beforeText) === atEnd) die("insert-row requires exactly one of --before n or --at-end");
  var word = getWord();
  var doc = getMutationDocument(word);
  requireTableTrackChangesOff(doc, "insert-row");
  var table = getGuardedTable(doc, tableIndex);
  var dimensions = getRegularTableDimensions(table, "insert-row");
  var before = beforeText ? parsePositiveInt(beforeText, "--before") : 0;
  if (before > dimensions.rows) die("--before row is out of range; table has " + dimensions.rows + " rows");
  var contextJson = "\"action\":\"insert-row\",\"table\":" + tableIndex + ",\"before\":" + (before || "null")
    + ",\"at_end\":" + boolJson(atEnd) + ",\"document\":{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var applied = null, errors = [];
  try {
    if (atEnd) table.Rows.Add();
    else table.Rows.Add(table.Rows(before));
    applied = true;
  } catch (e) { errors.push("write:" + (e.message || String(e))); }
  emitTableStructureResult(table, contextJson, dimensions.rows + 1, dimensions.cols, applied, errors);
}

function commandDeleteRow() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var row = parsePositiveInt(opt("--row", "0"), "--row");
  var word = getWord();
  var doc = getMutationDocument(word);
  requireTableTrackChangesOff(doc, "delete-row");
  var table = getGuardedTable(doc, tableIndex);
  var dimensions = getRegularTableDimensions(table, "delete-row");
  if (dimensions.rows <= 1) die("refusing to delete the last row; use delete-table when deleting the whole table is intended");
  if (row > dimensions.rows) die("row index out of range; table has " + dimensions.rows + " rows");
  var contextJson = "\"action\":\"delete-row\",\"table\":" + tableIndex + ",\"deleted_row\":" + row
    + ",\"document\":{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var applied = null, errors = [];
  try { table.Rows(row).Delete(); applied = true; } catch (e) { errors.push("write:" + (e.message || String(e))); }
  emitTableStructureResult(table, contextJson, dimensions.rows - 1, dimensions.cols, applied, errors);
}

function commandInsertColumn() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var beforeText = opt("--before", "");
  var atEnd = hasFlag("--at-end");
  if (Boolean(beforeText) === atEnd) die("insert-column requires exactly one of --before n or --at-end");
  var word = getWord();
  var doc = getMutationDocument(word);
  requireTableTrackChangesOff(doc, "insert-column");
  var table = getGuardedTable(doc, tableIndex);
  var dimensions = getRegularTableDimensions(table, "insert-column");
  var before = beforeText ? parsePositiveInt(beforeText, "--before") : 0;
  if (before > dimensions.cols) die("--before column is out of range; table has " + dimensions.cols + " columns");
  var contextJson = "\"action\":\"insert-column\",\"table\":" + tableIndex + ",\"before\":" + (before || "null")
    + ",\"at_end\":" + boolJson(atEnd) + ",\"document\":{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var applied = null, errors = [];
  try {
    if (atEnd) table.Columns.Add();
    else table.Columns.Add(table.Columns(before));
    applied = true;
  } catch (e) { errors.push("write:" + (e.message || String(e))); }
  emitTableStructureResult(table, contextJson, dimensions.rows, dimensions.cols + 1, applied, errors);
}

function commandDeleteColumn() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var col = parsePositiveInt(opt("--col", "0"), "--col");
  var word = getWord();
  var doc = getMutationDocument(word);
  requireTableTrackChangesOff(doc, "delete-column");
  var table = getGuardedTable(doc, tableIndex);
  var dimensions = getRegularTableDimensions(table, "delete-column");
  if (dimensions.cols <= 1) die("refusing to delete the last column; use delete-table when deleting the whole table is intended");
  if (col > dimensions.cols) die("column index out of range; table has " + dimensions.cols + " columns");
  var contextJson = "\"action\":\"delete-column\",\"table\":" + tableIndex + ",\"deleted_col\":" + col
    + ",\"document\":{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var applied = null, errors = [];
  try { table.Columns(col).Delete(); applied = true; } catch (e) { errors.push("write:" + (e.message || String(e))); }
  emitTableStructureResult(table, contextJson, dimensions.rows, dimensions.cols - 1, applied, errors);
}

function commandSetCellShading() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var colorHex = String(opt("--color", "")).replace(/^#/, "").toUpperCase();
  var color = parseHexColor(colorHex);
  var word = getWord();
  var doc = getMutationDocument(word);
  var table = getGuardedTable(doc, tableIndex);
  var target = resolveTableCell(table, "--cell", "--row", "--col", "shading target");
  var shading = target.cell.Shading;
  var previousColor = Number(shading.BackgroundPatternColor);
  var previousTexture = Number(shading.Texture);
  var documentJson = "{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var applied = null, rolledBack = false, readback = "null", failures = [], rollbackFailures = [];
  try {
    shading.Texture = 0; // wdTextureNone: retain a solid background without a pattern.
    shading.BackgroundPatternColor = color;
    applied = true;
    var actualTexture = Number(shading.Texture), actualColor = Number(shading.BackgroundPatternColor);
    if (!isFinite(actualTexture) || !isFinite(actualColor)) throw new Error("shading readback returned a non-finite value");
    var matches = actualTexture === 0 && actualColor === Number(color);
    readback = "{\"matches_requested\":" + boolJson(matches) + ",\"color_value\":" + actualColor + ",\"texture\":" + actualTexture + "}";
    if (!matches) throw new Error("shading readback did not match requested texture or color");
  } catch (e1) {
    failures.push(e1.message || String(e1));
    try { shading.Texture = previousTexture; } catch (e2) { rollbackFailures.push("texture:" + (e2.message || String(e2))); }
    try { shading.BackgroundPatternColor = previousColor; } catch (e3) { rollbackFailures.push("color:" + (e3.message || String(e3))); }
    try {
      if (Number(shading.Texture) !== previousTexture || Number(shading.BackgroundPatternColor) !== previousColor) throw new Error("original shading values were not restored");
    } catch (e4) { rollbackFailures.push("readback:" + (e4.message || String(e4))); }
    rolledBack = rollbackFailures.length === 0;
    applied = rolledBack ? false : null;
  }
  var errors = [], fingerprint = postWriteTableFingerprint(table, errors);
  var verified = applied === true && failures.length === 0 && errors.length === 0 && fingerprint !== null;
  var payload = "{\"ok\":" + boolJson(verified) + ",\"action\":\"set-cell-shading\",\"table\":" + tableIndex + ",\"target\":" + cellTargetJson(target)
    + ",\"color\":" + q(colorHex) + ",\"color_value\":" + color + ",\"texture\":0"
    + ",\"document\":" + documentJson + ",\"applied\":" + (applied === null ? "null" : boolJson(applied))
    + ",\"verified\":" + boolJson(verified) + ",\"inspection_complete\":" + boolJson(verified) + ",\"readback\":" + readback
    + ",\"failure_count\":" + failures.length + ",\"failures\":" + stringArrayJson(failures)
    + ",\"rolled_back\":" + boolJson(rolledBack) + ",\"rollback_failures\":" + stringArrayJson(rollbackFailures)
    + ",\"errors\":" + stringArrayJson(errors) + ",\"fingerprint\":" + (verified ? q(fingerprint) : "null") + "}";
  if (!verified) failJson(payload, 3);
  emit(payload);
}

function commandSetBorders(scope) {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var styleName = String(opt("--style", "single")).toLowerCase();
  var lineStyle = parseBorderStyle(styleName);
  var edges = parseBorderEdges(opt("--edges", ""), scope);
  var colorHex = null;
  var color = 0;
  var widthText = null;
  var lineWidth = 4;
  if (lineStyle !== 0) {
    colorHex = String(opt("--color", "000000")).replace(/^#/, "").toUpperCase();
    color = parseHexColor(colorHex);
    widthText = String(opt("--width", "0.5"));
    lineWidth = parseBorderWidth(widthText);
  }
  var word = getWord();
  var doc = getMutationDocument(word);
  var table = getGuardedTable(doc, tableIndex);
  var target = null;
  var borders;
  if (scope === "table") {
    borders = table.Borders;
  } else {
    target = resolveTableCell(table, "--cell", "--row", "--col", "border target");
    borders = target.cell.Borders;
  }
  var documentJson = "{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var result = applyBordersAtomically(borders, edges.types, color, lineWidth, lineStyle);
  var errors = [], fingerprint = postWriteTableFingerprint(table, errors);
  var verified = result.failures.length === 0 && result.rollbackFailures.length === 0 && errors.length === 0 && fingerprint !== null;
  var payload = "{\"ok\":" + boolJson(verified) + ",\"action\":" + q("set-" + scope + "-borders")
    + ",\"table\":" + tableIndex + ",\"target\":" + (target ? cellTargetJson(target) : "null")
    + ",\"edges\":" + stringArrayJson(edges.names) + ",\"style\":" + q(styleName) + ",\"line_style\":" + lineStyle
    + ",\"color\":" + (colorHex ? q(colorHex) : "null") + ",\"width_points\":" + (widthText ? q(widthText) : "null")
    + ",\"line_width\":" + (lineStyle !== 0 ? lineWidth : "null")
    + ",\"failure_count\":" + result.failures.length + ",\"failures\":" + stringArrayJson(result.failures)
    + ",\"rolled_back\":" + boolJson(result.rolledBack) + ",\"rollback_failures\":" + stringArrayJson(result.rollbackFailures)
    + ",\"document\":" + documentJson + ",\"applied\":" + (result.applied === null ? "null" : boolJson(result.applied))
    + ",\"verified\":" + boolJson(verified) + ",\"inspection_complete\":" + boolJson(verified)
    + ",\"errors\":" + stringArrayJson(errors) + ",\"fingerprint\":" + (verified ? q(fingerprint) : "null") + "}";
  if (!verified) failJson(payload, 3);
  emit(payload);
}

function emitDeletionResult(doc, collectionName, contextJson, countBefore, applied, errors) {
  var remaining = null;
  try {
    var afterCount = Number(doc[collectionName].Count);
    if (!isFinite(afterCount) || afterCount < 0 || afterCount !== Math.floor(afterCount)) throw new Error("invalid remaining object count");
    remaining = afterCount;
  } catch (e) { errors.push("readback:" + (e.message || String(e))); }
  var expected = countBefore - 1, matches = remaining === null ? null : remaining === expected;
  if (matches === false) errors.push("readback:object count did not decrease by one; inspect the target and any pending revisions before retrying");
  var verified = applied === true && matches === true && errors.length === 0;
  var countName = collectionName === "Tables" ? "remaining_tables" : "remaining_equations";
  var payload = "{\"ok\":" + boolJson(verified) + "," + contextJson + ",\"" + countName + "\":" + (remaining === null ? "null" : remaining)
    + ",\"applied\":" + (applied === null ? "null" : boolJson(applied)) + ",\"verified\":" + boolJson(verified)
    + ",\"inspection_complete\":" + boolJson(verified) + ",\"readback\":{\"expected_remaining\":" + expected
    + ",\"matches_requested\":" + (matches === null ? "null" : boolJson(matches)) + "}"
    + ",\"errors\":" + stringArrayJson(errors) + ",\"fingerprint\":null}";
  if (!verified) failJson(payload, 3);
  emit(payload);
}

function commandDeleteTable() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "0"), "--table");
  var word = getWord();
  var doc = getMutationDocument(word);
  requireTableTrackChangesOff(doc, "delete-table");
  var tableCount = Number(doc.Tables.Count);
  if (!isFinite(tableCount) || tableCount < 0 || tableCount !== Math.floor(tableCount)) die("invalid table count before deletion");
  if (tableIndex > tableCount) die("table index out of range; document has " + tableCount + " tables");
  var table = doc.Tables(tableIndex);
  requireFingerprint(tableFingerprint(table), "--expect-table-fingerprint", "--allow-unverified-target");
  var contextJson = "\"action\":\"delete-table\",\"table\":" + tableIndex + ",\"track_revisions\":" + boolJson(Boolean(doc.TrackRevisions))
    + ",\"document\":{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var applied = null, errors = [];
  try {
    table.Delete();
    applied = true;
  } catch (e) { errors.push("write:" + (e.message || String(e))); }
  emitDeletionResult(doc, "Tables", contextJson, tableCount, applied, errors);
}

function commandNormalizeTableBorders() {
  requireYes();
  var tableIndex = parsePositiveInt(opt("--table", "1"), "--table");
  var colorHex = opt("--color", "D9DEE8");
  var lineWidth = parsePositiveInt(opt("--line-width", "4"), "--line-width");
  var color = parseHexColor(colorHex);
  var word = getWord();
  var doc = getMutationDocument(word);
  var tableCount = Number(doc.Tables.Count);
  if (tableIndex > tableCount) die("table index out of range; document has " + tableCount + " tables");
  var table = doc.Tables(tableIndex);
  requireFingerprint(tableFingerprint(table), "--expect-table-fingerprint", "--allow-unverified-target");
  var tableBorderTypes = [-1, -2, -3, -4, -5, -6]; // top, left, bottom, right, insideH, insideV
  var cellBorderTypes = [-1, -2, -3, -4];
  var cellCount = 0;
  var failures = [];
  var documentJson = "{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";

  try { failures = failures.concat(setBorderCollection(table.Borders, tableBorderTypes, color, lineWidth, "table")); }
  catch (tableError) { failures.push("table:" + (tableError.message || String(tableError))); }
  try { table.Borders.Enable = true; } catch (e1) { failures.push("table-enable:" + (e1.message || String(e1))); }

  try {
    cellCount = Number(table.Range.Cells.Count);
    for (var i = 1; i <= cellCount; i++) {
      try {
        failures = failures.concat(setBorderCollection(table.Range.Cells(i).Borders, cellBorderTypes, color, lineWidth, "cell[" + i + "]"));
      } catch (e2) {
        failures.push("cell[" + i + "]:" + (e2.message || String(e2)));
      }
    }
  } catch (e3) {
    failures.push("cell-enumeration:" + (e3.message || String(e3)));
  }

  var errors = [], fingerprint = postWriteTableFingerprint(table, errors);
  var verified = failures.length === 0 && errors.length === 0 && fingerprint !== null;
  var payload = "{\"ok\":" + boolJson(verified) + ",\"action\":\"normalize-table-borders\",\"table\":" + tableIndex
    + ",\"tables\":" + tableCount
    + ",\"cells_seen\":" + cellCount
    + ",\"failure_count\":" + failures.length
    + ",\"failures\":" + stringArrayJson(failures)
    + ",\"color\":" + q(colorHex)
    + ",\"line_width\":" + lineWidth
    + ",\"document\":" + documentJson + ",\"applied\":" + (failures.length ? "null" : "true")
    + ",\"verified\":" + boolJson(verified) + ",\"inspection_complete\":" + boolJson(verified)
    + ",\"errors\":" + stringArrayJson(errors) + ",\"fingerprint\":" + (verified ? q(fingerprint) : "null") + "}";
  if (!verified) failJson(payload, 3);
  emit(payload);
}

function commandInsertEquation() {
  requireYes();
  var linearText = equationInputText();
  if (!linearText) die("equation input is empty after conversion");
  var at = opt("--at", "selection");
  if (at !== "selection" && at !== "end") die("--at must be selection or end");
  var word = getWord();
  var doc = getMutationDocument(word);
  var range;
  if (at === "end") range = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
  else range = requireExpectedSelection(word, true).range;
  try {
    var equation = buildEquationInRange(doc, range, linearText);
  } catch (e) {
    die("failed to insert equation: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"insert-equation\",\"linear\":" + q(linearText)
    + ",\"equation_count\":" + Number(doc.OMaths.Count)
    + ",\"fingerprint\":" + q(equationFingerprint(equation.OMaths(1))) + "}");
}

function commandSetEquation() {
  requireYes();
  var idx = parsePositiveInt(opt("--index", "0"), "--index");
  var linearText = equationInputText();
  if (!linearText) die("equation input is empty after conversion");
  var word = getWord();
  var doc = getMutationDocument(word);
  var count = Number(doc.OMaths.Count);
  if (idx > count) die("equation index out of range; document has " + count + " equations");
  requireFingerprint(equationFingerprint(doc.OMaths(idx)), "--expect-equation-fingerprint", "--allow-unverified-target");
  try {
    var oldRange = doc.OMaths(idx).Range;
    var range = oldRange.Duplicate;
    buildEquationInRange(doc, range, linearText);
  } catch (e) {
    die("failed to set equation: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"set-equation\",\"index\":" + idx + ",\"linear\":" + q(linearText)
    + ",\"equation_count\":" + Number(doc.OMaths.Count)
    + ",\"fingerprint\":" + q(equationFingerprint(doc.OMaths(idx))) + "}");
}

function commandDeleteEquation() {
  requireYes();
  var idx = parsePositiveInt(opt("--index", "0"), "--index");
  var word = getWord();
  var doc = getMutationDocument(word);
  var count = Number(doc.OMaths.Count);
  if (!isFinite(count) || count < 0 || count !== Math.floor(count)) die("invalid equation count before deletion");
  if (idx > count) die("equation index out of range; document has " + count + " equations");
  var equation = doc.OMaths(idx);
  requireFingerprint(equationFingerprint(equation), "--expect-equation-fingerprint", "--allow-unverified-target");
  var tracked = Boolean(doc.TrackRevisions);
  var contextJson = "\"action\":\"delete-equation\",\"index\":" + idx + ",\"track_revisions\":" + boolJson(tracked)
    + ",\"document\":{\"name\":" + q(doc.Name) + ",\"path\":" + q(safeDocPath(doc)) + "}";
  var prepareInline = !tracked && Number(equation.Type) === 0;
  var applied = null, deletedUnits = null, preparedInline = false, errors = [];
  if (prepareInline) {
    // A display equation can leave an empty OMath at its paragraph mark after Range.Delete.
    preparedInline = null;
    try {
      equation.Type = 1; // wdOMathInline; keep tracked revision objects in their original form.
      if (Number(equation.Type) !== 1) throw new Error("equation did not become inline before deletion");
      preparedInline = true;
    } catch (prepareError) { errors.push("prepare:" + (prepareError.message || String(prepareError))); }
  }
  if (preparedInline !== null) try {
    var reportedUnits = Number(equation.Range.Delete());
    if (!isFinite(reportedUnits) || reportedUnits < 0 || reportedUnits !== Math.floor(reportedUnits)) throw new Error("Range.Delete returned an invalid deletion count");
    deletedUnits = reportedUnits;
    applied = deletedUnits > 0;
    if (!applied) errors.push("write:Range.Delete returned 0; deletion was not reported as successful");
  } catch (e) { errors.push("write:" + (e.message || String(e))); }
  emitDeletionResult(doc, "OMaths", contextJson + ",\"deleted_units\":" + (deletedUnits === null ? "null" : deletedUnits)
    + ",\"prepared_inline\":" + (preparedInline === null ? "null" : boolJson(preparedInline)), count, applied, errors);
}

function confirmSavedDocument(word, doc) {
  // Save can return after cancellation or while background saving is queued.
  var deadline = new Date().getTime() + 30000;
  while (Number(word.BackgroundSavingStatus) !== 0) {
    if (new Date().getTime() >= deadline) throw new Error("saving is still pending");
    WScript.Sleep(100);
  }
  if (doc.Saved !== true) throw new Error("save was cancelled or could not be confirmed");
  var path = safeDocPath(doc);
  if (!String(doc.Path) || !path) throw new Error("save has no confirmed file path");
  verifyOutputFile(path, false);
  return path;
}

function saveAndConfirm(word, doc) {
  doc.Save();
  return confirmSavedDocument(word, doc);
}

function commandSaveActive() {
  requireYes();
  var word = getWord();
  var doc = getMutationDocument(word);
  try {
    saveAndConfirm(word, doc);
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
  if (hasFlag("--quit-if-empty")) die("--quit-if-empty is disabled; normal commands must not quit a user-owned Word instance");
  var word = getWord();
  var doc = getMutationDocument(word);
  var path = safeDocPath(doc);
  try {
    if (save) path = saveAndConfirm(word, doc);
    // Preserve edits introduced by a close event after our explicit save check.
    doc.Close(save ? -1 : 0);
  } catch (e) {
    die("failed to close active document: " + e.message);
  }
  emit("{\"ok\":true,\"action\":\"close-active\",\"saved\":" + q(String(save)) + ",\"path\":" + q(path) + "}");
}

function commandSaveCopy() {
  requireYes();
  var path = prepareDocumentOutputPath(opt("--path", ""), "save-copy");
  var word = getWord();
  var doc = getMutationDocument(word);
  var source = safeDocPath(doc);
  if (source && sameOutputFile(source, path)) die("backup path must differ from the active document path");
  var savedBefore = false;
  try { savedBefore = doc.Saved === true; } catch (e0) {}
  var stage = null, failure = null, method = "SaveCopyAs";
  try {
    stage = createOutputStage(path);
    try {
      doc.SaveCopyAs(stage.output);
    } catch (copyError) {
      if (!source || !fso.FileExists(source)) throw new Error("active document has no saved source path: " + copyError.message);
      if (!savedBefore || doc.Saved !== true || canonicalPath(safeDocPath(doc)) !== canonicalPath(source)) {
        throw new Error("active document has unsaved changes or changed identity; refusing stale disk-copy fallback: " + copyError.message);
      }
      if (Number(word.BackgroundSavingStatus) !== 0) throw new Error("source saving is still pending; refusing disk-copy fallback");
      if (fso.FileExists(stage.output)) fso.DeleteFile(stage.output, true);
      fso.CopyFile(source, stage.output, false);
      if (doc.Saved !== true || canonicalPath(safeDocPath(doc)) !== canonicalPath(source)) throw new Error("document changed during disk-copy fallback");
      method = "FileSystemObject.CopyFile";
    }
    verifyOutputFile(stage.output, false);
    publishOutput(stage, path);
  } catch (e) { failure = e; }
  var warning = cleanupOutputStage(stage);
  if (failure) die("failed to save copy: " + failure.message + (warning ? "; " + warning : ""));
  // Result delivery must not enter the generation fallback after publication.
  emit("{\"ok\":true,\"action\":\"save-copy\",\"method\":" + q(method) + ",\"path\":" + q(path)
    + (method === "FileSystemObject.CopyFile" ? ",\"source\":" + q(source) : "")
    + ",\"document_saved_before\":" + boolJson(savedBefore) + ",\"includes_current_document_state\":true"
    + (warning ? ",\"cleanup_warning\":" + q(warning) : "") + "}");
}

function commandExportPdf() {
  requireYes();
  var path = prepareDocumentOutputPath(opt("--path", ""), "export-pdf");
  if (!/\.pdf$/i.test(path)) die("export-pdf output must use a .pdf extension");
  var word = getWord();
  var doc = getMutationDocument(word);
  if (sameOutputFile(safeDocPath(doc), path)) die("PDF output path must differ from the active document path");
  var stage = null, failure = null;
  try {
    stage = createOutputStage(path);
    doc.ExportAsFixedFormat(stage.output, 17);
    verifyOutputFile(stage.output, true);
    publishOutput(stage, path);
  } catch (e) { failure = e; }
  var warning = cleanupOutputStage(stage);
  if (failure) die("failed to export PDF: " + failure.message + (warning ? "; " + warning : ""));
  emit("{\"ok\":true,\"action\":\"export-pdf\",\"path\":" + q(path)
    + (warning ? ",\"cleanup_warning\":" + q(warning) : "") + "}");
}

function commandOpen() {
  var path = absPath(opt("--path", ""));
  if (!path || !fso.FileExists(path)) die("open requires existing --path file");
  var word, owned = false;
  try {
    word = GetObject("", "Word.Application");
  } catch (e) {
    word = new ActiveXObject("Word.Application");
    owned = true;
  }
  var previousSecurity = null, restored = null, attempted = false, opened = false, errors = [];
  try {
    var security = Number(word.AutomationSecurity);
    if (security !== 1 && security !== 2 && security !== 3) throw new Error("unrecognized automation security setting");
    previousSecurity = security;
    word.AutomationSecurity = 3; // msoAutomationSecurityForceDisable, only during this open call.
    if (Number(word.AutomationSecurity) !== 3) throw new Error("cannot confirm that document macros are disabled");
    word.Visible = true;
    attempted = true;
    if (!word.Documents.Open(path)) throw new Error("Word did not return an opened document");
    opened = true;
  } catch (e) { errors.push((attempted ? "open:" : "prepare:") + (e.message || String(e))); }
  if (previousSecurity !== null) {
    try {
      word.AutomationSecurity = previousSecurity;
      if (Number(word.AutomationSecurity) !== previousSecurity) throw new Error("original automation security setting was not restored");
      restored = true;
    } catch (restoreError) { restored = false; errors.push("restore:" + (restoreError.message || String(restoreError))); }
  }
  if (errors.length && owned) {
    try { if (Number(word.Documents.Count) === 0) word.Quit(0); }
    catch (cleanupError) { errors.push("cleanup:" + (cleanupError.message || String(cleanupError))); }
  }
  var ok = opened && restored === true && errors.length === 0;
  var payload = "{\"ok\":" + boolJson(ok) + ",\"action\":\"open\",\"path\":" + q(path)
    + ",\"opened\":" + (opened ? "true" : attempted ? "null" : "false")
    + ",\"automation_security_restored\":" + (restored === null ? "null" : boolJson(restored))
    + ",\"errors\":" + stringArrayJson(errors) + "}";
  if (!ok) failJson(payload, 3);
  emit(payload);
}

function commandSmoke() {
  requireYes();
  var path = prepareDocumentOutputPath(opt("--path", ""), "smoke");
  if (!/\.docx$/i.test(path)) die("smoke output must use a .docx extension");
  var word = null;
  var doc = null;
  var stage = null;
  var tableCount = 0;
  var equationCount = 0;
  var failure = null;
  try {
    stage = createOutputStage(path);
    word = new ActiveXObject("Word.Application");
    word.Visible = false;
    word.DisplayAlerts = 0;
    doc = word.Documents.Add();
    doc.Content.Text = "Word control smoke test.\rCreated by word_control.js.\r";

    var tableRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
    var table = doc.Tables.Add(tableRange, 2, 2);
    setCellText(table.Cell(1, 1), "A");
    setCellText(table.Cell(1, 2), "B");
    setCellText(table.Cell(2, 1), "1");
    setCellText(table.Cell(2, 2), "2");
    var borderFailures = setBorderCollection(table.Borders, [-1, -2, -3, -4, -5, -6], parseHexColor("D9DEE8"), 4, "smoke-table");
    if (borderFailures.length) throw new Error("border smoke failed: " + borderFailures.join("; "));

    var mathRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
    buildEquationInRange(doc, mathRange, "x^2+y^2=z^2");
    tableCount = Number(doc.Tables.Count);
    equationCount = Number(doc.OMaths.Count);
    if (tableCount !== 1 || equationCount !== 1) throw new Error("unexpected smoke object counts");

    // Explicit DOCX format; task-only staging paths must not enter Word's recent files.
    try { doc.SaveAs2(stage.output, 12, false, "", false); } catch (e1) { doc.SaveAs(stage.output, 12, false, "", false); }
    if (!sameOutputFile(confirmSavedDocument(word, doc), stage.output)) throw new Error("smoke save changed the expected output path");
    doc.Close(false);
    doc = null;
    publishOutput(stage, path);
  } catch (e2) {
    failure = e2;
  } finally {
    if (doc !== null) {
      try { doc.Close(false); } catch (e3) {}
      doc = null;
    }
    if (word !== null) {
      try { word.Quit(0); } catch (e4) {}
      word = null;
    }
    try { CollectGarbage(); } catch (e5) {}
  }
  var warning = cleanupOutputStage(stage);
  if (failure !== null) die("smoke test failed: " + (failure.message || String(failure)) + (warning ? "; " + warning : ""));
  emit("{\"ok\":true,\"action\":\"smoke\",\"path\":" + q(path)
    + ",\"table_count\":" + tableCount + ",\"equation_count\":" + equationCount
    + (warning ? ",\"cleanup_warning\":" + q(warning) : "") + "}");
}

try {
  preflightOutput();
  if (COMMAND === "help" || COMMAND === "--help" || COMMAND === "-h") commandHelp();
  else if (COMMAND === "status") commandStatus();
  else if (COMMAND === "selection") commandSelection();
  else if (COMMAND === "selection-info") commandSelectionInfo();
  else if (COMMAND === "document-text") commandDocumentText();
  else if (COMMAND === "paragraphs") commandParagraphs();
  else if (COMMAND === "find-text") commandFindText();
  else if (COMMAND === "tables") commandTables();
  else if (COMMAND === "equations") commandEquations();
  else if (COMMAND === "convert-equation") commandConvertEquation();
  else if (COMMAND === "enable-track-changes") commandTrack(true);
  else if (COMMAND === "disable-track-changes") commandTrack(false);
  else if (COMMAND === "replace-selection") commandReplaceSelection();
  else if (COMMAND === "replace-paragraph") commandReplaceParagraph();
  else if (COMMAND === "insert-comment") commandInsertComment();
  else if (COMMAND === "create-table") commandCreateTable();
  else if (COMMAND === "set-cell") commandSetCell();
  else if (COMMAND === "swap-cell-text") commandSwapCellText();
  else if (COMMAND === "insert-row") commandInsertRow();
  else if (COMMAND === "delete-row") commandDeleteRow();
  else if (COMMAND === "insert-column") commandInsertColumn();
  else if (COMMAND === "delete-column") commandDeleteColumn();
  else if (COMMAND === "set-cell-shading") commandSetCellShading();
  else if (COMMAND === "set-cell-borders") commandSetBorders("cell");
  else if (COMMAND === "set-table-borders") commandSetBorders("table");
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
