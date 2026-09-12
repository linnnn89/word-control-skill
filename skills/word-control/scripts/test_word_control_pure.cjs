// Node-only regression checks; .cjs avoids inheriting a host package's module mode.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

if (process.argv[2] === '--validate-json-dir') {
  const directory = process.argv[3];
  const files = fs.readdirSync(directory).filter(name => name.endsWith('.json'));
  assert.ok(files.length > 0, 'No JSON outputs to validate');
  for (const file of files) {
    JSON.parse(fs.readFileSync(path.join(directory, file), 'utf8').replace(/^\uFEFF/, ''));
  }
  console.log(JSON.stringify({ ok: true, strict_json_files: files.length }));
  process.exit(0);
}

const source = fs.readFileSync(path.join(__dirname, 'word_control.js'), 'utf8');
const existingFiles = new Set();
const writes = [];
const context = {
  ActiveXObject: function (name) {
    if (name === 'Scripting.FileSystemObject') return {
      GetAbsolutePathName: value => path.win32.resolve(value),
      GetParentFolderName: value => path.win32.dirname(value),
      FolderExists: () => true,
      FileExists: value => existingFiles.has(value),
    };
    if (name === 'ADODB.Stream') return {
      Open() {}, Close() {}, WriteText() {},
      SaveToFile(file, mode) { writes.push({ file, mode }); },
    };
    throw new Error('Real COM is prohibited: ' + name);
  },
  WScript: {
    Arguments: [], Echo() {},
    StdErr: { WriteLine(message) { context.lastError = message; } },
    Quit(code) { throw new Error(context.lastError || 'Exit ' + code); },
  },
};
vm.createContext(context);
vm.runInContext(source, context);
let passed = 0;
function test(name, fn) { fn(); passed++; }

test('All JSON control characters round-trip through a strict parser', () => {
  for (let code = 0; code < 32; code++) {
    const text = 'a' + String.fromCharCode(code) + 'b';
    assert.equal(JSON.parse(context.q(text)), text);
  }
  const text = '中文 "quoted" \\ text';
  assert.equal(JSON.parse(context.q(text)), text);
});

test('LaTeX validates original tokens and braces', () => {
  context.ARGS = ['convert-equation', '--format', 'latex'];
  for (const input of ['\\leftarrow', '\\alphabogus', '\\matrix{a&b}', 'x^{2', 'x^2}']) {
    context.readUtf8 = () => input;
    assert.throws(() => context.equationInputText(), /unsupported|unbalanced/);
  }
  context.readUtf8 = () => '\\frac{a+b}{c} + \\alpha';
  assert.equal(context.equationInputText(), '(a+b)/c + α');
  context.readUtf8 = () => '\\sqrt{\\frac{a}{b}}';
  assert.equal(context.equationInputText(), '√(a/b)');
});

const range = { Start: 0, End: 4, Text: 'same', StoryType: 1 };
const selection = context.selectionSnapshot({ Selection: { Range: range } });
const guards = ['--expect-story-type', '1', '--expect-start', '0', '--expect-end', '4', '--expect-selection-hash', selection.hash];
test('Identical text and coordinates in a different story are rejected', () => {
  context.ARGS = ['replace-selection', ...guards];
  assert.throws(() => context.requireExpectedSelection({ Selection: { Range: { ...range, StoryType: 7 } } }, false), /story changed/);
  assert.equal(context.requireExpectedSelection({ Selection: { Range: range } }, false).hash, selection.hash);
  context.ARGS = ['replace-selection', ...guards.slice(2)];
  assert.throws(() => context.requireExpectedSelection({ Selection: { Range: range } }, false), /expect-story-type/);
});

test('A saved document cannot be verified by filename alone', () => {
  context.ARGS = ['save-active', '--expect-name', 'report.docx'];
  assert.throws(() => context.requireExpectedDocument({ Name: 'report.docx', Path: 'C:\\other', FullName: 'C:\\other\\report.docx' }), /saved documents require/);
  context.requireExpectedDocument({ Name: 'report.docx', Path: '' });
});

test('Failed tracked replacement restores previous tracking state', () => {
  const doc = { Name: 'test.docx', FullName: 'C:\\test\\test.docx', TrackRevisions: false };
  const target = { ...range };
  Object.defineProperty(target, 'Text', { get() { return 'same'; }, set() { throw new Error('Protected range'); } });
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: doc, Selection: { Range: target } });
  context.readUtf8 = () => 'changed';
  context.ARGS = ['replace-selection', '--yes', '--track', '--expect-path', doc.FullName, ...guards];
  assert.throws(() => context.commandReplaceSelection(), /Protected range/);
  assert.equal(doc.TrackRevisions, false);
  doc.TrackRevisions = true;
  assert.throws(() => context.commandReplaceSelection(), /Protected range/);
  assert.equal(doc.TrackRevisions, true);
});

test('Incomplete formatting inspection has no reusable fingerprint', () => {
  const borders = () => { throw new Error('Formatting unavailable'); };
  const cell = { Range: { Text: 'cell\r\x07' }, Shading: { BackgroundPatternColor: 0, Texture: 0 }, Borders: borders };
  const cells = () => cell; cells.Count = 1;
  const table = { Rows: { Count: 1 }, Columns: { Count: 1 }, Range: { Text: 'cell\r\x07', Cells: cells }, Cell: () => cell, Borders: borders };
  const errors = [];
  assert.equal(context.tableFingerprint(table, errors), null);
  assert.ok(errors.length > 0);
  assert.throws(() => context.tableFingerprint(table), /cannot verify/);
  context.ARGS = ['set-cell', '--allow-unverified-target'];
  assert.throws(() => context.requireFingerprint(null, '--expect-table-fingerprint', '--allow-unverified-target'), /incomplete/);
  const tables = () => table; tables.Count = 1;
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: { Tables: tables } });
  context.ARGS = ['tables'];
  let result;
  context.emit = value => { result = JSON.parse(value); };
  context.commandTables();
  assert.equal(result.tables[0].inspection_complete, false);
  assert.equal(result.tables[0].fingerprint, null);
});

test('Output preflight rejects document paths, collisions, and unapproved overwrite', () => {
  context.ARGS = ['help', '--output', 'C:\\test\\source.docx', '--overwrite'];
  assert.throws(() => context.preflightOutput(), /scratch/);
  context.ARGS = ['convert-equation', '--output', 'C:\\test\\input.txt', '--input', 'C:\\test\\input.txt'];
  assert.throws(() => context.preflightOutput(), /must differ/);
  existingFiles.add('C:\\test\\status.json');
  context.ARGS = ['status', '--output', 'C:\\test\\status.json'];
  assert.throws(() => context.preflightOutput(), /overwrite/);
  context.ARGS.push('--overwrite');
  context.preflightOutput();
  context.ARGS = ['status', '--output', 'C:\\test\\new.json'];
  context.writeUtf8('C:\\test\\new.json', '{}');
  assert.equal(writes.at(-1).mode, 1, 'New outputs must use exclusive creation');
});

test('Paragraph pages preserve legacy output and read only the requested absolute indices', () => {
  let allowed = [1, 80];
  const visited = [];
  const paragraphs = index => {
    assert.ok(index >= allowed[0] && index <= allowed[1], 'Unrequested paragraph read');
    visited.push(index);
    return { Range: { Text: 'paragraph ' + index + '\r' } };
  };
  paragraphs.Count = 85;
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: { Paragraphs: paragraphs } });
  let result;
  context.emit = value => { result = JSON.parse(value); };
  context.ARGS = ['paragraphs'];
  context.commandParagraphs();
  assert.equal(result.returned, 80);
  assert.equal(result.paragraphs[79].index, 80);
  assert.equal('next_from' in result, false, 'Legacy output shape must stay unchanged');

  allowed = [1, 2]; visited.length = 0;
  context.ARGS = ['paragraphs', '--max', '2'];
  context.commandParagraphs();
  assert.deepEqual(result, { ok: true, paragraph_count: 85, returned: 2,
    paragraphs: [{ index: 1, text: 'paragraph 1\n' }, { index: 2, text: 'paragraph 2\n' }] });
  assert.deepEqual(visited, [1, 2]);

  allowed = [43, 44]; visited.length = 0;
  context.ARGS = ['paragraphs', '--from', '43', '--max', '2'];
  context.commandParagraphs();
  assert.deepEqual(visited, [43, 44]);
  assert.equal(result.from, 43);
  assert.equal(result.next_from, 45);
  assert.deepEqual(result.paragraphs.map(item => item.index), [43, 44]);
  assert.equal(result.paragraphs[0].text, 'paragraph 43\n');

  allowed = [85, 85]; visited.length = 0;
  context.ARGS = ['paragraphs', '--from', '85', '--max', '5'];
  context.commandParagraphs();
  assert.equal(result.returned, 1);
  assert.equal(result.next_from, null);
  context.ARGS = ['paragraphs', '--from', '86'];
  visited.length = 0;
  context.commandParagraphs();
  assert.equal(result.returned, 0);
  assert.equal(result.next_from, null);
  assert.deepEqual(visited, []);
  for (const invalid of ['0', '-1', 'abc']) {
    context.ARGS = ['paragraphs', '--from', invalid];
    assert.throws(() => context.commandParagraphs(), /positive integer/);
  }
});

test('Failed reads stay unknown and equation text/fingerprints share one snapshot', () => {
  let result;
  context.emit = value => { result = JSON.parse(value); };
  const paragraphs = index => {
    if (index === 2) throw new Error('Paragraph unavailable');
    return { Range: { Text: index === 1 ? '\r' : 'readable\r' } };
  };
  paragraphs.Count = 3;
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: { Paragraphs: paragraphs } });
  context.ARGS = ['paragraphs'];
  context.commandParagraphs();
  assert.equal(result.inspection_complete, false);
  assert.equal(result.paragraphs[0].text, '\n');
  assert.equal(result.paragraphs[1].text, null, 'Unreadable is distinct from empty');
  assert.equal(result.paragraphs[2].text, 'readable\n');
  assert.match(result.read_errors[0], /paragraph\[2\].*unavailable/);

  const reads = [0, 0, 0];
  const equations = index => ({ Range: { get Text() {
    reads[index - 1]++;
    if (index === 2) throw new Error('Equation unavailable');
    return reads[index - 1] === 1 ? 'x\r' : 'changed';
  } } });
  equations.Count = 3;
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: { OMaths: equations } });
  context.ARGS = ['equations'];
  context.commandEquations();
  assert.deepEqual(reads, [1, 1, 1], 'Read text exactly once per equation');
  assert.equal(result.equations[0].text, 'x\n');
  assert.equal(result.equations[0].fingerprint, context.textHash('x\n'));
  assert.equal(result.equations[1].text, null);
  assert.equal(result.equations[1].fingerprint, null);
  assert.equal(result.equations[1].inspection_complete, false);
  assert.match(result.equations[1].read_errors[0], /unavailable/);
  assert.equal(result.equations[2].text, 'x\n');
});

test('Single-table inspection is bounded and read-only detail cannot supply mutation fingerprints', () => {
  let summaryOnly = false;
  let textOnly = false;
  const border = { LineStyle: 1, LineWidth: 4, Color: 0 };
  const makeTable = label => {
    const items = Array.from({ length: 4 }, (_, index) => ({
      get Range() { assert.equal(summaryOnly, false, 'Summary read a cell'); return { Text: label + index + '\r\x07' }; },
      get Shading() { assert.equal(summaryOnly || textOnly, false, 'Read-only detail read shading'); return { BackgroundPatternColor: 0, Texture: 0 }; },
      Borders() { assert.equal(summaryOnly || textOnly, false, 'Read-only detail read cell borders'); return border; },
    }));
    const cells = index => { assert.equal(summaryOnly, false, 'Summary enumerated cells'); return items[index - 1]; };
    cells.Count = items.length;
    return { Rows: { Count: 2 }, Columns: { Count: 2 },
      Range: { Cells: cells, get Text() { assert.equal(summaryOnly, false, 'Summary read table text'); return label + '\r\x07'; } },
      Cell: (row, column) => items[(row - 1) * 2 + column - 1],
      Borders() { assert.equal(summaryOnly || textOnly, false, 'Read-only detail read table borders'); return border; } };
  };
  const items = ['first', 'second', 'third'].map(makeTable);
  let allowed = [1, 2, 3];
  const visited = [];
  const tables = index => { assert.ok(allowed.includes(index), 'Unrequested table read'); visited.push(index); return items[index - 1]; };
  tables.Count = items.length;
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: { Tables: tables } });
  let result;
  context.emit = value => { result = JSON.parse(value); };
  context.ARGS = ['tables'];
  context.commandTables();
  const legacy = result;
  assert.equal(legacy.returned, 3);
  assert.equal('detail' in legacy, false);

  allowed = [2]; visited.length = 0;
  context.ARGS = ['tables', '--table', '2'];
  context.commandTables();
  assert.deepEqual(visited, [2]);
  assert.equal(result.table_count, 3);
  assert.equal(result.returned, 1);
  assert.deepEqual(result.tables[0], legacy.tables[1]);

  textOnly = true; visited.length = 0;
  context.ARGS = ['tables', '--table', '2', '--detail', 'text'];
  context.commandTables();
  assert.deepEqual(visited, [2]);
  assert.deepEqual(result.tables[0].cells, legacy.tables[1].cells);
  assert.equal(result.tables[0].fingerprint, null);
  assert.equal(result.tables[0].inspection_complete, false);
  assert.deepEqual(result.tables[0].read_errors, []);
  context.ARGS = ['set-cell', '--allow-unverified-target'];
  assert.throws(() => context.requireFingerprint(result.tables[0].fingerprint, '--expect-table-fingerprint', '--allow-unverified-target'), /incomplete/);
  textOnly = false;

  summaryOnly = true; visited.length = 0;
  context.ARGS = ['tables', '--table', '2', '--detail', 'summary'];
  context.commandTables();
  assert.deepEqual(visited, [2]);
  assert.equal(result.detail, 'summary');
  const brief = result.tables[0];
  assert.equal(brief.index, 2);
  assert.equal(brief.cell_count, 4);
  assert.equal(brief.layout, 'unverified');
  assert.equal(brief.inspection_complete, false);
  assert.equal(brief.fingerprint, null);
  assert.deepEqual(brief.cells, []);
  assert.deepEqual(brief.read_errors, []);
  context.ARGS = ['set-cell', '--allow-unverified-target'];
  assert.throws(() => context.requireFingerprint(brief.fingerprint, '--expect-table-fingerprint', '--allow-unverified-target'), /incomplete/);

  allowed = [1]; visited.length = 0;
  context.ARGS = ['tables', '--max', '1', '--detail', 'summary'];
  context.commandTables();
  assert.deepEqual(visited, [1]);
  for (const [args, message] of [
    [['--table', '2', '--max', '1'], /--table.*--max/],
    [['--table', '4'], /out of range/],
    [['--table', '0'], /positive integer/],
    [['--detail', 'brief'], /full, text or summary/],
  ]) {
    context.ARGS = ['tables', ...args];
    assert.throws(() => context.commandTables(), message);
  }
});

console.log(JSON.stringify({ ok: true, pure_regression_groups: passed }));
