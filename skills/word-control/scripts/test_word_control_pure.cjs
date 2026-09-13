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

  table.Borders = cell.Borders = () => ({ LineStyle: 1, LineWidth: 4, Color: 0 });
  const before = context.tableFingerprint(table);
  cell.Shading.Texture = 100;
  assert.notEqual(context.tableFingerprint(table), before, 'Texture must remain part of the guard');
  cell.Shading.Texture = 0;
  cell.Shading.BackgroundPatternColor = 255;
  assert.notEqual(context.tableFingerprint(table), before, 'Background color must remain part of the guard');
  for (const property of ['BackgroundPatternColor', 'Texture']) {
    cell.Shading = { BackgroundPatternColor: 0, Texture: 0 };
    Object.defineProperty(cell.Shading, property, { get() { throw new Error(property + ' unavailable'); } });
    const shadingErrors = [];
    assert.equal(context.tableFingerprint(table, shadingErrors), null);
    assert.ok(shadingErrors.some(error => error.includes(property + ' unavailable')));
  }
});

test('Table guards distinguish trailing paragraphs, line breaks and cell boundaries', () => {
  const borders = () => ({ LineStyle: 1, LineWidth: 4, Color: 0 });
  const cell = { Shading: { BackgroundPatternColor: 0, Texture: 0 }, Borders: borders };
  const cells = () => cell; cells.Count = 2;
  const table = { Rows: { Count: 1 }, Columns: { Count: 2 }, Range: { Text: '', Cells: cells }, Borders: borders };
  const states = ['A\r\x07B\r\x07', 'A\r\x07B\r\r\x07', 'A\rB\r\x07', 'A\n\x07B\r\x07', 'A\x0b\x07B\r\x07'];
  const fingerprints = states.map(text => { table.Range.Text = text; return context.tableFingerprint(table); });
  assert.equal(new Set(fingerprints).size, states.length, 'Distinct control characters must not collapse into the same guard input');
  table.Range.Text = states[1];
  context.ARGS = ['set-cell', '--expect-table-fingerprint', fingerprints[0]];
  assert.throws(() => context.requireFingerprint(context.tableFingerprint(table), '--expect-table-fingerprint', '--allow-unverified-target'), /changed/);
});

test('Cell text swaps preserve exact bodies on success and rollback, and reject nested cell markers before writing', () => {
  const original = { tableFingerprint: context.tableFingerprint, getWord: context.getWord, emit: context.emit };
  const initial = ['A\r\r', 'B\x0bC\t\r'];
  let bodies, writes, mode, result;
  const cell = index => ({ get Range() { return { Start: index * 20, End: index * 20 + 10,
    get Text() { return bodies[index] + '\r\x07'; },
    set Text(value) {
      writes++;
      if (mode === 'mismatch' && writes === 2) value = value.replace(/\r+$/, '');
      bodies[index] = value;
      if (mode === 'write-error' && writes === 2) throw new Error('Second write interrupted');
    } }; } });
  const cells = n => cell(n - 1); cells.Count = 2;
  const table = { Range: { Cells: cells } };
  const tables = () => table; tables.Count = 1;
  const doc = { Name: 'source.docx', Path: 'C:\\test', FullName: 'C:\\test\\source.docx', TrackRevisions: false, Tables: tables };
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: doc });
  context.tableFingerprint = () => 'guard';
  context.emit = text => { result = JSON.parse(text); };
  context.ARGS = ['swap-cell-text', '--table', '1', '--from-cell', '1', '--to-cell', '2', '--expect-path', doc.FullName, '--expect-table-fingerprint', 'guard', '--yes'];
  try {
    for (mode of ['success', 'mismatch', 'write-error', 'nested']) {
      bodies = initial.slice(); writes = 0; result = null; context.lastError = '';
      if (mode === 'nested') bodies[1] = 'nested\r\x07text';
      const before = bodies.slice();
      if (mode === 'success') {
        context.commandSwapCellText();
        assert.deepEqual(bodies, [initial[1], initial[0]]);
        assert.equal(result.ok, true);
      } else {
        assert.throws(() => context.commandSwapCellText(), mode === 'nested' ? /nested/ : /Exit 3/);
        assert.deepEqual(bodies, before);
        if (mode === 'nested') {
          assert.equal(result, null);
          assert.equal(writes, 0, 'Inspect both bodies before replacing either one');
        } else {
          assert.equal(result.ok, false);
          assert.equal(result.applied, false);
          assert.equal(result.rolled_back, true);
          assert.equal(result.fingerprint, null);
        }
      }
    }
  } finally { Object.assign(context, original); }
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
  reads.fill(0);
  context.ARGS = ['equations', '--index', '3'];
  context.commandEquations();
  assert.deepEqual(reads, [0, 0, 1]);
  assert.equal(result.equation_count, 3);
  assert.equal(result.returned, 1);
  assert.equal(result.equations[0].index, 3);
  context.ARGS = ['equations', '--index', '4'];
  assert.throws(() => context.commandEquations(), /out of range/);
});

test('Single-table inspection is bounded and read-only detail cannot supply mutation fingerprints', () => {
  let summaryOnly = false;
  let textOnly = false;
  let allowedCells = null;
  const border = { LineStyle: 1, LineWidth: 4, Color: 0 };
  const makeTable = label => {
    const items = Array.from({ length: 4 }, (_, index) => ({
      RowIndex: index < 2 ? 1 : 2, ColumnIndex: index % 2 + 1,
      get Range() { assert.equal(summaryOnly, false, 'Summary read a cell'); return { Text: label + index + '\r\x07' }; },
      get Shading() { assert.equal(summaryOnly || textOnly, false, 'Read-only detail read shading'); return { BackgroundPatternColor: 0, Texture: 0 }; },
      Borders() { assert.equal(summaryOnly || textOnly, false, 'Read-only detail read cell borders'); return border; },
    }));
    const cells = index => { assert.equal(summaryOnly, false, 'Summary enumerated cells');
      if (allowedCells) assert.ok(allowedCells.includes(index), 'Read an unrequested cell');
      return items[index - 1]; };
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

  textOnly = true; allowedCells = [3];
  context.ARGS = ['tables', '--table', '2', '--detail', 'text', '--cell-from', '3', '--cell-max', '1'];
  context.commandTables();
  assert.equal(result.tables[0].cell_count, 4);
  assert.equal(result.tables[0].returned_cells, 1);
  assert.equal(result.tables[0].next_cell, 4);
  assert.deepEqual(result.tables[0].linear_cells[0], { index: 3, row: 2, col: 1, text: 'second2' });
  assert.equal(result.tables[0].layout, 'unverified');
  assert.equal(result.tables[0].fingerprint, null);
  context.ARGS = ['tables', '--table', '2', '--detail', 'text', '--cell-from', '5'];
  context.commandTables();
  assert.equal(result.tables[0].returned_cells, 0);
  assert.equal(result.tables[0].next_cell, null);
  allowedCells = null; textOnly = false;

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
    [['--cell-from', '1'], /requires --table and --detail text/],
    [['--table', '2', '--cell-max', '2'], /requires --table and --detail text/],
    [['--table', '2', '--detail', 'text', '--cell-max', '201'], /<= 200/],
  ]) {
    context.ARGS = ['tables', ...args];
    assert.throws(() => context.commandTables(), message);
  }
});

test('Set-cell reports verified readback or an explicit unverified write without a reusable guard', () => {
  const original = { readUtf8: context.readUtf8, tableFingerprint: context.tableFingerprint };
  let result, value, mode, fingerprints;
  const range = { End: 10, get Text() { return value + '\r\x07'; }, set Text(input) {
    value = mode === 'mismatch' ? 'unexpected' : input.replace(/\r\n|\n/g, '\r');
    if (mode === 'write-error') throw new Error('Write interrupted');
  } };
  const table = { Rows: { Count: 1 }, Columns: { Count: 1 }, Cell: () => ({ Range: range }) };
  const tables = () => table; tables.Count = 1;
  const doc = { Name: 'source.docx', Path: 'C:\\test', FullName: 'C:\\test\\source.docx', Tables: tables };
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: doc });
  context.readUtf8 = () => '中文\r\nsecond\n';
  context.tableFingerprint = () => {
    fingerprints++;
    if (fingerprints > 1 && mode === 'fingerprint-error') throw new Error('Format unreadable');
    return fingerprints === 1 ? 'before' : 'after';
  };
  context.emit = output => { result = JSON.parse(output); };
  try {
    for (mode of ['success', 'mismatch', 'fingerprint-error', 'write-error']) {
      value = 'old'; fingerprints = 0; result = null; context.lastError = '';
      context.ARGS = ['set-cell', '--table', '1', '--row', '1', '--col', '1', '--input', 'input.txt',
        '--expect-path', doc.FullName, '--expect-table-fingerprint', 'before', '--yes'];
      if (mode === 'success') {
        context.commandSetCell();
        assert.equal(result.verified, true);
        assert.equal(result.applied, true);
        assert.equal(result.readback.matches_requested, true);
        assert.equal(result.readback.text_hash, context.textHash('中文\nsecond\n'));
        assert.equal(result.fingerprint, 'after');
        assert.equal(result.document.path, doc.FullName);
        assert.equal(fingerprints, 2, 'Reuse the existing post-write fingerprint');
      } else {
        assert.throws(() => context.commandSetCell());
        assert.equal(result.ok, false);
        assert.equal(result.verified, false);
        assert.equal(result.fingerprint, null);
        assert.equal(result.inspection_complete, false);
        assert.ok(result.errors.length);
        assert.equal(result.applied, mode === 'write-error' ? null : true);
        if (mode === 'mismatch') assert.equal(result.readback.matches_requested, false);
      }
    }
  } finally { Object.assign(context, original); }
});

test('Literal search keeps story offsets and restores Find settings on success and partial failure', () => {
  const read = context.readUtf8;
  const settings = { Text: 'user query', MatchCase: false, MatchWholeWord: true, MatchWildcards: true,
    MatchSoundsLike: false, MatchAllWordForms: false, Forward: false, Wrap: 1, Format: true,
    MatchByte: true, MatchFuzzy: false, MatchPrefix: true, MatchSuffix: true, IgnoreSpace: true, IgnorePunct: true };
  const initial = { ...settings };
  let result, query = 'needle', searches = 0, failAfterFirst = false;
  function range(text, start = 0, end = text.length) {
    const value = { Start: start, End: end, SetRange(a, b) { this.Start = a; this.End = b; },
      get Text() { return text.slice(this.Start, this.End); }, get Duplicate() { return range(text, this.Start, this.End); } };
    const finder = { Execute() {
      searches++;
      if (failAfterFirst && searches > 1) throw new Error('Find interrupted');
      assert.equal(settings.Wrap, 0); assert.equal(settings.Format, false);
      assert.equal(settings.MatchWildcards, false); assert.equal(settings.MatchCase, true);
      const literal = settings.Text.replace(/\^\^/g, '^');
      const found = text.indexOf(literal, value.Start);
      if (found < 0 || found + literal.length > value.End) return false;
      value.Start = found; value.End = found + literal.length; return true;
    } };
    for (const key of Object.keys(settings)) Object.defineProperty(finder, key, { get: () => settings[key], set: v => { settings[key] = v; } });
    value.Find = finder; return value;
  }
  const doc = { Name: 'query.docx', FullName: 'C:\\test\\query.docx', Content: range('A needle B needle C ^p\r'),
    Footnotes: { Count: 0 }, Endnotes: { Count: 1 }, StoryRanges: () => range('note needle\r') };
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: doc });
  context.emit = output => { result = JSON.parse(output); };
  context.readUtf8 = () => query;
  try {
    context.ARGS = ['find-text', '--input', 'query.txt', '--max', '1', '--context', '2'];
    context.commandFindText();
    assert.equal(result.matches[0].start, 2); assert.equal(result.matches[0].end, 8);
    assert.equal(result.matches[0].context, 'A needle B');
    assert.equal(result.has_more, true); assert.equal(result.next_from, 8);
    assert.deepEqual(settings, initial);
    context.ARGS.push('--from', '8');
    context.commandFindText();
    assert.equal(result.matches[0].start, 11); assert.equal(result.has_more, false);
    context.ARGS = ['find-text', '--input', 'query.txt', '--story', 'endnotes'];
    context.commandFindText();
    assert.equal(result.story_type, 3); assert.equal(result.matches[0].start, 5);
    context.ARGS = ['find-text', '--input', 'query.txt', '--story', 'footnotes'];
    context.commandFindText();
    assert.equal(result.story_available, false); assert.equal(result.returned, 0);
    query = '^p'; context.ARGS = ['find-text', '--input', 'query.txt'];
    context.commandFindText();
    assert.equal(result.matches[0].text, '^p'); assert.equal(result.matches[0].start, 20);
    query = 'needle'; searches = 0; failAfterFirst = true; context.lastError = '';
    assert.throws(() => context.commandFindText());
    assert.equal(result.ok, false); assert.equal(result.returned, 1);
    assert.equal(result.has_more, null); assert.equal(result.next_from, null);
    assert.equal(result.inspection_complete, false); assert.match(result.read_errors[0], /interrupted/);
    assert.deepEqual(settings, initial, 'Failed search must restore user Find settings');
  } finally { context.readUtf8 = read; }
});

test('Table creation reports partial writes and never loses them to post-check errors', () => {
  const original = { getWord: context.getWord, readUtf8: context.readUtf8, tableFingerprint: context.tableFingerprint };
  let mode, created, values, result;
  const table = { Rows: { Count: 1 }, Columns: { Count: 2 }, Cell: (row, col) => ({ Range: {
    End: 10,
    get Text() { return values[col - 1] + '\r\x07'; },
    set Text(text) {
      values[col - 1] = mode === 'mismatch' ? 'unexpected' : text;
      if (mode === 'fill-error' && col === 2) throw new Error('Fill interrupted');
    },
  } }) };
  const tables = { Add() {
    created++;
    if (mode === 'create-error') throw new Error('Creation interrupted');
    return table;
  }, get Count() {
    if (mode === 'fill-error') throw new Error('Table count unavailable');
    return created;
  } };
  const doc = { Name: 'source.docx', Path: 'C:\\test', FullName: 'C:\\test\\source.docx',
    Tables: tables, Content: { End: 1 }, Range: () => ({}) };
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: doc });
  context.readUtf8 = () => 'A\tB';
  context.tableFingerprint = () => {
    if (mode === 'fill-error' || mode === 'postcheck-error') throw new Error('Post-check unavailable');
    return 'after';
  };
  context.emit = output => { result = JSON.parse(output); };
  try {
    for (mode of ['success', 'fill-error', 'postcheck-error', 'create-error', 'mismatch']) {
      created = 0; values = ['', '']; result = null; context.lastError = '';
      context.ARGS = ['create-table', '--rows', '1', '--cols', '2', '--input', 'table.tsv',
        '--at', 'end', '--expect-path', doc.FullName, '--yes'];
      if (mode === 'success') {
        context.commandCreateTable();
        assert.equal(result.verified, true);
        assert.equal(result.fingerprint, 'after');
        assert.deepEqual(values, ['A', 'B']);
      } else {
        assert.throws(() => context.commandCreateTable(), /Exit 3/);
        assert.equal(result.ok, false);
        assert.equal(result.verified, false);
        assert.equal(result.inspection_complete, false);
        assert.equal(result.fingerprint, null);
        if (mode === 'fill-error') {
          assert.equal(result.fill_complete, false);
          assert.equal(result.table_count, null);
          for (const error of ['Fill interrupted', 'Table count unavailable', 'Post-check unavailable']) {
            assert.ok(result.errors.some(value => value.includes(error)), error);
          }
        }
        if (mode === 'mismatch') assert.ok(result.errors.some(value => value.includes('readback')));
      }
      assert.equal(created, 1, 'A failed creation must not be retried or silently deleted');
      assert.equal(result.applied, mode === 'create-error' ? null : true);
      assert.equal(result.document.path, doc.FullName);
    }
  } finally { Object.assign(context, original); }
});

test('Border results retain write and rollback failures when final inspection also fails', () => {
  const original = { getWord: context.getWord, tableFingerprint: context.tableFingerprint };
  let mode, result, inspections, colorWrites, color;
  const border = { LineStyle: 1, LineWidth: 4, get Color() { return color; }, set Color(value) {
    color = value;
    if (mode !== 'postcheck-only') throw new Error(++colorWrites === 1 ? 'Write failed' : 'Rollback failed');
  } };
  const borders = () => border;
  const cells = () => ({ Borders: borders }); cells.Count = 1;
  const table = { Borders: borders, Range: { Cells: cells } };
  const tables = () => table; tables.Count = 1;
  const doc = { Name: 'source.docx', Path: 'C:\\test', FullName: 'C:\\test\\source.docx', Tables: tables };
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: doc });
  context.tableFingerprint = () => {
    if (++inspections === 1) return 'before';
    if (mode !== 'write-only') throw new Error('Post-check unavailable');
    return 'after';
  };
  context.emit = output => { result = JSON.parse(output); };
  try {
    for (const command of ['set-table-borders', 'set-cell-borders', 'normalize-table-borders']) {
      for (mode of ['write-and-postcheck', 'write-only', 'postcheck-only']) {
        inspections = 0; colorWrites = 0; color = 0; result = null; context.lastError = '';
        context.ARGS = [command, '--table', '1', '--cell', '1', '--edges', 'top',
          '--expect-path', doc.FullName, '--expect-table-fingerprint', 'before', '--yes'];
        assert.throws(() => command === 'normalize-table-borders' ? context.commandNormalizeTableBorders()
          : context.commandSetBorders(command === 'set-table-borders' ? 'table' : 'cell'), /Exit 3/);
        assert.equal(result.ok, false);
        assert.equal(result.verified, false);
        assert.equal(result.inspection_complete, false);
        assert.equal(result.fingerprint, null, 'Even a readable post-check cannot authorize retry after failed writes');
        assert.equal(result.applied, mode === 'postcheck-only' ? true : null);
        assert.equal(result.document.path, doc.FullName);
        assert.equal(inspections, 2, 'Keep both live pre-check and post-check');
        if (mode !== 'postcheck-only') {
          assert.ok(result.failures.some(value => value.includes('Write failed')));
          assert.equal(result.failure_count, result.failures.length);
          if (command !== 'normalize-table-borders') {
            assert.ok(result.rollback_failures.some(value => value.includes('Rollback failed')));
            assert.equal(result.rolled_back, false);
          }
        }
        if (mode !== 'write-only') assert.ok(result.errors.some(value => value.includes('Post-check unavailable')));
      }
    }
  } finally { Object.assign(context, original); }
});

test('XML comparison is bounded, optional, and cannot replace a live mutation guard', () => {
  const original = { getWord: context.getWord, tableFingerprint: context.tableFingerprint,
    compareTableXml: context.compareTableXml };
  let result, compareCalls = 0, fail = false;
  const table = { Rows: { Count: 1 }, Columns: { Count: 1 }, Cell: () => ({ Range: { Text: 'A\r\x07' } }) };
  const tables = () => table; tables.Count = 1;
  const doc = { Name: 'source.docx', Path: 'C:\\test', FullName: 'C:\\test\\source.docx', Tables: tables };
  context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: doc });
  context.tableFingerprint = () => 'abc12345';
  context.compareTableXml = () => {
    compareCalls++;
    return { candidate: fail ? null : 'xml-v1:12345678', stable: !fail, sourceChars: 100,
      normalizedChars: 80, elapsedMs: 1, errors: fail ? ['XML unavailable'] : [] };
  };
  context.emit = output => { result = JSON.parse(output); };
  try {
    context.ARGS = ['tables', '--table', '1'];
    context.commandTables();
    assert.equal(compareCalls, 0);
    assert.equal('xml_comparison' in result.tables[0], false);
    context.ARGS.push('--compare-xml', '--expect-path', doc.FullName);
    context.commandTables();
    assert.equal(result.tables[0].fingerprint, 'abc12345');
    assert.equal(result.tables[0].xml_comparison.candidate_hash, 'xml-v1:12345678');
    assert.equal(result.tables[0].xml_comparison.diagnostic_only, true);
    assert.equal(result.document.path, doc.FullName);
    context.ARGS = ['set-cell', '--expect-table-fingerprint', 'xml-v1:12345678', '--allow-unverified-target'];
    assert.throws(() => context.requireFingerprint('abc12345', '--expect-table-fingerprint', '--allow-unverified-target'), /cannot authorize/);
    for (const args of [['tables', '--compare-xml'], ['tables', '--table', '1', '--detail', 'text', '--compare-xml']]) {
      context.ARGS = args;
      assert.throws(() => context.commandTables(), /single table.*full/);
    }
    context.ARGS = ['tables', '--table', '1', '--compare-xml', '--expect-path', doc.FullName];
    fail = true; context.lastError = '';
    assert.throws(() => context.commandTables(), /Exit 3/);
    assert.equal(result.ok, false);
    assert.equal(result.tables[0].fingerprint, null);
    assert.equal(result.tables[0].inspection_complete, false);
    assert.equal(result.tables[0].xml_comparison.candidate_hash, null);
    assert.deepEqual(result.tables[0].xml_comparison.errors, ['XML unavailable']);
  } finally { Object.assign(context, original); }
});

function outputGuardRuntime() {
  const files = new Map([['C:\\test\\source.docx', 'original document']]);
  const folders = new Set(['C:\\test']);
  const runtime = { WScript: { Arguments: [], Echo() {}, Sleep() {},
    StdErr: { WriteLine(value) { runtime.lastError = value; } },
    Quit(code) { throw new Error(runtime.lastError || 'Exit ' + code); } } };
  const disk = {
    GetAbsolutePathName: value => path.win32.resolve(value),
    GetParentFolderName: value => path.win32.dirname(value),
    GetExtensionName: value => path.win32.extname(value).slice(1),
    BuildPath: (parent, name) => path.win32.join(parent, name),
    GetTempName: () => 'guard.tmp',
    FileExists: value => files.has(value), FolderExists: value => folders.has(value),
    GetFile(value) { if (!files.has(value)) throw new Error('Missing file'); return { Size: files.get(value).length, ShortPath: value }; },
    CreateFolder(value) { assert.equal(folders.has(value), false); folders.add(value); },
    GetFolder(value) { return { Files: { Count: [...files.keys()].filter(p => path.win32.dirname(p) === value).length }, SubFolders: { Count: 0 } }; },
    DeleteFolder(value) { assert.equal(disk.GetFolder(value).Files.Count, 0); folders.delete(value); },
    DeleteFile(value) { files.delete(value); },
    CopyFile(from, to, overwrite) { assert.ok(files.has(from)); assert.equal(overwrite, false); assert.equal(files.has(to), false); files.set(to, files.get(from)); },
    MoveFile(from, to) {
      if (runtime.moveFailure) runtime.moveFailure(from, to);
      if (files.has(to) || folders.has(to)) throw new Error('Destination exists');
      assert.ok(files.has(from)); files.set(to, files.get(from)); files.delete(from);
    },
    OpenTextFile(value) { return { Read: count => files.get(value).slice(0, count), Close() {} }; },
  };
  runtime.ActiveXObject = function(name) { assert.equal(name, 'Scripting.FileSystemObject'); return disk; };
  vm.createContext(runtime);
  vm.runInContext(source, runtime);
  runtime.emit = text => { runtime.result = JSON.parse(text); };
  const doc = { FullName: 'C:\\test\\source.docx', Path: 'C:\\test', Saved: false, Save() {}, Close(mode) { runtime.closed = true; runtime.closeMode = mode; } };
  const word = { Documents: { Count: 1 }, ActiveDocument: doc, BackgroundSavingStatus: 0 };
  runtime.getWord = () => word;
  return { runtime, files, folders, disk, doc, word };
}

test('Save and close keep the document open when saving cannot be confirmed', () => {
  for (const command of ['save-active', 'close-active']) {
    for (const state of ['cancelled', 'unreadable', 'missing-file', 'throwing-save', 'background-error', 'pending']) {
      const { runtime: r, doc, files, word } = outputGuardRuntime();
      r.ARGS = [command, '--yes', '--save', '--expect-path', doc.FullName];
      if (state === 'unreadable') Object.defineProperty(doc, 'Saved', { get() { throw new Error('Saved unavailable'); } });
      if (state === 'missing-file') { doc.Save = () => { doc.Saved = true; }; files.clear(); }
      if (state === 'throwing-save') doc.Save = () => { throw new Error('Save failed'); };
      if (state === 'background-error') { doc.Saved = true; Object.defineProperty(word, 'BackgroundSavingStatus', { get() { throw new Error('Queue unavailable'); } }); }
      if (state === 'pending') {
        let now = 0; r.Date = class { getTime() { return now; } };
        r.WScript.Sleep = ms => { now += ms; }; doc.Saved = true; word.BackgroundSavingStatus = 1;
      }
      assert.throws(() => command === 'save-active' ? r.commandSaveActive() : r.commandCloseActive(), /save|saving|Saved|Queue|missing or empty/i);
      assert.equal(r.closed, undefined, state + ' closed the document');
      assert.equal(r.result, undefined, state + ' reported success');
    }
  }
  const { runtime: r, doc } = outputGuardRuntime();
  doc.Save = () => { doc.Saved = true; };
  r.ARGS = ['close-active', '--save', '--yes', '--expect-path', doc.FullName];
  r.commandCloseActive();
  assert.equal(r.closed, true);
  assert.equal(r.closeMode, -1, 'Close must retain save semantics for edits after the save check');
  assert.equal(r.result.saved, 'true', 'Preserve the existing result type');
});

test('Backup and PDF publication preserve old outputs through failures and reject unsafe paths', () => {
  for (const command of ['save-copy', 'export-pdf']) {
    for (const phase of ['generation', 'empty', 'old-move', 'publish', 'restore', 'success']) {
      const { runtime: r, doc, files, folders } = outputGuardRuntime();
      const target = command === 'save-copy' ? 'C:\\test\\backup.docx' : 'C:\\test\\preview.pdf';
      files.set(target, 'old output');
      const generate = output => {
        assert.equal(files.get(target), 'old output', 'Old output was removed before generation');
        files.set(output, phase === 'empty' ? '' : (command === 'export-pdf' ? '%PDF-new output' : 'new document'));
        if (phase === 'generation') throw new Error('Generation failed');
      };
      doc.SaveCopyAs = generate; doc.ExportAsFixedFormat = generate;
      r.moveFailure = (from, to) => {
        if (phase === 'old-move' && from === target) throw new Error('Old output locked');
        if ((phase === 'publish' || phase === 'restore') && to === target && path.win32.basename(from).startsWith('new.')) throw new Error('Publication failed');
        if (phase === 'restore' && to === target && path.win32.basename(from).startsWith('previous.')) throw new Error('Recovery blocked');
      };
      r.ARGS = [command, '--path', target, '--overwrite', '--yes', '--expect-path', doc.FullName];
      const invoke = () => command === 'save-copy' ? r.commandSaveCopy() : r.commandExportPdf();
      if (phase === 'success') {
        invoke(); assert.equal(r.result.ok, true); assert.notEqual(files.get(target), 'old output');
      } else {
        assert.throws(invoke, /failed|empty|saved|recovery|publication/i);
        assert.equal(r.result, undefined);
        if (phase === 'restore') {
          const recovery = [...files.entries()].find(([name, value]) => name !== target && value === 'old output');
          assert.ok(recovery, 'The prior output must survive even if rollback is blocked');
          assert.ok(r.lastError.includes(recovery[0]), 'Report the precise recovery file');
        } else { assert.equal(files.get(target), 'old output'); }
      }
      assert.equal(files.get(doc.FullName), 'original document');
      assert.equal(doc.Saved, false);
      if (phase !== 'restore') assert.equal(folders.size, 1, 'Staging folder leaked');
    }
  }
  for (const target of ['C:\\test\\source.docx', 'C:\\test\\source.docx:preview.pdf', 'C:\\test\\*.pdf', 'C:\\test\\preview.pdf.']) {
    const { runtime: r, doc, files } = outputGuardRuntime();
    doc.ExportAsFixedFormat = () => { throw new Error('Unexpected Word export'); };
    r.ARGS = ['export-pdf', '--path', target, '--overwrite', '--yes', '--expect-path', doc.FullName];
    assert.throws(() => r.commandExportPdf(), /\.pdf|literal|path/i);
    assert.equal(files.get(doc.FullName), 'original document');
  }
  const { runtime: r, doc, files, disk } = outputGuardRuntime();
  doc.FullName = 'C:\\test\\source.pdf'; files.set(doc.FullName, 'source PDF');
  r.ARGS = ['export-pdf', '--path', 'C:\\test\\SOURCE.pdf', '--overwrite', '--yes', '--expect-path', doc.FullName];
  assert.throws(() => r.commandExportPdf(), /differ/);
  const alias = 'C:\\test\\SOURCE~1.pdf'; files.set(alias, 'source PDF');
  disk.GetFile = () => ({ Size: 10, ShortPath: alias });
  r.ARGS = ['save-copy', '--path', alias, '--overwrite', '--yes', '--expect-path', doc.FullName];
  assert.throws(() => r.commandSaveCopy(), /differ/);
  const race = outputGuardRuntime(), target = 'C:\\test\\race.pdf';
  race.doc.ExportAsFixedFormat = output => { race.files.set(output, '%PDF-generated'); race.files.set(target, 'concurrent output'); };
  race.runtime.ARGS = ['export-pdf', '--path', target, '--yes', '--expect-path', race.doc.FullName];
  assert.throws(() => race.runtime.commandExportPdf(), /overwrite/);
  assert.equal(race.files.get(target), 'concurrent output', 'No overwrite approval existed for the concurrent output');
});

test('Smoke preserves old output until a completed save is closed and published', () => {
  for (const phase of ['startup', 'generation', 'save-cancelled', 'empty', 'publish', 'success']) {
    const { runtime: r, disk, doc, word, files, folders } = outputGuardRuntime();
    const target = 'C:\\test\\smoke.docx'; files.set(target, 'previous smoke document');
    let savedPath, saveCalls = 0, closed = false, quit = false;
    const table = { Cell() { return {}; } };
    doc.Content = { End: 5, Text: '' }; doc.Range = () => ({});
    doc.Tables = { Count: 1, Add() { return table; } }; doc.OMaths = { Count: 1 };
    doc.Close = () => { closed = true; }; word.Quit = () => { quit = true; };
    word.Documents.Add = () => doc;
    r.ActiveXObject = function(name) {
      if (name === 'Word.Application') {
        if (phase === 'startup') throw new Error('Word startup failed');
        return word;
      }
      assert.equal(name, 'Scripting.FileSystemObject'); return disk;
    };
    r.setCellText = () => {}; r.setBorderCollection = () => [];
    r.buildEquationInRange = () => {
      assert.equal(files.get(target), 'previous smoke document', 'Old file was deleted before generation finished');
      if (phase === 'generation') throw new Error('Equation generation failed');
    };
    doc.SaveAs2 = output => {
      saveCalls++; savedPath = output;
      files.set(output, phase === 'empty' ? '' : 'completed smoke document');
      doc.FullName = output; doc.Path = path.win32.dirname(output); doc.Saved = phase !== 'save-cancelled';
    };
    doc.SaveAs = () => { throw new Error('Unexpected legacy save fallback'); };
    doc.Save = () => { throw new Error('Smoke must not trigger an additional save'); };
    r.moveFailure = (from, to) => {
      if (from === savedPath) {
        assert.equal(closed, true, 'Close the saved document before publishing its file');
        if (phase === 'publish' && to === target) throw new Error('Publication failed');
      }
    };
    r.ARGS = ['smoke', '--path', target, '--overwrite', '--yes'];
    if (phase === 'success') {
      r.commandSmoke();
      assert.equal(r.result.ok, true); assert.equal(r.result.path, target);
      assert.equal(files.get(target), 'completed smoke document');
      assert.equal(saveCalls, 1); assert.notEqual(savedPath, target);
    } else {
      assert.throws(() => r.commandSmoke(), /smoke test failed/);
      assert.equal(files.get(target), 'previous smoke document');
      assert.equal(r.result, undefined);
    }
    assert.equal(files.get('C:\\test\\source.docx'), 'original document');
    assert.equal(folders.size, 1, 'Smoke staging folder leaked');
    if (phase !== 'startup') assert.equal(quit, true, 'Owned Word instance was not closed');
  }
});

test('Swap and shading retain applied state and rollback evidence when final inspection fails', () => {
  const original = { getWord: context.getWord, tableFingerprint: context.tableFingerprint, emit: context.emit };
  try {
    for (const command of ['swap-cell-text', 'set-cell-shading']) {
      for (const mode of ['postcheck-null', 'postcheck-throw', 'success', 'rollback-success', 'rollback-throws', 'rollback-silent', 'readback-unavailable']) {
        let writes = 0, inspections = 0, result;
        const bodies = ['LEFT\r', 'RIGHT\x0b'], initial = bodies.slice();
        let color = 17, texture = 25;
        function write(perform) {
          writes++;
          if (mode === 'rollback-silent' && writes >= 3) return;
          perform();
          if (mode.indexOf('rollback-') === 0 && writes === 2) throw new Error('Write interrupted');
          if (mode === 'rollback-throws' && writes >= 3) throw new Error('Restore interrupted');
        }
        const shading = {
          get Texture() { return mode === 'readback-unavailable' && writes === 2 ? NaN : texture; }, set Texture(value) { write(() => { texture = value; }); },
          get BackgroundPatternColor() { return color; }, set BackgroundPatternColor(value) { write(() => { color = value; }); }
        };
        const cells = n => ({ Shading: shading, get Range() { return {
          Start: n * 20, End: n * 20 + 10, get Text() {
            if (mode === 'readback-unavailable' && writes === 2) throw new Error('Cell text unreadable');
            return bodies[n - 1] + '\r\x07';
          },
          set Text(value) { write(() => { bodies[n - 1] = value; }); }
        }; } }); cells.Count = 2;
        const table = { Range: { Cells: cells } }, tables = () => table; tables.Count = 1;
        const doc = { Name: 'source.docx', Path: 'C:\\test', FullName: 'C:\\test\\source.docx', TrackRevisions: false, Tables: tables };
        context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: doc });
        context.tableFingerprint = (table, errors) => {
          if (++inspections === 1) return 'before';
          if (mode === 'postcheck-null') { if (errors) errors.push('Final inspection incomplete'); return null; }
          if (mode === 'postcheck-throw' || mode === 'rollback-throws') throw new Error('Final inspection unavailable');
          return 'after';
        };
        context.emit = output => { result = JSON.parse(output); }; context.lastError = '';
        context.ARGS = [command, '--table', '1', '--from-cell', '1', '--to-cell', '2', '--cell', '1', '--color', 'FF0000',
          '--expect-path', doc.FullName, '--expect-table-fingerprint', 'before', '--yes'];
        const run = () => command === 'swap-cell-text' ? context.commandSwapCellText() : context.commandSetCellShading();
        if (mode === 'success') {
          run();
          assert.equal(result.ok, true); assert.equal(result.verified, true); assert.equal(result.applied, true);
          assert.equal(result.readback.matches_requested, true); assert.equal(result.fingerprint, 'after');
          assert.equal(result.failure_count, 0); assert.deepEqual(result.errors, []);
          assert.equal(writes, 2, 'A verified result must not repeat the successful write');
        } else {
          assert.throws(run, /Exit 3/);
          assert.equal(result.ok, false); assert.equal(result.verified, false); assert.equal(result.inspection_complete, false);
          assert.equal(result.fingerprint, null, 'Never offer an incomplete or failed operation as a reusable guard');
          if (mode.indexOf('postcheck-') === 0) {
            assert.equal(result.applied, true); assert.equal(result.rolled_back, false);
            assert.equal(result.readback.matches_requested, true); assert.equal(result.failure_count, 0);
            assert.ok(result.errors.some(error => error.includes('Final inspection')));
            assert.equal(writes, 2, 'A final-inspection failure must not replay or undo successful writes');
          } else {
            const restored = mode === 'rollback-success' || mode === 'readback-unavailable';
            assert.ok(result.failures.some(error => mode === 'readback-unavailable' ? /unreadable|non-finite/.test(error) : error.includes('Write interrupted')));
            assert.equal(result.applied, restored ? false : null);
            assert.equal(result.rolled_back, restored);
            assert.equal(result.rollback_failures.length === 0, restored);
            if (mode === 'rollback-silent') assert.ok(result.rollback_failures.some(error => error.includes('readback')));
            if (mode === 'rollback-throws') {
              assert.ok(result.rollback_failures.some(error => error.includes('Restore interrupted')));
              assert.ok(result.errors.some(error => error.includes('Final inspection unavailable')));
            }
          }
        }
        assert.equal(result.document.path, doc.FullName);
        assert.equal(inspections, 2, 'Reuse the existing pre-write and post-write full inspections');
        if (mode.indexOf('postcheck-') === 0 || mode === 'success') {
          if (command === 'swap-cell-text') assert.deepEqual(bodies, [initial[1], initial[0]]);
          else assert.deepEqual([texture, color], [0, 255]);
        }
        if (mode === 'rollback-success' || mode === 'readback-unavailable') {
          if (command === 'swap-cell-text') assert.deepEqual(bodies, initial);
          else assert.deepEqual([texture, color], [25, 17]);
        }
      }
    }
  } finally { Object.assign(context, original); }
});

test('Row and column results retain completed or uncertain writes through failed dimension and fingerprint reads', () => {
  const original = { getWord: context.getWord, tableFingerprint: context.tableFingerprint, emit: context.emit };
  const cases = [
    ['insert-row', 'commandInsertRow', 'rows', 1, 4, 3],
    ['delete-row', 'commandDeleteRow', 'rows', -1, 2, 3],
    ['insert-column', 'commandInsertColumn', 'cols', 1, 3, 4],
    ['delete-column', 'commandDeleteColumn', 'cols', -1, 3, 2]
  ];
  try {
    for (const [command, method, axis, delta, expectedRows, expectedCols] of cases) {
      for (const mode of ['postcheck-null', 'postcheck-throw', 'success', 'row-unreadable', 'col-unreadable', 'dimension-mismatch', 'invalid-count', 'write-throw', 'write-and-postcheck']) {
        const dimensions = { rows: 3, cols: 3 };
        let writes = 0, inspections = 0, result;
        function mutate() {
          writes++;
          if (mode !== 'dimension-mismatch') dimensions[axis] += delta;
          if (mode === 'write-throw' || mode === 'write-and-postcheck') throw new Error('Structural write interrupted');
        }
        function collection(name) {
          const items = index => { assert.equal(index, 2); return { Delete: mutate }; };
          items.Add = mutate;
          Object.defineProperty(items, 'Count', { get() {
            if (writes && (mode === 'row-unreadable' && name === 'rows' || mode === 'col-unreadable' && name === 'cols')) throw new Error(name + ' unavailable');
            return writes && mode === 'invalid-count' && name === 'rows' ? NaN : dimensions[name];
          } });
          return items;
        }
        const table = { Rows: collection('rows'), Columns: collection('cols') }, tables = () => table; tables.Count = 1;
        const doc = { Name: 'source.docx', Path: 'C:\\test', FullName: 'C:\\test\\source.docx', TrackRevisions: false, Tables: tables };
        context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: doc });
        context.tableFingerprint = (table, errors) => {
          if (++inspections === 1) return 'before';
          if (mode === 'postcheck-null') { if (errors) errors.push('Final inspection incomplete'); return null; }
          if (mode === 'postcheck-throw' || mode === 'write-and-postcheck') throw new Error('Final inspection unavailable');
          return 'after';
        };
        context.emit = text => { result = JSON.parse(text); }; context.lastError = '';
        context.ARGS = [command, '--table', '1', '--before', '2', '--row', '2', '--col', '2',
          '--expect-path', doc.FullName, '--expect-table-fingerprint', 'before', '--yes'];
        if (mode === 'success') {
          context[method]();
          assert.equal(result.ok, true); assert.equal(result.verified, true); assert.equal(result.inspection_complete, true);
          assert.equal(result.fingerprint, 'after'); assert.deepEqual(result.errors, []);
          assert.equal(result.rows, expectedRows); assert.equal(result.cols, expectedCols);
        } else {
          assert.throws(() => context[method](), /Exit 3/);
          assert.equal(result.ok, false); assert.equal(result.verified, false); assert.equal(result.inspection_complete, false);
          assert.equal(result.fingerprint, null); assert.ok(result.errors.length);
          if (mode === 'write-throw' || mode === 'write-and-postcheck') assert.ok(result.errors.some(error => error.includes('Structural write interrupted')));
          if (mode === 'postcheck-throw' || mode === 'write-and-postcheck') assert.ok(result.errors.some(error => error.includes('Final inspection unavailable')));
          if (mode === 'dimension-mismatch') assert.equal(result.error, 'dimension readback mismatch');
          if (mode === 'row-unreadable' || mode === 'invalid-count') assert.equal(result.rows, null);
          if (mode === 'col-unreadable') assert.equal(result.cols, null);
        }
        assert.equal(result.applied, mode === 'write-throw' || mode === 'write-and-postcheck' ? null : true);
        assert.equal(result.document.path, doc.FullName);
        assert.equal(result.readback.expected_rows, expectedRows); assert.equal(result.readback.expected_cols, expectedCols);
        assert.equal(result.readback.matches_requested, ['row-unreadable', 'col-unreadable', 'invalid-count'].includes(mode) ? null : mode !== 'dimension-mismatch');
        assert.equal(writes, 1, 'Structural writes must never be replayed or automatically reversed');
        assert.equal(inspections, 2, 'Retain one live guard check and one final inspection');
      }
    }
  } finally { Object.assign(context, original); }
});

test('Object deletion blocks tracked table writes and preserves uncertain or incomplete removal results', () => {
  const original = { getWord: context.getWord, tableFingerprint: context.tableFingerprint, equationFingerprint: context.equationFingerprint, emit: context.emit };
  try {
    for (const kind of ['table', 'equation']) {
      for (const mode of ['success', 'tracked-blocked', 'tracked-retained', 'no-removal', 'readback-throw', 'invalid-count', 'write-throw', 'write-and-readback', 'zero-return',
        'display-success', 'display-prepare-throw', 'display-prepare-silent', 'display-delete-throw']) {
        if (mode === 'tracked-blocked' && kind !== 'table' || (mode === 'zero-return' || mode.startsWith('display-')) && kind !== 'equation') continue;
        let count = 3, writes = 0, inspections = 0, preparationWrites = 0, equationType = mode.startsWith('display-') ? 0 : 1, result;
        const preparationFailed = mode.startsWith('display-prepare-');
        const writeThrew = ['write-throw', 'write-and-readback', 'display-delete-throw'].includes(mode);
        const success = mode === 'success' || mode === 'display-success';
        function remove() {
          writes++;
          if (!['tracked-retained', 'no-removal', 'zero-return'].includes(mode)) count--;
          if (writeThrew) throw new Error('Deletion interrupted');
          return mode === 'zero-return' ? 0 : 1;
        }
        const target = kind === 'table' ? { Delete: remove } : { Range: { Delete: remove },
          get Type() { return equationType; }, set Type(value) {
            preparationWrites++;
            if (mode !== 'display-prepare-silent') equationType = value;
            if (mode === 'display-prepare-throw') throw new Error('Inline preparation interrupted');
          } };
        const collection = index => { assert.equal(index, 2); return target; };
        Object.defineProperty(collection, 'Count', { get() {
          if (writes && (mode === 'readback-throw' || mode === 'write-and-readback')) throw new Error('Collection unreadable');
          return writes && mode === 'invalid-count' ? NaN : count;
        } });
        const tracked = mode.startsWith('tracked-');
        const doc = { Name: 'source.docx', Path: 'C:\\test', FullName: 'C:\\test\\source.docx', TrackRevisions: tracked,
          [kind === 'table' ? 'Tables' : 'OMaths']: collection };
        context.getWord = () => ({ Documents: { Count: 1 }, ActiveDocument: doc });
        context.tableFingerprint = context.equationFingerprint = () => { inspections++; return 'before'; };
        context.emit = text => { result = JSON.parse(text); }; context.lastError = '';
        context.ARGS = ['delete-' + kind, kind === 'table' ? '--table' : '--index', '2',
          '--expect-' + kind + '-fingerprint', 'before', '--expect-path', doc.FullName, '--yes'];
        if (mode === 'tracked-retained' && kind === 'table') context.ARGS.push('--allow-track-changes');
        const run = () => kind === 'table' ? context.commandDeleteTable() : context.commandDeleteEquation();
        if (mode === 'tracked-blocked') {
          assert.throws(run, /Track Changes/);
          assert.equal(writes, 0); assert.equal(result, undefined); assert.equal(count, 3);
          continue;
        }
        if (success) run(); else assert.throws(run, /Exit 3/);
        const unknownCount = ['readback-throw', 'invalid-count', 'write-and-readback'].includes(mode);
        assert.equal(result.ok, success); assert.equal(result.verified, success);
        assert.equal(result.inspection_complete, success); assert.equal(result.fingerprint, null);
        assert.equal(result.applied, writeThrew || preparationFailed ? null : mode !== 'zero-return');
        assert.equal(result[kind === 'table' ? 'remaining_tables' : 'remaining_equations'], unknownCount ? null : count);
        assert.equal(result.readback.expected_remaining, 2);
        assert.equal(result.readback.matches_requested, unknownCount ? null : count === 2);
        assert.equal(result.document.path, doc.FullName); assert.equal(result.track_revisions, tracked);
        assert.equal(result[kind === 'table' ? 'table' : 'index'], 2);
        assert.equal(result.errors.length === 0, success);
        if (mode === 'write-and-readback') {
          assert.ok(result.errors.some(error => error.includes('Deletion interrupted')));
          assert.ok(result.errors.some(error => error.includes('Collection unreadable')));
        }
        if (kind === 'equation') {
          assert.equal(result.deleted_units, writeThrew || preparationFailed ? null : mode === 'zero-return' ? 0 : 1);
          assert.equal(result.prepared_inline, preparationFailed ? null : mode.startsWith('display-'));
        }
        assert.equal(preparationWrites, mode.startsWith('display-') ? 1 : 0);
        assert.equal(writes, preparationFailed ? 0 : 1, 'Never delete after failed preparation or replay a deletion');
        assert.equal(inspections, 1, 'Inspect the target before deletion, without fingerprinting a removed object');
        assert.equal(doc.TrackRevisions, tracked, 'Deletion must not change revision policy');
      }
    }
  } finally { Object.assign(context, original); }
});

console.log(JSON.stringify({ ok: true, pure_regression_groups: passed }));
