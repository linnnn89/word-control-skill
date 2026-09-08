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

console.log(JSON.stringify({ ok: true, pure_regression_groups: passed }));
