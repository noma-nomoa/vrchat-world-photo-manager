'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { crc32 } = require('node:zlib');
const { parseReleases, assertInside, validateRegressionReport, verifyNupkgPayload, verifyExecutionStub } = require('./build-review-candidate');
const hashes = { 'src/index.js': 'a'.repeat(64), 'package.json': 'b'.repeat(64) };
const plan = ['check:syntax', 'test', 'smoke:packaged'];
function report() { return { ok: true, status: 'passed', suite: 'all', sourceUnchanged: true,
  sourceHashes: { ...hashes }, plannedSteps: [...plan], results: plan.map(name => ({ name, ok: true, exitCode: 0, signal: null, aborted: false, timedOut: false })) }; }

test('accepts complete passed checks tied to the current full source inventory', () => assert.doesNotThrow(() => validateRegressionReport(report(), hashes, plan)));
for (const [name, mutate] of [
  ['failed run', value => { value.ok = false; }],
  ['partial suite', value => { value.suite = 'unit'; }],
  ['source changed during checks', value => { value.sourceUnchanged = false; }],
  ['stale or edited source inventory', value => { value.sourceHashes['src/index.js'] = 'c'.repeat(64); }],
  ['missing source entry', value => { delete value.sourceHashes['package.json']; }],
  ['shortened plan and matching shortened results', value => { value.plannedSteps.pop(); value.results.pop(); }],
  ['missing result', value => { value.results.pop(); }],
  ['reordered result', value => { value.results.reverse(); }],
  ['nonzero child exit hidden by success flag', value => { value.results[1].exitCode = 1; }],
  ['timeout hidden by success flag', value => { value.results[1].timedOut = true; }],
  ['termination error hidden by success flag', value => { value.results[1].terminationError = 'controlled cleanup failure'; }],
]) test(`rejects ${name}`, () => { const value = report(); mutate(value); assert.throws(() => validateRegressionReport(value, hashes, plan)); });

test('parses the exact Squirrel RELEASES hash, package name and size', () => {
  assert.deepEqual(parseReleases(`${'AB'.repeat(20)} worldshot_log-2.5.3-full.nupkg 123\r\n`), [
    { sha1: 'ab'.repeat(20), name: 'worldshot_log-2.5.3-full.nupkg', bytes: 123 },
  ]);
});
for (const [name, line] of [
  ['empty manifest', ''], ['malformed hash', `abc package.nupkg 12`],
  ['parent path', `${'a'.repeat(40)} ../package.nupkg 12`],
  ['Windows path', `${'a'.repeat(40)} ..\\package.nupkg 12`],
  ['absolute path', `${'a'.repeat(40)} C:\\package.nupkg 12`],
  ['zero size', `${'a'.repeat(40)} package.nupkg 0`],
  ['oversized integer', `${'a'.repeat(40)} package.nupkg 9007199254740992`],
  ['wrong extension', `${'a'.repeat(40)} package.exe 123`],
]) test(`rejects RELEASES ${name}`, () => assert.throws(() => parseReleases(line)));

test('output paths must remain strictly inside the generated candidate', () => {
  const root = path.resolve(os.tmpdir(), 'worldshot-path-fixture');
  assert.doesNotThrow(() => assertInside(root, path.join(root, 'build', 'file')));
  for (const target of [root, path.dirname(root), path.resolve(root, '..', 'outside')]) assert.throws(() => assertInside(root, target));
});

test('malformed installer archive is rejected without extracting anything', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-candidate-validator-'));
  const app = path.join(root, 'app'); fs.mkdirSync(app);
  fs.writeFileSync(path.join(app, 'WorldShotLog.exe'), 'fixed application bytes');
  const nupkg = path.join(root, 'invalid.nupkg'); fs.writeFileSync(nupkg, 'not a ZIP archive');
  await assert.rejects(verifyNupkgPayload(nupkg, app));
  assert.deepEqual(fs.readdirSync(root).sort(), ['app', 'invalid.nupkg']);
});

// Small stored ZIP fixtures cover content integrity independently of Forge.
function storedZip(entries) {
  const files = []; const central = []; let offset = 0;
  for (const [name, value] of entries) {
    const nameBytes = Buffer.from(name); const bytes = Buffer.from(value); const crc = crc32(bytes);
    const local = Buffer.alloc(30); local.writeUInt32LE(0x04034b50); local.writeUInt16LE(20, 4);
    local.writeUInt32LE(crc, 14); local.writeUInt32LE(bytes.length, 18); local.writeUInt32LE(bytes.length, 22); local.writeUInt16LE(nameBytes.length, 26);
    const directory = Buffer.alloc(46); directory.writeUInt32LE(0x02014b50); directory.writeUInt16LE(20, 4); directory.writeUInt16LE(20, 6);
    directory.writeUInt32LE(crc, 16); directory.writeUInt32LE(bytes.length, 20); directory.writeUInt32LE(bytes.length, 24);
    directory.writeUInt16LE(nameBytes.length, 28); directory.writeUInt32LE(offset, 42);
    files.push(local, nameBytes, bytes); central.push(directory, nameBytes); offset += local.length + nameBytes.length + bytes.length;
  }
  const directory = Buffer.concat(central); const end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50);
  end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10); end.writeUInt32LE(directory.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...files, directory, end]);
}
for (const variant of ['complete', 'unknown', 'duplicate', 'missing-application', 'missing-helper', 'changed-application', 'changed-helper']) {
  test(`installer payload ${variant === 'complete' ? 'accepts complete files and explicit helper' : `rejects ${variant}`}`, async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-payload-fixture-'));
    const app = path.join(root, 'app'); fs.mkdirSync(app);
    const original = Buffer.from('original application'); const helper = Buffer.from('standard helper');
    fs.writeFileSync(path.join(app, 'WorldShotLog.exe'), original);
    const entries = [['lib/net45/WorldShotLog.exe', original], ['lib/net45/squirrel.exe', helper]];
    if (variant === 'unknown') entries.push(['lib/net45/unknown.exe', helper]);
    if (variant === 'duplicate') entries.push(entries[0]);
    if (variant === 'missing-application') entries.shift();
    if (variant === 'missing-helper') entries.pop();
    if (variant === 'changed-application') entries[0][1] = Buffer.from('modified application');
    if (variant === 'changed-helper') entries[1][1] = Buffer.from('modified helper');
    const archive = path.join(root, 'fixture.nupkg'); fs.writeFileSync(archive, storedZip(entries));
    const validate = () => verifyNupkgPayload(archive, app, { 'squirrel.exe': {
      bytes: helper.length, sha256: createHash('sha256').update(helper).digest('hex'),
    } });
    if (variant === 'complete') { const result = await validate(); assert.equal(result.comparedFiles, 2); assert.equal(result.applicationFiles, 1); }
    else await assert.rejects(validate());
  });
}
test('Squirrel stub requires original vendor code and exact application resources', async () => {
  const pe = await import('pe-library');
  const vendorBytes = fs.readFileSync(path.resolve(__dirname, '../node_modules/electron-winstaller/vendor/StubExecutable.exe'));
  const application = pe.NtExecutable.createEmpty(true, false);
  const resources = pe.NtExecutableResource.from(application);
  resources.entries = [{ type: 10, id: 1, lang: 1033, codepage: 0, bin: Uint8Array.from([1, 2, 3, 4]).buffer }];
  resources.outputResource(application);
  const applicationBytes = Buffer.from(application.generate());
  const stub = pe.NtExecutable.from(vendorBytes); resources.outputResource(stub, false, true);
  const stubBytes = Buffer.from(stub.generate());
  const result = await verifyExecutionStub(stubBytes, vendorBytes, applicationBytes);
  assert.equal(result.comparedResourceEntries, 1); assert.equal(result.comparedCodeSections.length, 4);
  const changedCode = Buffer.from(stubBytes); changedCode[stub.getAllSections()[0].info.pointerToRawData] ^= 1;
  await assert.rejects(verifyExecutionStub(changedCode, vendorBytes, applicationBytes), /code differs/);
  resources.entries[0].bin = Uint8Array.from([1, 2, 3, 5]).buffer; resources.outputResource(stub, false, true);
  await assert.rejects(verifyExecutionStub(Buffer.from(stub.generate()), vendorBytes, applicationBytes), /resources differ/);
});
