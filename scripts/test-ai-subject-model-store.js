'use strict';

const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { pathToFileURL } = require('node:url');
const { spawnSync } = require('node:child_process');
// This same suite was first run against the frozen, original main-process functions.
const { createAiSubjectModelStore } = require(process.env.MODEL_STORE_BASELINE
  ? '../out/refactor-model-baseline/adapter.cjs' : '../src/ai-subject-model-store');
const { AI_SUBJECT_MODELS, AI_MODEL_TIERS } = require('../src/ai-subject-model-catalog');
const catalogBaseline = require('./fixtures/ai-subject-model-catalog-baseline.json');

const binary = Buffer.from('synthetic model binary, never used for inference');
const digest = bytes => createHash('sha256').update(bytes).digest('hex');
const definition = (id, extra = {}) => ({ id, tier: 'standard', label: id,
  displayName: 'Standard', modelName: id, license: 'MIT', directoryName: id,
  fileName: 'model.onnx', metadataFileName: 'model.json', inputSize: 320,
  visible: true, downloadUrl: `https://fixture.invalid/${id}`, sha256: digest(binary), ...extra });

async function fixture(t, definitions = [definition('download')], ports = {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'worldshot-model-store-'));
  t.after(async () => {
    assert(path.resolve(root).startsWith(`${path.resolve(os.tmpdir())}${path.sep}worldshot-model-store-`));
    await fs.rm(root, { recursive: true, force: true });
  });
  const roots = { bundledRootPath: path.join(root, 'bundle'), managedRootPath: path.join(root, 'managed'),
    legacyRootPath: path.join(root, 'legacy'), wasmRootPath: path.join(root, 'wasm') };
  const calls = [];
  const store = createAiSubjectModelStore({ ...roots, models: Object.fromEntries(definitions.map(d => [d.id, d])),
    fetch: async url => { calls.push(url); return new Response(binary); },
    openPath: async target => { calls.push(target); return ''; }, ...ports });
  const directory = d => path.join(d.bundled ? roots.bundledRootPath
    : d.storageRoot === 'legacy' ? roots.legacyRootPath : roots.managedRootPath, d.directoryName);
  const put = async (d, fileName = d.fileName, bytes = binary) => {
    await fs.mkdir(directory(d), { recursive: true });
    const target = path.join(directory(d), fileName);
    await fs.writeFile(target, bytes);
    return target;
  };
  return { root, ...roots, store, calls, directory, put,
    status: async () => (await store.getAiSubjectModelStatus()).models };
}

test('catalog keeps shipped defaults, paths, hashes and eligibility', () => {
  assert.deepEqual(AI_MODEL_TIERS, catalogBaseline.tiers);
  assert.deepEqual(AI_SUBJECT_MODELS, catalogBaseline.models);
  assert.equal(AI_MODEL_TIERS.lightweight.defaultModelId, 'u2netp');
  assert.equal(AI_MODEL_TIERS.standard.defaultModelId, 'withoutbg-snap');
  assert.equal(AI_MODEL_TIERS.highQuality.defaultModelId, 'withoutbg-focus');
  assert.equal(AI_SUBJECT_MODELS.u2netp.sha256, '309c8469258dda742793dce0ebea8e6dd393174f89934733ecc8b14c76f4ddd8');
  for (const d of Object.values(AI_SUBJECT_MODELS)) {
    assert.equal(path.basename(d.directoryName), d.directoryName);
    for (const file of d.files || [d]) {
      assert.equal(path.basename(file.fileName), file.fileName);
      if (file.sha256) assert.match(file.sha256, /^[a-f0-9]{64}$/);
    }
  }
});

test('missing files report download capability without touching disk or network', async t => {
  const f = await fixture(t);
  const [s] = await f.status();
  assert.equal(s.installed, false);
  assert.equal(s.partial, false);
  assert.equal(s.ready, false);
  assert.equal(s.canDownload, true);
  assert.equal(s.canDelete, false);
  assert.equal(s.message, 'モデルファイルが見つかりません');
  assert.deepEqual(f.calls, []);
  assert.deepEqual(await fs.readdir(f.root), []);
});

test('bundle, managed and legacy status preserve paths, role map and metadata precedence', async t => {
  const defs = [definition('bundle', { bundled: true }), definition('managed'),
    definition('legacy', { storageRoot: 'legacy', deprecated: true })];
  const f = await fixture(t, defs);
  for (const d of defs) {
    await f.put(d);
    await f.put(d, 'model.json', JSON.stringify({ license: 'Apache-2.0', inputSize: 512,
      description: 'installed metadata', provider: 'metadata provider' }));
  }
  const states = await f.status();
  for (let i = 0; i < defs.length; i++) {
    const s = states[i];
    assert.equal(s.ready, true);
    assert.equal(s.verified, true);
    assert.equal(s.modelDirectoryPath, f.directory(defs[i]));
    assert.equal(s.modelUrl, pathToFileURL(path.join(f.directory(defs[i]), 'model.onnx')).href);
    assert.equal(s.modelFiles.model, s.modelUrl);
    assert.equal(s.wasmBaseUrl, pathToFileURL(`${f.wasmRootPath}${path.sep}`).href);
    assert.equal(s.license, 'Apache-2.0');
    assert.equal(s.inputSize, 512);
    assert.equal(s.canDelete, !defs[i].bundled);
    assert.equal(s.canDownload, !defs[i].bundled);
  }
});

test('multi-file partial, corrupted, repaired and missing states cannot become ready', async t => {
  const d = definition('multi', { files: [
    { role: 'alpha', fileName: 'alpha.onnx', sha256: digest(binary), downloadUrl: 'https://fixture.invalid/a' },
    { role: 'beta', fileName: 'beta.onnx', sha256: digest(binary), downloadUrl: 'https://fixture.invalid/b' },
  ] });
  const f = await fixture(t, [d]);
  await f.put(d, 'alpha.onnx');
  let [s] = await f.status();
  assert.equal(s.partial, true);
  assert.equal(s.ready, false);
  assert.equal(s.canDelete, true);
  assert.equal(s.message, 'モデルファイルが一部不足しています');
  await f.put(d, 'beta.onnx', 'bad');
  [s] = await f.status();
  assert.equal(s.installed, true);
  assert.equal(s.verified, false);
  assert.equal(s.message, 'モデルファイルの検証に失敗しました');
  await f.put(d, 'beta.onnx');
  [s] = await f.status();
  assert.equal(s.ready, true);
  assert.equal(s.files[0].sha256, digest(binary));
  await fs.rm(path.join(f.directory(d), 'alpha.onnx'));
  assert.equal((await f.status())[0].ready, false);
});

test('invalid metadata and optional hashes retain established status behavior', async t => {
  const d = definition('no-hash', { sha256: '' });
  const f = await fixture(t, [d]);
  await f.put(d);
  await f.put(d, 'model.json', '{invalid');
  const [s] = await f.status();
  assert.equal(s.ready, true);
  assert.equal(s.license, 'MIT');
  assert.equal(s.canDownload, false);
});

for (const flags of [{ bundled: true }, { future: true }, { licenseRestricted: true },
  { runtimeSupported: false }, { downloadUrl: '' }, { sha256: '' }]) {
  test(`download refuses unavailable model ${JSON.stringify(flags)}`, async t => {
    const f = await fixture(t, [definition('blocked', flags)]);
    await assert.rejects(f.store.downloadAiSubjectModel('blocked'), /ダウンロード対象/);
    assert.deepEqual(f.calls, []);
    assert.deepEqual(await fs.readdir(f.root), []);
  });
}

test('unknown model actions fail without side effects; main-process inference stays disabled', async t => {
  const f = await fixture(t);
  for (const action of ['downloadAiSubjectModel', 'deleteAiSubjectModel', 'openAiSubjectModelFolder']) {
    await assert.rejects(f.store[action]('unknown'), /見つかりません/);
  }
  await assert.rejects(f.store.runAiSubjectModelInMain({ modelId: 'unknown' }), /見つかりません/);
  await assert.rejects(f.store.runAiSubjectModelInMain({ modelId: 'download' }), /利用できません/);
  assert.deepEqual(f.calls, []);
});

test('multi-file streamed download writes exact bytes, progress, metadata and ignores optional text errors', async t => {
  const second = Buffer.alloc(160000, 123);
  const d = definition('multi', { files: [
    { role: 'alpha', fileName: 'alpha.onnx', downloadUrl: 'https://fixture.invalid/a', sha256: digest(binary), sizeBytes: binary.length },
    { role: 'beta', fileName: 'beta.onnx', downloadUrl: 'https://fixture.invalid/b', sha256: digest(second), sizeBytes: second.length },
  ], supplementalFiles: [{ fileName: 'README.md', downloadUrl: 'https://fixture.invalid/readme' },
    { fileName: 'optional.txt', downloadUrl: 'https://fixture.invalid/fail' }] });
  const f = await fixture(t, [d], { fetch: async url => {
    if (url.endsWith('fail')) throw new Error('optional network failure');
    const bytes = url.endsWith('/a') ? binary : url.endsWith('/b') ? second : Buffer.from('fixture readme');
    let offset = 0;
    return new Response(new ReadableStream({ pull(controller) {
      if (offset >= bytes.length) { controller.close(); return; }
      controller.enqueue(bytes.subarray(offset, offset + 2000)); offset += 2000;
    } }), { headers: { 'content-length': String(bytes.length) } });
  } });
  const events = [];
  const result = await f.store.downloadAiSubjectModel('multi', { send: (channel, event) => {
    assert.equal(channel, 'ai-subject-model-download-progress'); events.push(event);
  } });
  assert.equal(result.models[0].ready, true);
  assert.deepEqual(await fs.readFile(path.join(f.directory(d), 'alpha.onnx')), binary);
  assert.deepEqual(await fs.readFile(path.join(f.directory(d), 'beta.onnx')), second);
  assert.equal(events.at(-1).receivedBytes, binary.length + second.length);
  assert.equal(events.at(-1).totalBytes, binary.length + second.length);
  assert(events.every((e, i) => i === 0 || e.receivedBytes >= events[i - 1].receivedBytes));
  const metadata = JSON.parse(await fs.readFile(path.join(f.directory(d), 'model.json'), 'utf8'));
  assert.equal(metadata.size, binary.length + second.length);
  assert.equal(metadata.files.length, 2);
  assert.equal(metadata.source, 'official');
  assert.equal(metadata.localOnly, true);
  assert.equal(metadata.usesCloudApi, false);
  assert(!Number.isNaN(Date.parse(metadata.downloadedAt)));
  assert.deepEqual((await fs.readdir(f.directory(d))).sort(), ['README.md', 'alpha.onnx', 'beta.onnx', 'model.json']);
});

test('unavailable progress receiver does not fail a valid download', async t => {
  const f = await fixture(t);
  const result = await f.store.downloadAiSubjectModel('download', { send: () => { throw Error('window closed'); } });
  assert.equal(result.models[0].ready, true);
});

for (const mode of ['network', 'http', 'hash', 'body']) {
  test(`${mode} failure preserves existing model and removes temporary file`, async t => {
    const d = definition('download');
    const f = await fixture(t, [d], { fetch: async () => {
      if (mode === 'network') throw Error('network failed');
      if (mode === 'http') return new Response('error', { status: 503 });
      if (mode === 'hash') return new Response('incorrect binary');
      return new Response(new ReadableStream({ start(c) { c.enqueue(binary); c.error(Error('body failed')); } }));
    } });
    await f.put(d);
    await f.put(d, 'model.json', 'original metadata');
    await assert.rejects(f.store.downloadAiSubjectModel('download'));
    assert.deepEqual(await fs.readFile(path.join(f.directory(d), d.fileName)), binary);
    assert.equal(await fs.readFile(path.join(f.directory(d), 'model.json'), 'utf8'), 'original metadata');
    assert.deepEqual((await fs.readdir(f.directory(d))).sort(), ['model.json', 'model.onnx']);
  });
}

test('delete protects bundle and sibling directories and removes only chosen legacy/managed model', async t => {
  const defs = [definition('bundle', { bundled: true }), definition('managed'),
    definition('legacy', { storageRoot: 'legacy' }), definition('escape', { directoryName: '../managed-other' })];
  const f = await fixture(t, defs);
  for (const d of defs) await f.put(d);
  await assert.rejects(f.store.deleteAiSubjectModel('bundle'), /同梱AIモデル/);
  await assert.rejects(f.store.deleteAiSubjectModel('escape'), /削除先が不正/);
  for (const id of ['managed', 'legacy']) {
    await f.store.deleteAiSubjectModel(id);
    await assert.rejects(fs.stat(f.directory(defs.find(d => d.id === id))), { code: 'ENOENT' });
  }
  assert.deepEqual(await fs.readFile(path.join(f.directory(defs[0]), 'model.onnx')), binary);
  assert.deepEqual(await fs.readFile(path.join(f.directory(defs[3]), 'model.onnx')), binary);
});

test('folder opening creates the chosen directory and returns shell failure', async t => {
  const d = definition('download');
  const f = await fixture(t, [d]);
  assert.deepEqual(await f.store.openAiSubjectModelFolder('download'), { ok: true });
  assert.equal(f.calls[0], f.directory(d));
  assert.equal((await fs.stat(f.directory(d))).isDirectory(), true);
  const failed = await fixture(t, [d], { openPath: async () => 'shell failure' });
  await assert.rejects(failed.store.openAiSubjectModelFolder('download'), /shell failure/);
});

test('filesystem write failure rejects download without an uncaught process error', async t => {
  const f = await fixture(t);
  const modulePath = require.resolve(process.env.MODEL_STORE_BASELINE
    ? '../out/refactor-model-baseline/adapter.cjs' : '../src/ai-subject-model-store');
  const child = spawnSync(process.execPath, [path.join(__dirname, 'lib/model-stream-error-probe.js'),
    modulePath, f.root], { encoding: 'utf8', timeout: 5000, windowsHide: true });
  assert.equal(child.error, undefined);
  assert.equal(child.status, 0, child.stderr);
  assert.equal(JSON.parse(child.stdout).ok, true);
});
