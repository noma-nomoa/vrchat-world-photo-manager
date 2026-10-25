// Run with: node --test scripts/test-photo-editor-save.js
// Save-controller contracts use the real session controller and fake render/IPC
// boundaries. No Electron, image files, database, or native dialogs are involved.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createController: createSession } = require('../src/photo-editor-session');
const PhotoEditorSave = require('../src/photo-editor-save');

const DATA_URL = 'data:image/png;base64,c3ludGhldGlj';
const clone = value => JSON.parse(JSON.stringify(value));

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function createState(name = 'A') {
  return {
    isSaving: false,
    sourcePhoto: {
      filePath: `synthetic/${name}.png`,
      fileName: `${name}.png`,
      takenAtTimestamp: name === 'A' ? 1788912000000 : 1788912060000,
    },
  };
}

function harness(options = {}) {
  const h = {
    state: createState(),
    sourceMetadata: {
      worldId: 'wrld_synthetic_A', worldNameManual: 'Manual A', memoText: 'Memo A',
      labels: [{ name: 'Label A', colorHex: '#14B8A6' }], isFavorite: true,
    },
    exportSettings: { format: 'png', quality: 0.94, maxEdge: 2048 },
    events: [], requests: [], starts: [], finishes: [], renders: [], payloads: [], synchronizations: [],
    renderStarted: deferred(), persistStarted: deferred(), syncStarted: deferred(),
    successResult: { ok: true, canceled: false, fileName: 'saved.png', importFailed: false,
      importResult: { importedCount: 1, selectedMonth: { year: 2026, month: 9 } } },
  };
  h.session = createSession({ captureSnapshot: () => ({}), applySnapshot() {} });
  // Record request lifecycle while delegating every operation to the real module.
  const beginRequest = h.session.beginRequest;
  h.session.beginRequest = channel => {
    const request = beginRequest(channel);
    h.requests.push(request);
    h.events.push(`begin:${channel}`);
    return request;
  };
  const finishRequest = h.session.finishRequest;
  h.session.finishRequest = request => {
    h.events.push('finish-request');
    finishRequest(request);
  };
  if (options.active !== false) h.session.activate(h.state);
  h.controller = PhotoEditorSave.createController({
    session: h.session,
    render: async (settings, renderOptions) => {
      h.events.push('render');
      h.renders.push({ settings, options: renderOptions });
      h.renderStarted.resolve();
      return options.render ? options.render(settings, renderOptions) : { dataUrl: DATA_URL };
    },
    persist: async payload => {
      h.events.push('persist');
      h.payloads.push(payload);
      h.persistStarted.resolve();
      return options.persist ? options.persist(payload) : h.successResult;
    },
    synchronize: async (result, context) => {
      h.events.push('synchronize');
      h.synchronizations.push({ result, context, saving: h.state.isSaving });
      h.syncStarted.resolve();
      if (options.synchronize) await options.synchronize(result, context);
    },
    onStart: state => {
      h.events.push('start');
      h.starts.push({ state, saving: state.isSaving });
      if (options.onStart) options.onStart(state);
    },
    onFinish: state => {
      h.events.push('finish');
      h.finishes.push({ state, saving: state.isSaving,
        current: h.session.isRequestCurrent(h.requests[0]) });
      if (options.onFinish) options.onFinish(state);
    },
  });
  h.save = overrides => h.controller.save({
    state: h.state,
    sourceMetadata: h.sourceMetadata,
    exportSettings: h.exportSettings,
    transparentSubjectBackground: false,
    ...overrides,
  });
  return h;
}

function assertCurrentCleanup(h) {
  assert.equal(h.state.isSaving, false);
  assert.equal(h.finishes.length, 1);
  assert.equal(h.finishes[0].state, h.state);
  assert.equal(h.finishes[0].saving, false);
  assert.equal(h.finishes[0].current, true, 'onFinish runs while its own request is still current');
  assert.equal(h.session.isRequestCurrent(h.requests[0]), false);
}

test('CommonJS and browser global expose the same DOM-free save-controller API', () => {
  const context = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.resolve(__dirname, '../src/photo-editor-save.js'), 'utf8'), context);
  assert.equal(typeof context.PhotoEditorSave.createController, 'function');
  assert.deepEqual(Object.keys(context.PhotoEditorSave).sort(), Object.keys(PhotoEditorSave).sort());
  const browserController = context.PhotoEditorSave.createController({
    session: createSession({ captureSnapshot: () => ({}), applySnapshot() {} }),
    render() {}, persist() {}, synchronize() {}, onStart() {}, onFinish() {},
  });
  assert.deepEqual(Object.keys(browserController).sort(), Object.keys(harness().controller).sort());
});

test('inactive state is ignored without a request or any render/UI side effect', async () => {
  const h = harness({ active: false });
  assert.deepEqual(await h.save(), { kind: 'ignored' });
  assert.equal(h.state.isSaving, false);
  assert.deepEqual(h.events, []);
});

test('a different inactive state is ignored even while another editor is active', async () => {
  const h = harness();
  const otherState = createState('B');
  assert.deepEqual(await h.save({ state: otherState }), { kind: 'ignored' });
  assert.equal(otherState.isSaving, false);
  assert.equal(h.state.isSaving, false);
  assert.deepEqual(h.events, []);
});

test('an active state already saving is ignored without clearing its saving flag', async () => {
  const h = harness();
  h.state.isSaving = true;
  assert.deepEqual(await h.save(), { kind: 'ignored' });
  assert.equal(h.state.isSaving, true);
  assert.deepEqual(h.events, []);
});

test('successful save preserves payload fields and lifecycle ordering', async () => {
  const h = harness();
  const outcome = await h.save({ transparentSubjectBackground: true });
  assert.deepEqual(outcome, { kind: 'saved', result: h.successResult });
  assert.equal(outcome.result, h.successResult);
  assert.deepEqual(h.events, [
    'begin:save', 'start', 'render', 'persist', 'synchronize', 'finish', 'finish-request',
  ]);
  assert.equal(h.starts[0].saving, true);
  assert.equal(h.starts[0].state, h.state);
  assert.deepEqual(h.renders, [{ settings: h.exportSettings, options: { transparentSubjectBackground: true } }]);
  assert.deepEqual(h.payloads, [{
    sourceFilePath: h.state.sourcePhoto.filePath,
    sourceFileName: h.state.sourcePhoto.fileName,
    sourceTakenAtTimestamp: h.state.sourcePhoto.takenAtTimestamp,
    sourcePhotoMetadata: h.sourceMetadata,
    outputFormat: h.exportSettings.format,
    outputQuality: h.exportSettings.quality,
    dataUrl: DATA_URL,
  }]);
  assert.equal(h.synchronizations.length, 1);
  assert.equal(h.synchronizations[0].result, h.successResult);
  assertCurrentCleanup(h);
});

test('a second save is blocked while the first render is outstanding', { timeout: 5000 }, async () => {
  const gate = deferred();
  const h = harness({ render: () => gate.promise });
  const first = h.save();
  await h.renderStarted.promise;
  assert.equal(h.state.isSaving, true);
  assert.deepEqual(await h.save(), { kind: 'ignored' });
  assert.equal(h.renders.length, 1);
  assert.equal(h.starts.length, 1);
  assert.equal(h.requests.length, 1);
  gate.resolve({ dataUrl: DATA_URL });
  assert.equal((await first).kind, 'saved');
  assertCurrentCleanup(h);
});

for (const action of ['close', 'switch', 'same-state-new-session', 'invalidate-save']) {
  test(`${action} while rendering prevents persistence and ignores stale UI completion`, { timeout: 5000 }, async () => {
    const gate = deferred();
    const h = harness({ render: () => gate.promise });
    const pending = h.save();
    await h.renderStarted.promise;
    const otherState = createState('B');
    otherState.isSaving = true;
    if (action === 'close') h.session.deactivate();
    else if (action === 'switch') h.session.activate(otherState);
    else if (action === 'same-state-new-session') h.session.activate(h.state);
    else h.session.invalidateRequest('save');
    gate.resolve({ dataUrl: DATA_URL });
    assert.deepEqual(await pending, { kind: 'superseded' });
    assert.equal(h.payloads.length, 0);
    assert.equal(h.synchronizations.length, 0);
    assert.equal(h.finishes.length, 0);
    assert.equal(h.state.isSaving, false);
    assert.equal(otherState.isSaving, true);
  });
}

test('finishing an old save does not finish a newer request on the same channel', { timeout: 5000 }, async () => {
  const gate = deferred();
  const h = harness({ render: () => gate.promise });
  const pending = h.save();
  await h.renderStarted.promise;
  const replacement = h.session.beginRequest('save');
  gate.resolve({ dataUrl: DATA_URL });
  assert.deepEqual(await pending, { kind: 'superseded' });
  assert.equal(h.session.isRequestCurrent(replacement), true);
  assert.equal(h.finishes.length, 0);
  assert.equal(h.payloads.length, 0);
});

test('independent subject-mask requests do not cancel a save request', { timeout: 5000 }, async () => {
  const gate = deferred();
  const h = harness({ render: () => gate.promise });
  const pending = h.save();
  await h.renderStarted.promise;
  const subjectRequest = h.session.beginRequest('subject-mask');
  gate.resolve({ dataUrl: DATA_URL });
  assert.equal((await pending).kind, 'saved');
  assert.equal(h.session.isRequestCurrent(subjectRequest), true);
  assertCurrentCleanup(h);
});

test('source photo, nested metadata, and export settings are fixed before render awaits', { timeout: 5000 }, async () => {
  const gate = deferred();
  const h = harness({ render: () => gate.promise });
  const expectedPhoto = clone(h.state.sourcePhoto);
  const expectedMetadata = clone(h.sourceMetadata);
  const expectedSettings = clone(h.exportSettings);
  const pending = h.save();
  await h.renderStarted.promise;
  h.state.sourcePhoto.filePath = 'changed/path.png';
  h.state.sourcePhoto.fileName = 'changed.png';
  h.state.sourcePhoto.takenAtTimestamp = 999;
  h.state.sourcePhoto = createState('B').sourcePhoto;
  h.sourceMetadata.labels[0].name = 'Changed label';
  h.sourceMetadata.labels.push({ name: 'Added later', colorHex: '#000000' });
  h.sourceMetadata.memoText = 'Changed memo';
  h.exportSettings.format = 'jpeg';
  h.exportSettings.quality = 0.1;
  h.exportSettings.maxEdge = 320;
  assert.deepEqual(h.renders[0].settings, expectedSettings);
  gate.resolve({ dataUrl: DATA_URL });
  assert.equal((await pending).kind, 'saved');
  assert.deepEqual(h.payloads[0], {
    sourceFilePath: expectedPhoto.filePath,
    sourceFileName: expectedPhoto.fileName,
    sourceTakenAtTimestamp: expectedPhoto.takenAtTimestamp,
    sourcePhotoMetadata: expectedMetadata,
    outputFormat: expectedSettings.format,
    outputQuality: expectedSettings.quality,
    dataUrl: DATA_URL,
  });
  assert.notEqual(h.payloads[0].sourcePhotoMetadata, h.sourceMetadata);
  assert.notEqual(h.payloads[0].sourcePhotoMetadata.labels, h.sourceMetadata.labels);
  assertCurrentCleanup(h);
});

for (const [name, persistResult, expectedKind] of [
  ['canceled', { ok: false, canceled: true }, 'canceled'],
  ['failed', { ok: false, message: 'Save rejected' }, 'failed'],
  ['missing result', undefined, 'failed'],
]) {
  test(`${name} persistence result skips synchronization and cleans up`, async () => {
    const h = harness({ persist: () => persistResult });
    assert.deepEqual(await h.save(), { kind: expectedKind, result: persistResult });
    assert.equal(h.synchronizations.length, 0);
    assertCurrentCleanup(h);
  });
}

for (const [name, result] of [
  ['import failed', { ok: true, importFailed: true, importResult: { importedCount: 1 } }],
  ['no import result', { ok: true, importFailed: false }],
  ['null import result', { ok: true, importFailed: false, importResult: null }],
]) {
  test(`successful file save with ${name} skips database synchronization`, async () => {
    const h = harness({ persist: () => result });
    assert.deepEqual(await h.save(), { kind: 'saved', result });
    assert.equal(h.synchronizations.length, 0);
    assertCurrentCleanup(h);
  });
}

for (const rendered of [undefined, null, {}, { dataUrl: '' }]) {
  test(`invalid render output ${JSON.stringify(rendered)} rejects before persistence`, async () => {
    const h = harness({ render: () => rendered });
    await assert.rejects(h.save(), { message: '編集結果を描画できませんでした' });
    assert.equal(h.payloads.length, 0);
    assert.equal(h.synchronizations.length, 0);
    assertCurrentCleanup(h);
  });
}

for (const stage of ['render', 'persist', 'synchronize', 'onStart']) {
  test(`${stage} exception propagates and clears the active save lifecycle`, async () => {
    const failure = new Error(`${stage} controlled failure`);
    const h = harness({ [stage]: () => { throw failure; } });
    await assert.rejects(h.save(), error => error === failure);
    assertCurrentCleanup(h);
    if (stage === 'render' || stage === 'onStart') assert.equal(h.payloads.length, 0);
    if (stage !== 'synchronize') assert.equal(h.synchronizations.length, 0);
  });
}

test('onFinish exception still releases its request and propagates', async () => {
  const failure = new Error('onFinish controlled failure');
  const h = harness({ onFinish: () => { throw failure; } });
  await assert.rejects(h.save(), error => error === failure);
  assertCurrentCleanup(h);
});

test('late render rejection after close propagates without stale UI callbacks', { timeout: 5000 }, async () => {
  const gate = deferred();
  const h = harness({ render: () => gate.promise });
  const pending = h.save();
  await h.renderStarted.promise;
  h.session.deactivate();
  const failure = new Error('late render failure');
  const rejected = assert.rejects(pending, error => error === failure);
  gate.reject(failure);
  await rejected;
  assert.equal(h.payloads.length, 0);
  assert.equal(h.finishes.length, 0);
  assert.equal(h.state.isSaving, false);
});

test('switch after IPC starts still synchronizes the saved file but leaves the newer editor untouched', { timeout: 5000 }, async () => {
  const gate = deferred();
  const h = harness({ persist: () => gate.promise });
  const pending = h.save();
  await h.persistStarted.promise;
  const otherState = createState('B');
  otherState.isSaving = true;
  h.session.activate(otherState);
  const otherSave = h.session.beginRequest('save');
  gate.resolve(h.successResult);
  assert.deepEqual(await pending, { kind: 'saved', result: h.successResult });
  assert.equal(h.synchronizations.length, 1);
  assert.equal(h.synchronizations[0].context.isCurrent(), false);
  assert.equal(h.finishes.length, 0);
  assert.equal(h.state.isSaving, false);
  assert.equal(otherState.isSaving, true);
  assert.equal(h.session.isRequestCurrent(otherSave), true);
});

test('late IPC rejection cannot finish the newer editor or remove its save request', { timeout: 5000 }, async () => {
  const gate = deferred();
  const h = harness({ persist: () => gate.promise });
  const pending = h.save();
  await h.persistStarted.promise;
  const otherState = createState('B');
  otherState.isSaving = true;
  h.session.activate(otherState);
  const otherSave = h.session.beginRequest('save');
  const failure = new Error('late IPC failure');
  const rejected = assert.rejects(pending, error => error === failure);
  gate.reject(failure);
  await rejected;
  assert.equal(h.synchronizations.length, 0);
  assert.equal(h.finishes.length, 0);
  assert.equal(otherState.isSaving, true);
  assert.equal(h.state.isSaving, false);
  assert.equal(h.session.isRequestCurrent(otherSave), true);
});

test('saving remains active until synchronization finishes and isCurrent is evaluated dynamically', { timeout: 5000 }, async () => {
  const gate = deferred();
  const h = harness({ synchronize: () => gate.promise });
  let settled = false;
  const pending = h.save().then(value => { settled = true; return value; });
  await h.syncStarted.promise;
  const context = h.synchronizations[0].context;
  assert.equal(context.isCurrent(), true);
  assert.equal(h.state.isSaving, true);
  assert.equal(settled, false);
  assert.equal(h.finishes.length, 0);
  h.session.deactivate();
  assert.equal(context.isCurrent(), false);
  gate.resolve();
  assert.equal((await pending).kind, 'saved');
  assert.equal(settled, true);
  assert.equal(h.state.isSaving, false);
  assert.equal(h.finishes.length, 0);
});
