'use strict';
const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const resources = require('../src/photo-editor-image-resources');
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
};
function fixture() {
  const cache = new Map(), images = [], ready = [];
  const ports = { createImage: () => { const image = {}; images.push(image); return image; }, onReady: entry => ready.push(entry) };
  return { cache, images, ready, get: key => resources.getOrLoad(cache, key, ports) };
}
const descriptor = entry => ({ getEntry: () => entry, errorMessage: 'resource unavailable' });
const readyEntry = () => ({ loaded: true, failed: false, image: {}, loadPromise: Promise.resolve() });

test('cache loads each URL once and preserves the same entry and promise', () => {
  const h = fixture();
  assert.equal(h.get(''), null);
  const first = h.get('asset.png');
  assert.equal(h.get('asset.png'), first);
  assert.equal(h.images.length, 1);
  assert.equal(first.image.src, 'asset.png');
  assert.equal(first.image.decoding, 'async');
  assert.equal(first.loaded, false);
  assert.equal(first.failed, false);
});
test('successful decode settles readiness and notifies redraw with the cached entry', async () => {
  const h = fixture(), entry = h.get('asset.png');
  entry.image.onload();
  assert.equal(await entry.loadPromise, entry);
  assert.equal(entry.loaded, true);
  assert.equal(entry.failed, false);
  assert.deepEqual(h.ready, [entry]);
});
test('decode failure settles the promise as failed without triggering redraw', async () => {
  const h = fixture(), entry = h.get('missing.png');
  entry.image.onerror();
  assert.equal(await entry.loadPromise, entry);
  assert.equal(entry.loaded, false);
  assert.equal(entry.failed, true);
  assert.deepEqual(h.ready, []);
});
test('loading an evicted entry cannot notify redraw for its replacement', async () => {
  const h = fixture(), old = h.get('asset.png');
  h.cache.clear();
  const next = h.get('asset.png');
  old.image.onload();
  await old.loadPromise;
  assert.deepEqual(h.ready, []);
  assert.equal(h.cache.get('asset.png'), next);
  next.image.onload();
  assert.deepEqual(h.ready, [next]);
});
test('separate caches do not share image or error state for the same URL', () => {
  const first = fixture(), second = fixture();
  const a = first.get('same.png'), b = second.get('same.png');
  assert.notEqual(a, b);
  a.image.onerror();
  assert.equal(b.failed, false);
});
test('a synchronous image source failure cannot leave a permanently pending cache entry', () => {
  const cache = new Map();
  const error = new Error('source rejected');
  assert.throws(() => resources.getOrLoad(cache, 'bad.png', {
    createImage: () => ({ set src(_value) { throw error; } }),
  }), value => value === error);
  assert.equal(cache.has('bad.png'), false);
});
test('readiness waits for every required entry before allowing export', async () => {
  const h = fixture(), a = h.get('a.png'), b = h.get('b.png');
  let completed = false;
  const pending = resources.waitForReady([descriptor(a), descriptor(b)])
    .then(guard => { completed = true; return guard; });
  a.image.onload();
  await Promise.resolve();
  assert.equal(completed, false);
  b.image.onload();
  const guard = await pending;
  assert.equal(completed, true);
  assert.doesNotThrow(guard);
});
test('missing, failed, undecoded or imageless resources prevent export', async () => {
  for (const entry of [null, { ...readyEntry(), failed: true }, { ...readyEntry(), loaded: false }, { ...readyEntry(), image: null }]) {
    await assert.rejects(resources.waitForReady([descriptor(entry)]), /resource unavailable/);
  }
});
test('a rejected loader is reported using the resource-specific message', async () => {
  const load = deferred();
  const pending = resources.waitForReady([descriptor({ ...readyEntry(), loadPromise: load.promise })]);
  load.reject(new Error('internal decoder error'));
  await assert.rejects(pending, /resource unavailable/);
});
test('cancellation before readiness starts does not start a load', async () => {
  const error = new Error('canceled');
  await assert.rejects(resources.waitForReady([{ getEntry: () => assert.fail('load should not start') }],
    { assertCurrent: () => { throw error; } }), value => value === error);
});
test('cancellation during decode is checked before returning a readiness guard', async () => {
  const h = fixture(), entry = h.get('asset.png');
  let current = true;
  const pending = resources.waitForReady([descriptor(entry)], { assertCurrent: () => { if (!current) throw new Error('session changed'); } });
  current = false;
  entry.image.onload();
  await assert.rejects(pending, /session changed/);
});
test('Undo-style replacement during decode cancels instead of trusting the detached entry', async () => {
  const h = fixture();
  let entry = h.get('asset.png');
  const pending = resources.waitForReady([{ getEntry: () => entry, errorMessage: 'bad image' }], { canceledMessage: 'resource replaced' });
  const old = entry;
  entry = readyEntry();
  old.image.onload();
  await assert.rejects(pending, /resource replaced/);
});
test('the returned guard detects resource replacement or failure during a later Worker await', async () => {
  let entry = readyEntry();
  const guard = await resources.waitForReady([{ getEntry: () => entry, errorMessage: 'bad image' }], { canceledMessage: 'resource replaced' });
  const old = entry;
  entry = readyEntry();
  assert.throws(guard, /resource replaced/);
  entry = old;
  entry.failed = true;
  assert.throws(guard, /bad image/);
});
test('an empty requirement list still preserves the caller ownership check', async () => {
  let active = true;
  const guard = await resources.waitForReady([], { assertCurrent: () => { if (!active) throw new Error('closed'); } });
  guard();
  active = false;
  assert.throws(guard, /closed/);
});
test('browser entry point loads without DOM, Image or Electron globals', () => {
  const context = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/photo-editor-image-resources.js'), 'utf8'), context);
  assert.equal(typeof context.PhotoEditorImageResources.waitForReady, 'function');
  assert.equal(Object.isFrozen(context.PhotoEditorImageResources), true);
  assert.equal(Object.isFrozen(resources), true);
});

function fakeTimer() {
  let callback;
  const clock = { delay: null, cleared: [], setTimer: (run, delay) => { callback = run; clock.delay = delay; return 7; },
    clearTimer: id => clock.cleared.push(id), fire: () => callback() };
  return clock;
}
const tick = () => new Promise(resolve => setImmediate(resolve));

test('optional fonts start concurrently and wait for all successful or failed loads', async () => {
  const a = deferred(), b = deferred(), clock = fakeTimer(), calls = [];
  const pending = resources.waitForOptional([
    () => { calls.push('a'); return a.promise; },
    () => { calls.push('b'); return b.promise; },
  ], clock);
  await tick();
  assert.deepEqual(calls, ['a', 'b']);
  assert.equal(clock.delay, 5000);
  a.resolve();
  b.reject(new Error('offline font'));
  assert.deepEqual(await pending, { timedOut: false });
  assert.deepEqual(clock.cleared, [7]);
});
test('a never-settling optional font times out instead of blocking export forever', async () => {
  const font = deferred(), clock = fakeTimer();
  const pending = resources.waitForOptional([() => font.promise], clock);
  clock.fire();
  assert.deepEqual(await pending, { timedOut: true });
  assert.deepEqual(clock.cleared, [7]);
  font.reject(new Error('late network failure'));
  await tick();
});
test('late optional font completion cannot re-run ownership checks after timeout', async () => {
  const font = deferred(), clock = fakeTimer();
  let checks = 0;
  const pending = resources.waitForOptional([() => font.promise], { ...clock, assertCurrent: () => { checks++; } });
  clock.fire();
  await pending;
  assert.equal(checks, 2);
  font.resolve();
  await tick();
  assert.equal(checks, 2);
});
test('optional font loader exceptions preserve fallback and clear the wait timer', async () => {
  const clock = fakeTimer();
  assert.deepEqual(await resources.waitForOptional([() => { throw new Error('font API unavailable'); }], clock), { timedOut: false });
  assert.deepEqual(clock.cleared, [7]);
});
test('an empty optional list does not allocate a timer', async () => {
  assert.deepEqual(await resources.waitForOptional([], { setTimer: () => assert.fail('unexpected timer') }), { timedOut: false });
});
test('a changed editor during font loading cancels export and clears its timer', async () => {
  const font = deferred(), clock = fakeTimer();
  let active = true;
  const pending = resources.waitForOptional([() => font.promise], { ...clock,
    assertCurrent: () => { if (!active) throw new Error('editor changed'); } });
  active = false;
  font.resolve();
  await assert.rejects(pending, /editor changed/);
  assert.deepEqual(clock.cleared, [7]);
});
test('a changed editor when the font deadline expires still cancels instead of exporting fallback', async () => {
  const clock = fakeTimer();
  let active = true;
  const pending = resources.waitForOptional([() => new Promise(() => {})], { ...clock,
    assertCurrent: () => { if (!active) throw new Error('editor changed'); } });
  active = false;
  clock.fire();
  await assert.rejects(pending, /editor changed/);
  assert.deepEqual(clock.cleared, [7]);
});
test('invalid optional font timeout values retain a finite default', async () => {
  for (const timeoutMs of [Infinity, NaN, -1]) {
    const clock = fakeTimer();
    const pending = resources.waitForOptional([() => new Promise(() => {})], { ...clock, timeoutMs });
    assert.equal(clock.delay, 5000);
    clock.fire();
    assert.deepEqual(await pending, { timedOut: true });
  }
});
