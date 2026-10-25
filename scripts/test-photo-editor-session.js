const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createController } = require('../src/photo-editor-session');

const clone = value => JSON.parse(JSON.stringify(value));

function harness(options = {}) {
  const h = { state: null, timers: new Map(), nextTimer: 0, changes: 0, onApply: null };
  h.controller = createController({
    captureSnapshot: () => clone(h.state.document),
    applySnapshot: snapshot => {
      if (h.onApply) h.onApply(snapshot);
      else h.state.document = clone(snapshot);
    },
    onHistoryChanged: () => { h.changes += 1; },
    setTimer: (callback, delay) => {
      const id = ++h.nextTimer;
      h.timers.set(id, { callback, delay, cleared: false });
      return id;
    },
    clearTimer: id => { h.timers.get(id).cleared = true; },
    ...options,
  });
  h.activate = (value = 0) => {
    h.state = { document: { value } };
    h.controller.activate(h.state);
    return h.state;
  };
  h.edit = value => {
    h.controller.beginMutation();
    h.state.document.value = value;
    h.controller.commitMutation();
  };
  // Deliberately invoke callbacks even after clearTimer, as if already queued.
  h.fire = id => h.timers.get(id).callback();
  h.history = () => h.controller.getHistoryState();
  return h;
}

test('CommonJS and browser global expose the same DOM-free controller API', () => {
  const context = { setTimeout() {}, clearTimeout() {} };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.resolve(__dirname, '../src/photo-editor-session.js'), 'utf8'), context);
  assert.equal(typeof context.PhotoEditorSession.createController, 'function');
  const browserController = context.PhotoEditorSession.createController({
    captureSnapshot: () => ({ value: 1 }), applySnapshot() {},
  });
  assert.deepEqual(Object.keys(browserController).sort(), Object.keys(harness().controller).sort());
});

test('inactive methods are harmless and activation begins an empty history', () => {
  const h = harness();
  const c = h.controller;
  c.beginMutation(); c.commitMutation(); c.scheduleCommit(); c.undo(); c.redo();
  c.restoreSnapshot({ value: 1 }); c.deactivate();
  assert.equal(c.beginRequest('subject-mask'), null);
  assert.equal(c.isActive(null), false);
  assert.deepEqual(h.history(), { undo: [], redo: [], pending: null });
  assert.deepEqual(c.getHistoryStatus(), { canUndo: false, canRedo: false });
  const state = h.activate();
  assert.equal(c.isActive(state), true);
  assert.deepEqual(h.history(), { undo: [], redo: [], pending: null });
});

test('repeated begin coalesces changes and history inspection returns independent copies', () => {
  const h = harness();
  const c = h.controller;
  h.activate();
  c.beginMutation();
  h.state.document.value = 1;
  c.beginMutation();
  h.state.document.value = 2;
  const exposedPending = h.history();
  exposedPending.pending.value = 99;
  c.commitMutation();
  assert.deepEqual(h.history(), { undo: [{ value: 0 }], redo: [], pending: null });
  const exposedUndo = h.history();
  exposedUndo.undo[0].value = 99;
  exposedUndo.undo.push({ value: 99 });
  c.undo();
  assert.equal(h.state.document.value, 0);
  const exposedRedo = h.history();
  exposedRedo.redo[0].value = 99;
  c.redo();
  assert.equal(h.state.document.value, 2);
});

test('no-op commits create no history and preserve an existing redo branch', () => {
  const h = harness();
  const c = h.controller;
  h.activate();
  c.beginMutation(); c.commitMutation();
  assert.deepEqual(h.history(), { undo: [], redo: [], pending: null });
  h.edit(1); c.undo();
  const before = h.history();
  c.beginMutation(); c.commitMutation();
  assert.deepEqual(h.history(), before);
  assert.deepEqual(c.getHistoryStatus(), { canUndo: false, canRedo: true });
  c.redo();
  assert.equal(h.state.document.value, 1);
});

test('undo flushes an uncommitted edit before restoring the previous state', () => {
  const h = harness();
  const c = h.controller;
  h.activate();
  c.beginMutation(); h.state.document.value = 1; c.scheduleCommit();
  const timer = h.nextTimer;
  assert.deepEqual(c.getHistoryStatus(), { canUndo: true, canRedo: false });
  c.undo();
  assert.equal(h.state.document.value, 0);
  assert.deepEqual(h.history(), { undo: [], redo: [{ value: 1 }], pending: null });
  assert.equal(h.timers.get(timer).cleared, true);
  h.fire(timer);
  assert.deepEqual(h.history(), { undo: [], redo: [{ value: 1 }], pending: null });
});

test('redo commits a new pending edit and invalidates the old redo branch', () => {
  const h = harness();
  const c = h.controller;
  h.activate(); h.edit(1); h.edit(2); c.undo();
  c.beginMutation(); h.state.document.value = 3;
  c.redo();
  assert.equal(h.state.document.value, 3);
  assert.deepEqual(h.history(), { undo: [{ value: 0 }, { value: 1 }], redo: [], pending: null });
});

test('history retains only the most recent configured number of committed edits', () => {
  const h = harness({ historyLimit: 3 });
  const c = h.controller;
  h.activate();
  for (let value = 1; value <= 5; value += 1) h.edit(value);
  assert.deepEqual(h.history().undo, [{ value: 2 }, { value: 3 }, { value: 4 }]);
  for (const expected of [4, 3, 2, 2]) {
    c.undo(); assert.equal(h.state.document.value, expected);
  }
  for (const expected of [3, 4, 5, 5]) {
    c.redo(); assert.equal(h.state.document.value, expected);
  }
});

test('rescheduled or cleared timer callbacks cannot commit later work in the same session', () => {
  const h = harness();
  const c = h.controller;
  h.activate(); c.beginMutation(); h.state.document.value = 1;
  c.scheduleCommit(); const first = h.nextTimer;
  assert.equal(h.timers.get(first).delay, 420);
  c.scheduleCommit(75); const replacement = h.nextTimer;
  assert.equal(h.timers.get(first).cleared, true);
  assert.equal(h.timers.get(replacement).delay, 75);
  const pending = h.history();
  h.fire(first); assert.deepEqual(h.history(), pending);
  c.clearCommitTimer(); h.fire(replacement);
  assert.deepEqual(h.history(), pending);
  c.scheduleCommit(50); const final = h.nextTimer;
  h.fire(final);
  assert.deepEqual(h.history(), { undo: [{ value: 0 }], redo: [], pending: null });
  h.fire(final);
  assert.equal(h.history().undo.length, 1);
});

test('replacement and close discard the old owner timer, even if its callback runs later', () => {
  const h = harness();
  const c = h.controller;
  const stateA = h.activate(); c.beginMutation(); h.state.document.value = 1;
  c.scheduleCommit(); const timerA = h.nextTimer;
  const stateB = h.activate(10); c.beginMutation(); h.state.document.value = 11;
  c.scheduleCommit(); const timerB = h.nextTimer;
  assert.equal(c.isActive(stateA), false);
  assert.equal(c.isActive(stateB), true);
  const pendingB = h.history();
  h.fire(timerA);
  assert.deepEqual(h.history(), pendingB);
  assert.equal(h.timers.get(timerB).cleared, false);
  c.deactivate(); h.fire(timerB);
  assert.equal(h.timers.get(timerB).cleared, true);
  assert.deepEqual(h.history(), { undo: [], redo: [], pending: null });
});

test('restore suppresses nested mutations and history navigation', () => {
  const h = harness();
  const c = h.controller;
  h.activate(); h.edit(1);
  const before = h.history();
  h.onApply = snapshot => {
    c.beginMutation(); c.scheduleCommit(); c.commitMutation(); c.undo(); c.redo();
    h.state.document = clone(snapshot);
  };
  c.restoreSnapshot({ value: 7 });
  assert.equal(h.state.document.value, 7);
  assert.deepEqual(h.history(), before);
  assert.equal(h.timers.size, 0);
});

test('a throwing restore releases the restoration guard for subsequent edits', () => {
  const h = harness();
  const c = h.controller;
  h.activate();
  h.onApply = () => { throw new Error('controlled restore failure'); };
  assert.throws(() => c.restoreSnapshot({ value: 1 }), /controlled restore failure/);
  h.onApply = null;
  h.edit(2);
  assert.deepEqual(h.history(), { undo: [{ value: 0 }], redo: [], pending: null });
  c.undo();
  assert.equal(h.state.document.value, 0);
});

test('latest request wins per channel; an old finish cannot cancel its replacement', () => {
  const h = harness();
  const c = h.controller;
  h.activate();
  const first = c.beginRequest('subject-mask');
  const image = c.beginRequest('source-image');
  const replacement = c.beginRequest('subject-mask');
  assert.equal(c.isRequestCurrent(first), false);
  c.finishRequest(first);
  assert.equal(c.isRequestCurrent(replacement), true);
  assert.equal(c.isRequestCurrent(image), true);
  c.invalidateRequest('subject-mask');
  assert.equal(c.isRequestCurrent(replacement), false);
  assert.equal(c.isRequestCurrent(image), true);
  c.finishRequest(image);
  assert.equal(c.isRequestCurrent(image), false);
});

test('close and reopen invalidate old requests even when reusing the same state object', () => {
  const h = harness();
  const c = h.controller;
  const state = h.activate();
  const first = c.beginRequest('subject-mask');
  c.deactivate();
  assert.equal(c.isRequestCurrent(first), false);
  c.activate(state);
  const replacement = c.beginRequest('subject-mask');
  c.finishRequest(first);
  assert.equal(c.isRequestCurrent(first), false);
  assert.equal(c.isRequestCurrent(replacement), true);
  h.activate();
  assert.equal(c.isRequestCurrent(replacement), false);
});

test('restoring history cancels subject selection but keeps independent request channels', () => {
  const h = harness();
  const c = h.controller;
  h.activate();
  const subject = c.beginRequest('subject-mask');
  const image = c.beginRequest('source-image');
  c.restoreSnapshot({ value: 1 });
  assert.equal(c.isRequestCurrent(subject), false);
  assert.equal(c.isRequestCurrent(image), true);
});
