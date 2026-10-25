(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PhotoEditorSession = api;
})(globalThis, function () {
  'use strict';

  const clone = (value) => JSON.parse(JSON.stringify(value || null));
  const equal = (left, right) => JSON.stringify(left || null) === JSON.stringify(right || null);

  // Owns the lifetime of one editor, its undo history, and asynchronous requests.
  // Image/state normalization and all DOM/file effects are supplied by the caller.
  function createController({
    captureSnapshot, applySnapshot, onHistoryChanged = () => {},
    historyLimit = 40, commitDelay = 420,
    setTimer = setTimeout, clearTimer = clearTimeout,
  }) {
    let active = null;

    function clearCommitTimer() {
      if (active?.timer !== null && active?.timer !== undefined) clearTimer(active.timer);
      if (active) active.timer = null;
    }

    function deactivate() {
      clearCommitTimer();
      if (active) active.requests.clear();
      active = null;
    }

    function activate(state) {
      deactivate();
      active = { state, undo: [], redo: [], pending: null, timer: null,
        restoring: false, requests: new Map() };
    }

    function isActive(state) {
      return Boolean(active && active.state === state);
    }

    function beginRequest(channel) {
      if (!active) return null;
      const request = { owner: active, channel };
      active.requests.set(channel, request);
      return request;
    }

    function isRequestCurrent(request) {
      return Boolean(request && active === request.owner && active.requests.get(request.channel) === request);
    }

    function finishRequest(request) {
      if (isRequestCurrent(request)) active.requests.delete(request.channel);
    }

    function invalidateRequest(channel) {
      active?.requests.delete(channel);
    }

    function getHistoryStatus() {
      return { canUndo: Boolean(active && (active.undo.length || active.pending)),
        canRedo: Boolean(active?.redo.length) };
    }

    function getHistoryState() {
      return { undo: clone(active?.undo || []), redo: clone(active?.redo || []),
        pending: clone(active?.pending) };
    }

    function beginMutation() {
      if (active && !active.restoring && !active.pending) active.pending = captureSnapshot();
    }

    function commitMutation() {
      if (!active || active.restoring) return;
      clearCommitTimer();
      const previous = active.pending;
      active.pending = null;
      if (previous && !equal(previous, captureSnapshot())) {
        active.undo.push(previous);
        if (active.undo.length > historyLimit) active.undo.shift();
        active.redo = [];
      }
      onHistoryChanged();
    }

    function scheduleCommit(delay = commitDelay) {
      if (!active?.pending || active.restoring) return;
      clearCommitTimer();
      const owner = active;
      const timer = setTimer(() => {
        // Clearing a timer is not sufficient if its callback is already queued.
        if (active !== owner || owner.timer !== timer) return;
        owner.timer = null;
        commitMutation();
      }, delay);
      owner.timer = timer;
      onHistoryChanged();
    }

    function restoreSnapshot(snapshot) {
      if (!active || !snapshot) return;
      const owner = active;
      owner.requests.delete('subject-mask');
      owner.restoring = true;
      try {
        applySnapshot(snapshot);
      } finally {
        owner.restoring = false;
      }
    }

    function moveHistory(direction) {
      if (!active || active.restoring) return;
      commitMutation();
      const source = direction === 'undo' ? active.undo : active.redo;
      const target = direction === 'undo' ? active.redo : active.undo;
      const snapshot = source.pop();
      if (snapshot) {
        target.push(captureSnapshot());
        restoreSnapshot(snapshot);
      }
      onHistoryChanged();
    }

    return { activate, deactivate, isActive, beginRequest, isRequestCurrent,
      finishRequest, invalidateRequest, getHistoryStatus, getHistoryState,
      beginMutation, commitMutation, scheduleCommit, clearCommitTimer,
      restoreSnapshot, undo: () => moveHistory('undo'), redo: () => moveHistory('redo') };
  }

  return Object.freeze({ createController });
});
