(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PhotoEditorImageResources = api;
})(globalThis, function () {
  'use strict';

  // Cache ownership belongs to the caller; image creation and redraw notification
  // are ports so the loading lifecycle can be checked without DOM or Electron.
  function getOrLoad(cache, key, { createImage, onReady = () => {} }) {
    if (!key) return null;
    const cached = cache.get(key);
    if (cached) return cached;
    const image = createImage();
    let resolveLoad;
    const entry = { image, loaded: false, failed: false,
      loadPromise: new Promise(resolve => { resolveLoad = resolve; }) };
    image.onload = () => {
      entry.loaded = true;
      entry.failed = false;
      resolveLoad(entry);
      if (cache.get(key) === entry) onReady(entry);
    };
    image.onerror = () => {
      entry.failed = true;
      resolveLoad(entry);
    };
    image.decoding = 'async';
    cache.set(key, entry);
    try { image.src = key; }
    catch (error) {
      if (cache.get(key) === entry) cache.delete(key);
      throw error;
    }
    return entry;
  }

  // Return a guard as well as awaiting decode. Export checks it again after
  // Worker awaits, because Undo/Redo can evict a required entry in that time.
  async function waitForReady(resources, { assertCurrent = () => {}, canceledMessage } = {}) {
    assertCurrent();
    const required = resources.map(resource => ({ resource, entry: resource.getEntry() }));
    await Promise.all(required.map(async ({ resource, entry }) => {
      try { await entry?.loadPromise; }
      catch { throw new Error(resource.errorMessage); }
    }));
    const assertReady = () => {
      assertCurrent();
      for (const { resource, entry } of required) {
        if (resource.getEntry() !== entry) throw new Error(canceledMessage);
        if (!entry?.loaded || entry.failed || !entry.image) throw new Error(resource.errorMessage);
      }
    };
    assertReady();
    return assertReady;
  }

  // Fonts may legitimately be unavailable offline. Give optional resources a
  // bounded chance to finish, then preserve the existing browser fallback.
  async function waitForOptional(loaders, { assertCurrent = () => {}, timeoutMs = 5000,
    setTimer = setTimeout, clearTimer = clearTimeout } = {}) {
    assertCurrent();
    if (loaders.length === 0) return { timedOut: false };
    const waitMs = Number.isFinite(timeoutMs) && timeoutMs >= 0 ? timeoutMs : 5000;
    const ready = Promise.allSettled(loaders.map(load => Promise.resolve().then(load)))
      .then(() => ({ timedOut: false }));
    let timer;
    try {
      const timeout = new Promise(resolve => { timer = setTimer(() => resolve({ timedOut: true }), waitMs); });
      const result = await Promise.race([ready, timeout]);
      assertCurrent();
      return result;
    } finally {
      if (timer !== undefined) clearTimer(timer);
    }
  }

  return Object.freeze({ getOrLoad, waitForReady, waitForOptional });
});
