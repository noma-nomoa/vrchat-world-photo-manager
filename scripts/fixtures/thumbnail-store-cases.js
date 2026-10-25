'use strict';

const assert = require('node:assert/strict');
const path = require('node:path');
const { createHash } = require('node:crypto');

const digest = bytes => createHash('sha256').update(bytes).digest('hex');

async function runThumbnailStoreCases({ createStore, fs, nativeImage, root }) {
  const results = [];
  const sourceDirectory = path.join(root, 'sources');
  await fs.mkdir(sourceDirectory, { recursive: true });
  const sourcePaths = {};
  const sourceHashes = {};
  for (const name of ['plain.png', 'plain.jpg', 'plain.webp']) {
    const bytes = await fs.readFile(path.join(__dirname, 'photo-metadata', name));
    sourcePaths[name] = path.join(sourceDirectory, name);
    sourceHashes[name] = digest(bytes);
    await fs.writeFile(sourcePaths[name], bytes);
  }
  const sourceImage = nativeImage.createFromPath(sourcePaths['plain.png']);
  assert.equal(sourceImage.isEmpty(), false);

  const guard = target => {
    const relative = path.relative(root, path.resolve(target));
    assert(relative && relative !== '..' && !relative.startsWith(`..${path.sep}`)
      && !path.isAbsolute(relative), `Mutation must remain inside isolated test root: ${target}`);
  };
  const relative = target => path.relative(root, target).split(path.sep).join('/');
  const normalizeArg = arg => Buffer.isBuffer(arg) ? { bytes: arg.length, sha256: digest(arg) }
    : typeof arg === 'string' && path.isAbsolute(arg) ? relative(arg) : arg;

  async function setup(name, options = {}) {
    const home = path.join(root, name, 'home');
    const currentRoot = path.join(home, 'Current');
    const legacyRoot = path.join(home, options.sameRoot ? 'Current' : 'Legacy');
    const current = path.join(currentRoot, 'thumbnails');
    const legacy = path.join(legacyRoot, 'thumbnails');
    await fs.mkdir(home, { recursive: true });
    const calls = [];
    const updates = [];
    let rows = [];
    let database = {
      getAllPhotos: () => rows,
      updateThumbnailPath: (id, thumbnailPath) => {
        updates.push([id, relative(thumbnailPath)]);
        if (options.updateThrows) throw new Error('fixture DB failure');
        return options.returnRow ? { id, thumbnail_path: thumbnailPath, updated: true } : null;
      },
    };
    const checkedFs = {};
    for (const method of ['access', 'mkdir', 'readdir', 'rename', 'rm', 'writeFile']) {
      checkedFs[method] = async (...args) => {
        if (['mkdir', 'rename', 'rm', 'writeFile'].includes(method)) {
          guard(args[0]);
          if (method === 'rename') guard(args[1]);
        }
        calls.push([method, ...args.map(normalizeArg)]);
        if (options.fs?.[method]) return options.fs[method](...args);
        return fs[method](...args);
      };
    }
    const nativePort = options.native || nativeImage;
    const store = createStore({ fs: checkedFs, nativeImage: nativePort,
      getHomePath: () => home, getThumbnailDirectoryPath: () => current,
      getPhotoDb: () => database, storageDirname: 'Current',
      legacyStorageDirname: options.sameRoot ? 'Current' : 'Legacy', thumbnailDirname: 'thumbnails' });
    const write = async (target, bytes = 'fixture') => {
      guard(target); await fs.mkdir(path.dirname(target), { recursive: true }); await fs.writeFile(target, bytes);
    };
    return { store, home, currentRoot, legacyRoot, current, legacy, calls, updates, write,
      setRows: next => { rows = next; }, setDatabase: next => { database = next; } };
  }
  const exists = async target => { try { await fs.access(target); return true; } catch { return false; } };
  const record = (name, context, value) => results.push({ name, value,
    calls: context?.calls || [], updates: context?.updates || [] });
  const inspect = async imagePath => {
    const bytes = await fs.readFile(imagePath);
    const image = nativeImage.createFromPath(imagePath);
    assert.equal(image.isEmpty(), false);
    return { path: relative(imagePath), sha256: digest(bytes), size: image.getSize() };
  };

  {
    const ctx = await setup('paths');
    const { store, current, legacy } = ctx;
    const pathCases = [current, path.join(current, 'photo.jpg'), legacy, path.join(legacy, '海.jpg'),
      `${current}-sibling/photo.jpg`, path.dirname(current), sourcePaths['plain.png'],
      path.join(current, '..', 'outside.jpg'), '', null, undefined];
    const managed = pathCases.map(value => store.isManagedThumbnailPath(value));
    assert.deepEqual(managed, [true, true, true, true, false, false, false, false, false, false, false]);
    const rebased = pathCases.map(value => store.rebaseManagedThumbnailPath(value));
    assert.equal(rebased[2], current);
    assert.equal(rebased[3], path.join(current, '海.jpg'));
    assert(rebased.filter(Boolean).length === 2);
    const sizes = [[640, 360], [320, 180], [180, 320], [100, 100], [1, 1000],
      [0, 1], [-1, 3], [NaN, 100], ['640', 360], [Infinity, 2]];
    record('paths-and-sizing', ctx, { managed, rebased: rebased.map(value => value && relative(value)),
      roots: store.getManagedThumbnailRootPaths().map(relative),
      directories: store.getManagedThumbnailDirectoryPaths().map(relative),
      fitSizes: sizes.map(args => store.calculateThumbnailFitSize(...args)) });
  }
  {
    const ctx = await setup('same-root', { sameRoot: true });
    await ctx.write(path.join(ctx.current, 'keep.jpg'));
    await ctx.store.migrateLegacyThumbnailStorage();
    assert.equal(ctx.store.getManagedThumbnailRootPaths().length, 1);
    assert.equal(await exists(path.join(ctx.current, 'keep.jpg')), true);
    record('same-root-migration-is-noop', ctx, true);
  }
  {
    const ctx = await setup('no-legacy');
    await ctx.store.migrateLegacyThumbnailStorage();
    assert.equal(await exists(ctx.currentRoot), false);
    record('missing-legacy-is-noop', ctx, true);
  }
  {
    const ctx = await setup('whole-root');
    await ctx.write(path.join(ctx.legacy, 'old.jpg'), 'old thumbnail');
    await ctx.write(path.join(ctx.legacyRoot, 'root-sentinel'), 'retained by whole-root rename');
    await ctx.store.migrateLegacyThumbnailStorage();
    assert.equal(await exists(ctx.legacyRoot), false);
    assert.equal(await fs.readFile(path.join(ctx.current, 'old.jpg'), 'utf8'), 'old thumbnail');
    assert.equal(await exists(path.join(ctx.currentRoot, 'root-sentinel')), true);
    record('whole-root-rename-preserves-content', ctx, true);
  }
  {
    const ctx = await setup('prefer-current');
    await ctx.write(path.join(ctx.current, 'same.jpg'), 'current');
    await ctx.write(path.join(ctx.legacy, 'same.jpg'), 'legacy');
    await ctx.store.migrateLegacyThumbnailStorage();
    assert.equal(await fs.readFile(path.join(ctx.current, 'same.jpg'), 'utf8'), 'current');
    assert.equal(await fs.readFile(path.join(ctx.legacy, 'same.jpg'), 'utf8'), 'legacy');
    record('existing-current-root-is-never-merged', ctx, true);
  }
  for (const scenario of ['success', 'collision', 'missing-directory', 'entry-failure']) {
    let ctx;
    ctx = await setup(`migration-fallback-${scenario}`, { fs: {
      rename: async (source, destination) => {
        if (source === ctx.legacyRoot) {
          if (scenario === 'collision') await ctx.write(path.join(ctx.current, 'old.jpg'), 'new wins');
          throw new Error('fixture root rename failure');
        }
        if (scenario === 'entry-failure') throw new Error('fixture entry rename failure');
        return fs.rename(source, destination);
      },
    } });
    if (scenario === 'missing-directory') await ctx.write(path.join(ctx.legacyRoot, 'keep'));
    else await ctx.write(path.join(ctx.legacy, 'old.jpg'), 'legacy file');
    await ctx.store.migrateLegacyThumbnailStorage();
    if (scenario === 'success' || scenario === 'collision') {
      assert.equal(await exists(ctx.legacyRoot), false);
      assert.equal(await fs.readFile(path.join(ctx.current, 'old.jpg'), 'utf8'),
        scenario === 'collision' ? 'new wins' : 'legacy file');
    } else assert.equal(await exists(ctx.legacyRoot), true);
    record(`migration-fallback-${scenario}`, ctx, true);
  }
  {
    const ctx = await setup('reconcile');
    const migrated = path.join(ctx.current, 'found.jpg');
    await ctx.write(migrated);
    ctx.setRows([{ id: 1, thumbnail_path: path.join(ctx.legacy, 'found.jpg') },
      { id: 2, thumbnail_path: path.join(ctx.legacy, 'missing.jpg') },
      { id: 3, thumbnail_path: sourcePaths['plain.png'] }, { id: 4, thumbnail_path: migrated }, null]);
    await ctx.store.reconcileManagedThumbnailPaths();
    assert.deepEqual(ctx.updates, [[1, relative(migrated)]]);
    ctx.setDatabase(null);
    await ctx.store.reconcileManagedThumbnailPaths();
    record('reconcile-only-existing-rebased-files', ctx, true);
  }
  {
    const ctx = await setup('reconcile-db-error', { updateThrows: true });
    await ctx.write(path.join(ctx.current, 'found.jpg'));
    ctx.setRows([{ id: 1, thumbnail_path: path.join(ctx.legacy, 'found.jpg') }]);
    await ctx.store.reconcileManagedThumbnailPaths();
    record('reconcile-db-error-keeps-startup-running', ctx, true);
  }
  {
    const ctx = await setup('reset');
    await ctx.write(path.join(ctx.current, 'remove.jpg'));
    await ctx.write(path.join(ctx.legacy, 'remove.jpg'));
    await ctx.write(path.join(ctx.currentRoot, 'keep-original.png'), 'outside cache');
    await ctx.write(path.join(ctx.legacyRoot, 'keep-original.png'), 'outside cache');
    await ctx.write(`${ctx.current}-sibling/keep.png`, 'sibling');
    await ctx.store.resetManagedThumbnailDirectories();
    assert.equal(await exists(ctx.current), true);
    assert.equal(await exists(ctx.legacy), false);
    assert.equal((await fs.readdir(ctx.current)).length, 0);
    for (const target of [path.join(ctx.currentRoot, 'keep-original.png'),
      path.join(ctx.legacyRoot, 'keep-original.png'), `${ctx.current}-sibling/keep.png`]) {
      assert.equal(await exists(target), true);
    }
    record('reset-removes-only-managed-directories', ctx, true);
  }

  for (const name of ['plain.png', 'plain.jpg', 'plain.webp']) {
    const ctx = await setup(`create-${name}`);
    await fs.mkdir(ctx.current, { recursive: true });
    const image = nativeImage.createFromPath(sourcePaths[name]);
    const output = await ctx.store.createThumbnail(sourcePaths[name], sourceHashes[name], { sourceImage: image });
    assert.equal(typeof output, 'string');
    record(`create-${name}`, ctx, await inspect(output));
  }
  {
    const ctx = await setup('resize');
    await fs.mkdir(ctx.current, { recursive: true });
    const image = sourceImage.resize({ width: 640, height: 360, quality: 'good' });
    const output = await ctx.store.createThumbnailFromNativeImage(image, 'large');
    const info = await inspect(output);
    assert.deepEqual(info.size, { width: 320, height: 180 });
    record('native-image-resize', ctx, info);
  }
  {
    const ctx = await setup('reuse', { native: { createThumbnailFromPath: () => {
      throw new Error('Existing cache must avoid native decoding');
    } } });
    const target = path.join(ctx.current, 'existing.jpg');
    await ctx.write(target, 'existing cache bytes');
    assert.equal(await ctx.store.createThumbnail('not-read.png', 'existing'), target);
    assert.equal(await ctx.store.createThumbnailFromNativeImage(null, 'existing'), target);
    assert.equal(await fs.readFile(target, 'utf8'), 'existing cache bytes');
    record('existing-cache-is-reused-without-reencoding', ctx, true);
  }
  {
    const ctx = await setup('force');
    await ctx.write(path.join(ctx.current, 'existing.jpg'), 'old cache');
    const output = await ctx.store.createThumbnail(sourcePaths['plain.png'], 'existing',
      { sourceImage, force: true });
    record('force-regenerates-existing-cache', ctx, await inspect(output));
  }
  for (const scenario of ['empty-native', 'invalid-size', 'encoder-error', 'empty-encoding',
    'write-error', 'shell-error']) {
    const options = scenario === 'write-error' ? { fs: { writeFile: async () => {
      throw new Error('fixture write failure');
    } } } : scenario === 'shell-error' ? { native: { createThumbnailFromPath: async () => {
      throw new Error('fixture shell failure');
    } } } : {};
    const ctx = await setup(`failure-${scenario}`, options);
    await fs.mkdir(ctx.current, { recursive: true });
    let result;
    if (scenario === 'shell-error') result = await ctx.store.createThumbnail(sourcePaths['plain.png'], scenario);
    else {
      const image = scenario === 'empty-native' ? nativeImage.createEmpty()
        : scenario === 'invalid-size' ? { isEmpty: () => false, getSize: () => ({ width: 0, height: 0 }) }
          : scenario === 'encoder-error' || scenario === 'empty-encoding' ? {
            isEmpty: () => false, getSize: () => ({ width: 320, height: 180 }),
            toJPEG: () => { if (scenario === 'encoder-error') throw new Error('fixture encoding failure'); return Buffer.alloc(0); },
          } : sourceImage;
      result = await ctx.store.createThumbnailFromNativeImage(image, scenario);
    }
    assert.equal(result, null);
    assert.equal((await fs.readdir(ctx.current)).length, 0);
    record(`failure-${scenario}`, ctx, result);
  }
  {
    const ctx = await setup('native-error-shell-recovery');
    await fs.mkdir(ctx.current, { recursive: true });
    const brokenImage = { isEmpty: () => false, getSize: () => ({ width: 0, height: 0 }) };
    const result = await ctx.store.createThumbnail(sourcePaths['plain.png'], 'recovered', { sourceImage: brokenImage });
    record('native-failure-recovers-through-shell', ctx, await inspect(result));
  }
  {
    const ctx = await setup('invalid-rows');
    const rows = [null, {}, { id: '1', file_path: 'x', file_hash: 'y' },
      { id: 1, file_path: '', file_hash: 'y' }, { id: 1, file_path: 'x', file_hash: '' }];
    for (const row of rows) assert.equal(await ctx.store.ensureManagedThumbnailForRow(row), row);
    assert.equal(ctx.calls.length, 0);
    record('invalid-rows-are-unchanged', ctx, true);
  }
  {
    const ctx = await setup('existing-row');
    const row = { id: 1, file_path: sourcePaths['plain.png'], file_hash: 'hash', thumbnail_path: sourcePaths['plain.png'] };
    assert.equal(await ctx.store.ensureManagedThumbnailForRow(row), row);
    assert.equal(ctx.updates.length, 0);
    record('existing-row-thumbnail-is-kept', ctx, true);
  }
  for (const returnRow of [false, true]) {
    const ctx = await setup(`row-recovery-${returnRow}`, { returnRow });
    await fs.mkdir(ctx.current, { recursive: true });
    const row = { id: 1, file_path: sourcePaths['plain.png'], file_hash: 'hash', thumbnail_path: path.join(ctx.home, 'missing.jpg') };
    const result = await ctx.store.ensureManagedThumbnailForRow(row);
    assert.equal(result.thumbnail_path, path.join(ctx.current, 'hash.jpg'));
    assert.equal(ctx.updates.length, 1);
    assert.equal(row.thumbnail_path, path.join(ctx.home, 'missing.jpg'));
    record(`row-recovery-${returnRow}`, ctx, { thumbnail: await inspect(result.thumbnail_path),
      usedDatabaseResult: result.updated === true });
  }
  {
    const ctx = await setup('row-failure', { native: { createThumbnailFromPath: async () => {
      throw new Error('fixture missing image');
    } } });
    const row = { id: 1, file_path: sourcePaths['plain.png'], file_hash: 'hash', thumbnail_path: null };
    assert.equal(await ctx.store.ensureManagedThumbnailForRow(row), row);
    assert.equal(ctx.updates.length, 0);
    record('failed-row-recovery-keeps-database', ctx, true);
  }
  for (const name of Object.keys(sourcePaths)) {
    assert.equal(digest(await fs.readFile(sourcePaths[name])), sourceHashes[name], 'Original image changed');
  }
  return { cases: results, sourceHashes };
}

async function runThumbnailPathSafetyCases({ createStore, fs, nativeImage, root }) {
  const home = path.join(root, 'path-safety');
  const directory = path.join(home, 'Current', 'thumbnails');
  const sibling = path.join(home, 'Current', 'thumbnails-sibling');
  const sentinelPaths = [path.join(home, 'Current', 'sentinel.jpg'),
    path.join(sibling, 'sentinel.jpg'), path.join(directory, 'untouched.jpg')];
  const sentinel = Buffer.from('Synthetic sentinel that must never be overwritten');
  for (const target of sentinelPaths) {
    const relative = path.relative(root, path.resolve(target));
    assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative));
    await fs.mkdir(path.dirname(target), { recursive: true });
    await fs.writeFile(target, sentinel);
  }
  const unexpectedCalls = [];
  const rejectUnexpected = name => async () => { unexpectedCalls.push(name); throw new Error(`Unexpected unsafe ${name}`); };
  const store = createStore({
    fs: { access: rejectUnexpected('access'), writeFile: rejectUnexpected('writeFile') },
    nativeImage: { createThumbnailFromPath: rejectUnexpected('native decode') },
    getHomePath: () => home, getThumbnailDirectoryPath: () => directory,
    getPhotoDb: () => ({ updateThumbnailPath: () => { throw new Error('Unsafe row must not update the DB'); } }),
    storageDirname: 'Current', legacyStorageDirname: 'Legacy',
  });
  const hashes = ['../sentinel', '..\\sentinel', '../thumbnails-sibling/sentinel',
    '..\\thumbnails-sibling\\sentinel', 'folder/nested', 'folder\\nested',
    path.join(home, 'absolute'), '\\\\server\\share\\sentinel', '/sentinel',
    'C:sentinel', 'hash:alternate-stream'];
  const image = nativeImage.createFromPath(path.join(__dirname, 'photo-metadata', 'plain.png'));
  for (const fileHash of hashes) {
    assert.equal(store.getManagedThumbnailPath(fileHash), null, fileHash);
    assert.equal(await store.createThumbnailFromNativeImage(image, fileHash, { force: true }), null);
    assert.equal(await store.createThumbnail('never-read.png', fileHash, { sourceImage: image, force: true }), null);
    const row = { id: 1, file_path: 'never-read.png', file_hash: fileHash, thumbnail_path: null };
    assert.equal(await store.ensureManagedThumbnailForRow(row), row);
  }
  for (const [fileHash, extension] of [['', ''], ['.', ''], ['..', ''], ['safe', '/../bad']]) {
    assert.equal(store.getManagedThumbnailPath(fileHash, extension), null);
  }
  assert.deepEqual(unexpectedCalls, [], 'Unsafe names must be rejected before any filesystem or codec call');
  for (const target of sentinelPaths) assert.deepEqual(await fs.readFile(target), sentinel);
  return { caseCount: hashes.length + 4, invalidHashCount: hashes.length,
    rejectedBeforeIo: true, preservedSentinelCount: sentinelPaths.length };
}

module.exports = { runThumbnailStoreCases, runThumbnailPathSafetyCases };
