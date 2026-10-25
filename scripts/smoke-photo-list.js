// Actual renderer list controls and main/preload/SQLite data, with fixed pre-extraction expectations.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { app, BrowserWindow, nativeImage, session } = require('electron');

const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-photo-list-'));
const OUTPUT_DIR = path.resolve(__dirname, '..', 'out', 'smoke-photo-list', path.basename(ROOT));
const PROFILE = path.join(ROOT, 'Profile');
const APPDATA = path.join(PROFILE, 'AppData', 'Roaming');
const LOCALAPPDATA = path.join(PROFILE, 'AppData', 'Local');
const TEMP = path.join(ROOT, 'Temp');
for (const dir of [OUTPUT_DIR, PROFILE, APPDATA, LOCALAPPDATA, TEMP]) fs.mkdirSync(dir, { recursive: true });
process.env.APPDATA = process.env.WORLDSHOT_APPDATA_ROOT = APPDATA;
process.env.LOCALAPPDATA = process.env.WORLDSHOT_LOCALAPPDATA_ROOT = LOCALAPPDATA;
process.env.USERPROFILE = process.env.WORLDSHOT_HOME_ROOT = PROFILE;
process.env.TEMP = process.env.TMP = TEMP;
app.setPath('home', PROFILE);
app.setPath('appData', APPDATA);
app.setPath('userData', path.join(APPDATA, 'WorldShot Log'));
app.setPath('temp', TEMP);
const sha256 = value => createHash('sha256').update(value).digest('hex');
const WORLD_ID = 'wrld_11111111-1111-1111-1111-111111111111';
const report = { ok: false, root: ROOT, checks: [], blockedRequests: [],
  rendererSourceSha256: sha256(fs.readFileSync(path.resolve(__dirname, '..', 'src', 'renderer.js'))) };
const fixtures = [
  { date: '2026-09-01', time: '10:00:00', world: 'ＳＡＫＵＲＡ ガーデン', manual: '桜庭園', worldId: WORLD_ID,
    orientation: 'landscape', favorite: true, labels: ['旅', '赤'], memo: 'ＣＡＦＥ café 夜景' },
  { date: '2026-09-01', time: '10:00:00', world: 'Sunset Café', orientation: 'landscape', labels: ['旅'], memo: '猫と星' },
  { date: '2026-09-02', time: '12:30:00', world: '海辺', orientation: 'portrait', favorite: true,
    labels: ['旅', '青'], printNote: 'ＦＲＩＥＮＤＳ 集合' },
  { date: '2026-09-03', time: '08:00:00', world: '猫カフェ ＣＡＦＥ', orientation: 'square', labels: ['赤'], printNote: '星空 ２０２６' },
  { date: '2026-09-04', invalidDate: true, world: '', orientation: 'landscape' },
  { date: '2026-09-03', time: '08:00:00', world: '桜庭園 Original', manual: '手動 名称', orientation: 'portrait',
    favorite: true, labels: ['青'], memo: 'normalized cafe\u0301' },
  { date: '2026-08-02', time: '15:00:00', world: 'ＳＡＫＵＲＡ ガーデン', worldId: WORLD_ID,
    orientation: 'square', favorite: true, labels: ['旅'] },
  { date: '2026-08-01', time: '09:00:00', world: 'Sea', orientation: 'landscape' },
  { date: '2025-12-31', time: '23:59:59', world: '年越し', orientation: 'portrait', labels: ['赤'] },
];

function seed() {
  const dataDir = path.join(app.getPath('userData'), 'data');
  fs.mkdirSync(dataDir, { recursive: true });
  const db = require('../src/db').initDatabase(path.join(dataDir, 'app.sqlite'));
  try {
    return fixtures.map((fixture, index) => {
      const id = index + 1;
      const [year, month, day] = fixture.date.split('-').map(Number);
      const [width, height] = fixture.orientation === 'portrait' ? [48, 64]
        : fixture.orientation === 'square' ? [48, 48] : [64, 48];
      const pixels = Buffer.alloc(width * height * 4);
      for (let i = 0; i < pixels.length; i += 4) pixels.set([id * 20, 80, 180, 255], i);
      const filePath = path.join(ROOT, `fixture-${id}.png`);
      fs.writeFileSync(filePath, nativeImage.createFromBitmap(pixels, { width, height }).toPNG());
      const fileHash = sha256(fs.readFileSync(filePath));
      const photo = db.insertOrUpdatePhoto({ filePath, fileName: path.basename(filePath), fileHash,
        takenAt: fixture.invalidDate ? 'invalid-date' : `${fixture.date.replaceAll('-', '/')} ${fixture.time}`,
        takenAtTimestamp: new Date(`${fixture.date}T${fixture.time || '00:00:00'}`).getTime(),
        groupDate: fixture.date, year, month, day, worldId: fixture.worldId || null,
        worldName: fixture.world || null, worldNameManual: fixture.manual || null,
        worldUrl: fixture.worldId ? `https://vrchat.com/home/world/${fixture.worldId}` : null,
        thumbnailPath: filePath, imageWidth: width, imageHeight: height,
        resolutionTier: 'other', orientationTier: fixture.orientation,
        printNoteText: fixture.printNote || '', memoText: fixture.memo || null,
        createdAt: '2026-09-10T00:00:00.000Z', updatedAt: '2026-09-10T00:00:00.000Z' });
      assert.equal(photo.id, id);
      if (fixture.favorite) db.updateFavoriteStatus(id, true);
      if (fixture.manual) db.updateManualWorldName(id, fixture.manual);
      if (fixture.memo) db.updatePhotoMemo(id, fixture.memo);
      if (fixture.labels) db.replacePhotoTags(id, fixture.labels);
      return { id, filePath, fileHash };
    });
  } finally { db.close(); }
}

let win;
const inPage = (fn, ...args) => win.webContents.executeJavaScript(`(${fn.toString()})(...${JSON.stringify(args)})`);
async function until(label, predicate) {
  const deadline = Date.now() + 10000;
  while (Date.now() < deadline) {
    if (await predicate()) return;
    await new Promise(resolve => setTimeout(resolve, 30));
  }
  throw new Error(`Timed out: ${label}`);
}
async function snapshot() {
  return inPage(() => ({
    cards: [...document.querySelectorAll('#month-gallery-list .photo-card')].map(card => Number(card.dataset.photoId)),
    current: currentPhotos.map(photo => photo.id), total: allCurrentMonthPhotos.length,
    count: document.getElementById('current-month-count').textContent,
    selection: currentSelection, selectedIds: [...selectedPhotoIds].sort((a, b) => a - b),
    selectedCards: [...document.querySelectorAll('#month-gallery-list .photo-card.is-selected')].map(card => Number(card.dataset.photoId)),
    selectionMode: isSelectionMode, sort: currentPhotoSortOrder,
    filter: activeWorldNameFilter, scope: activeToolbarSearchScope, draftScope: draftToolbarSearchScope,
    emptyText: document.getElementById('month-gallery-empty').textContent,
    emptyVisible: getComputedStyle(document.getElementById('month-gallery-empty')).display !== 'none',
  }));
}
async function check(name, expected, action = async () => {}, extra = () => {}) {
  try {
    await action();
    await until(name, async () => {
      const value = await snapshot();
      return JSON.stringify(value.cards) === JSON.stringify(expected) && JSON.stringify(value.current) === JSON.stringify(expected);
    });
    const value = await snapshot();
    assert(value.count.startsWith(`${expected.length}枚`), `${name}: count text ${value.count}`);
    assert.equal(value.emptyVisible, expected.length === 0, `${name}: empty-state visibility`);
    await extra(value);
    report.checks.push({ name, ok: true, ...value });
    console.log(`PASS ${name}`);
  } catch (error) {
    report.checks.push({ name, ok: false, error: error.message, actual: await snapshot() });
    throw error;
  }
}
const click = selector => inPage(value => {
  const button = document.querySelector(value);
  if (!button || button.disabled || button.getClientRects().length === 0) throw new Error(`Control unavailable: ${value}`);
  button.click();
}, selector);
async function search(text, { scope, submit = true, composing = false } = {}) {
  if (scope) {
    await click('.toolbar-search-scope-button');
    await click(`[data-toolbar-search-scope="${scope}"]`);
  }
  await inPage((value, submit, composing) => {
    const input = document.getElementById('world-name-filter-input');
    input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    if (submit) input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, isComposing: composing }));
  }, text, submit, composing);
}
async function orientation(value) {
  await click('#orientation-filter-btn');
  await click(`[data-orientation-filter="${value}"]`);
}
async function label(value) {
  if (await inPage(() => document.getElementById('photo-label-filter-btn').getAttribute('aria-expanded') !== 'true')) {
    await click('#photo-label-filter-btn');
  }
  await click(`[data-photo-label-filter="${value}"]`);
}
const clearSearch = () => click('.toolbar-search-clear-button');

async function main() {
  await app.whenReady();
  session.defaultSession.webRequest.onBeforeRequest({ urls: ['http://*/*', 'https://*/*', 'ws://*/*', 'wss://*/*'] },
    (details, callback) => { report.blockedRequests.push(details.url); callback({ cancel: true }); });
  globalThis.fetch = async input => { report.blockedRequests.push(String(input)); throw new Error('Network disabled in photo list smoke'); };
  app.on('browser-window-created', (_event, window) => { window.hide(); window.webContents.setBackgroundThrottling(false); });
  const sources = seed();
  require('../src/index');
  for (const key of ['home', 'documents', 'pictures', 'userData', 'sessionData', 'temp']) {
    const relative = path.relative(ROOT, app.getPath(key));
    assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative), `Unisolated ${key}`);
  }
  await until('main window', () => { win = BrowserWindow.getAllWindows()[0]; return win; });
  await until('renderer initialization', () => inPage(() => typeof currentPhotos !== 'undefined' &&
    !document.querySelector('.app')?.classList.contains('is-app-initializing')));
  await inPage(() => {
    window.photoListSmokeErrors = [];
    window.addEventListener('error', event => photoListSmokeErrors.push(event.message));
    window.addEventListener('unhandledrejection', event => photoListSmokeErrors.push(String(event.reason)));
  });
  await check('latest month defaults to newest date, then descending ID', [5, 6, 4, 3, 2, 1], undefined,
    value => { assert.equal(value.count, '6枚'); assert.equal(value.selection.month, 9); });
  await check('oldest first reverses date and ID tie', [1, 2, 3, 4, 6, 5], () => click('#photo-sort-btn'));
  await check('newest first restores order', [5, 6, 4, 3, 2, 1], () => click('#photo-sort-btn'));
  await check('typing draft alone does not apply search', [5, 6, 4, 3, 2, 1], () => search('sakura', { submit: false }),
    value => assert.equal(value.filter, ''));
  await check('IME composing Enter does not submit', [5, 6, 4, 3, 2, 1], () => search('sakura', { composing: true }),
    value => assert.equal(value.filter, ''));
  await check('fullwidth raw world name matches folded search despite manual display name', [1], () => search('  SaKuRa  '),
    value => assert.equal(value.filter, 'SaKuRa'));
  await check('manual and original world names are both searchable', [6, 1], () => search('桜庭園'));
  await check('world ID is searchable', [1], () => search(WORLD_ID));
  await check('world name search normalizes fullwidth Latin', [4], () => search('ＣＡＦＥ'));
  await check('changing draft search scope alone keeps applied world scope', [4], async () => {
    await click('.toolbar-search-scope-button'); await click('[data-toolbar-search-scope="memo"]');
  }, value => { assert.equal(value.scope, 'world'); assert.equal(value.draftScope, 'memo'); });
  await check('same query submission applies changed memo scope without stripping accents', [1], () => search('ＣＡＦＥ'));
  await check('memo search normalizes decomposed accent and case', [6, 1], () => search('CAFÉ', { scope: 'memo' }));
  await check('print-note search folds fullwidth Latin', [3], () => search('friends', { scope: 'printNote' }));
  await check('print-note search folds fullwidth digits', [4], () => search('2026'));
  await check('unmatched search shows zero-count summary and empty message', [], () => search('no-match-fixture'), value => {
    assert(value.count.includes('/ 全6枚')); assert(value.emptyText.includes('一致する写真はありません'));
  });
  await check('clearing search restores original order and count', [5, 6, 4, 3, 2, 1], clearSearch);
  await check('favorite control filters existing favorites', [6, 3, 1], () => click('#favorite-filter-btn'));
  await check('favorite and portrait filters compose', [6, 3], () => orientation('portrait'));
  await check('favorite and square filters show empty state', [], () => orientation('square'));
  await check('square filter alone selects square source', [4], () => click('#favorite-filter-btn'));
  await check('landscape filter selects all landscape sources', [5, 2, 1], () => orientation('landscape'));
  await check('clearing orientation restores all photos', [5, 6, 4, 3, 2, 1], () => orientation('all'));
  await check('one label filters by normalized label identity', [3, 2, 1], () => label('旅'));
  await check('OR combines labels', [6, 3, 2, 1], () => label('青'));
  await check('AND intersects labels', [3], () => click('[data-photo-label-filter-mode="and"]'));
  await check('AND label and world search compose', [3], () => search('海辺', { scope: 'world' }));
  await check('combined filters preserve empty result', [], () => search('桜庭園'));
  await check('clear search keeps label constraints', [3], clearSearch);
  await check('clear labels restores complete list', [5, 6, 4, 3, 2, 1], () => label('__all__'));
  await check('selection mode and card clicks select expected IDs', [5, 6, 4, 3, 2, 1], async () => {
    await click('#selection-mode-btn');
    await click('#month-gallery-list .photo-card[data-photo-id="1"]');
    await click('#month-gallery-list .photo-card[data-photo-id="6"]');
  }, value => { assert.equal(value.selectionMode, true); assert.deepEqual(value.selectedIds, [1, 6]); assert.deepEqual(value.selectedCards, [6, 1]); });
  await check('filtering retains selected IDs while hiding unmatched cards', [6, 3], () => orientation('portrait'),
    value => { assert.deepEqual(value.selectedIds, [1, 6]); assert.deepEqual(value.selectedCards, [6]); });
  await check('clearing filter restores visible selection markings', [5, 6, 4, 3, 2, 1], () => orientation('all'),
    value => assert.deepEqual(value.selectedCards, [6, 1]));
  await check('leaving selection mode clears selection', [5, 6, 4, 3, 2, 1], () => click('#selection-mode-btn'),
    value => { assert.equal(value.selectionMode, false); assert.deepEqual(value.selectedIds, []); });
  await check('month sidebar selects previous month with fixed order', [7, 8], () => click('[data-year="2026"] .month-button[data-month="8"]'),
    value => { assert.equal(value.count, '2枚'); assert.equal(value.selection.month, 8); });
  await check('month selection preserves active favorite filter', [7], () => click('#favorite-filter-btn'));
  await check('year sidebar combines months and preserves active filter', [6, 3, 1, 7], () => click('[data-year="2026"] .year-button'),
    value => { assert.equal(value.total, 8); assert.equal(value.selection.mode, 'year'); });
  await check('year scope without filter has all year photos', [5, 6, 4, 3, 2, 1, 7, 8], () => click('#favorite-filter-btn'));
  await check('older year sidebar excludes current year', [9], () => click('[data-year="2025"] .year-button'),
    value => assert.equal(value.selection.year, 2025));
  await check('return to latest month restores unchanged complete list', [5, 6, 4, 3, 2, 1], () => click('[data-year="2026"] .month-button[data-month="9"]'));
  const catalog = await inPage(() => getCurrentMonthPhotoLabelCatalog().map(label => ({ name: label.name, count: label.photoCount })));
  assert.deepEqual(catalog, [{ name: '青', count: 2 }, { name: '赤', count: 2 }, { name: '旅', count: 3 }]);
  assert.deepEqual(await inPage(() => photoListSmokeErrors), []);
  for (const source of sources) assert.equal(sha256(fs.readFileSync(source.filePath)), source.fileHash, 'Original fixture changed');
  report.ok = true;
  report.catalog = catalog;
  report.sources = sources;
  console.log(`Photo list smoke passed (${report.checks.length} scenarios): ${OUTPUT_DIR}`);
}
const watchdog = setTimeout(() => { console.error('Photo list smoke exceeded 90 seconds'); app.exit(1); }, 90000);
main().catch(error => { report.error = error.stack || String(error); console.error(report.error); process.exitCode = 1; })
  .finally(() => {
    clearTimeout(watchdog);
    fs.writeFileSync(path.join(OUTPUT_DIR, 'result.json'), JSON.stringify(report, null, 2));
    app.exit(report.ok ? 0 : 1);
  });
