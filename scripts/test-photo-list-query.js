'use strict';

const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createHash } = require('node:crypto');
const list = require('../src/photo-list-query');
const fixture = require('./fixtures/photo-list-query-cases');
const baseline = require('./fixtures/photo-list-query-baseline.json');
const hash = value => createHash('sha256').update(JSON.stringify(value)).digest('hex');
const ids = photos => photos.map(photo => photo.caseId);
const copy = value => JSON.parse(JSON.stringify(value));

function withTimeZone(zone, run) {
  const previous = process.env.TZ;
  process.env.TZ = zone;
  try { return run(); }
  finally {
    if (previous === undefined) delete process.env.TZ;
    else process.env.TZ = previous;
  }
}

test('fixed pre-extraction oracle covers the unchanged photo and query inputs', () => {
  // No renderer code is read to derive expected values during tests.
  assert.equal(baseline.schemaVersion, 1);
  assert.equal(baseline.preExtractionSourceSha256, '55c1b517c67b47b857b4095e921dd87c6ac78e0757601671bcd3719a77815c77');
  assert.equal(hash(fixture.photos), baseline.fixtureSha256);
  assert.equal(hash(fixture.createQueries()), baseline.queriesSha256);
  assert.equal(hash(fixture.dateInputs), baseline.dateInputsSha256);
  assert.equal(hash(fixture.searchInputs), baseline.searchInputsSha256);
  assert.equal(baseline.queryResults.length, 1920);
});

test('1,920 complete filter and sort combinations preserve exact membership and order', () => {
  const photosBefore = JSON.stringify(fixture.photos);
  const queries = fixture.createQueries();
  const queriesBefore = JSON.stringify(queries);
  withTimeZone('Asia/Tokyo', () => queries.forEach((query, index) => {
    const result = list.queryPhotos(fixture.photos, query);
    assert.deepEqual(ids(result), baseline.queryResults[index], `query ${index}: ${JSON.stringify(query)}`);
    assert.notEqual(result, fixture.photos);
    for (const photo of result) assert(fixture.photos.includes(photo), 'Results retain original photo objects');
  }));
  assert.equal(JSON.stringify(fixture.photos), photosBefore);
  assert.equal(JSON.stringify(queries), queriesBefore);
});

test('24 date forms preserve local-time parsing and DST behavior in three time zones', () => {
  for (const [zone, expected] of Object.entries(baseline.timeZones)) {
    withTimeZone(zone, () => {
      assert.deepEqual(fixture.dateInputs.map(list.parsePhotoSortTimestamp), expected, zone);
    });
  }
});

test('invalid takenAt falls back to groupDate; numeric timestamp does not override the existing contract', () => {
  withTimeZone('UTC', () => {
    assert.equal(list.getPhotoSortTimestamp({ takenAt: '2026-09-10 12:00', groupDate: '2020-01-01', takenAtTimestamp: 1 }), 1789041600000);
    assert.equal(list.getPhotoSortTimestamp({ takenAt: 'bad', groupDate: '2026-09-10', takenAtTimestamp: 1 }), 1788998400000);
    assert.equal(list.getPhotoSortTimestamp({ takenAtTimestamp: 1788998400000 }), 0);
    assert.equal(list.getPhotoSortTimestamp(null), 0);
  });
});

test('972 pair comparisons preserve date, ID, locale filename and unknown-direction behavior', () => {
  withTimeZone('Asia/Tokyo', () => {
    for (const [sortOrder, expected] of Object.entries(baseline.comparatorSigns)) {
      const actual = fixture.photos.flatMap(left => fixture.photos.map(right => Math.sign(list.comparePhotos(left, right, sortOrder))));
      assert.deepEqual(actual, expected, sortOrder);
    }
  });
});

test('filename ties remain ascending in both directions and exact ties retain input order', () => {
  const tied = fixture.photos.filter(photo => photo.caseId.startsWith('tie-'));
  const expected = ['tie-alpha', 'tie-duplicate-first', 'tie-duplicate-second', 'tie-zulu'];
  for (const order of ['asc', 'desc']) assert.deepEqual(ids(list.sortPhotos(tied, order)), expected);
  assert.deepEqual(list.sortPhotos(null), []);
  assert.deepEqual(list.queryPhotos(undefined), []);
  assert(Number.isNaN(list.comparePhotos({ id: 'invalid' }, { id: 1 })), 'Malformed IDs retain the comparator fallback rather than acquiring a new order');
});

test('orientation preserves explicit metadata, numeric dimensions and existing null/empty quirks', () => {
  assert.deepEqual(fixture.photos.map(list.getPhotoOrientationTier), baseline.orientationResults);
  assert.equal(list.getPhotoOrientationTier(null), null);
  assert.equal(list.getPhotoOrientationTier({}), null);
  assert.equal(list.getPhotoOrientationTier({ imageWidth: null, imageHeight: null }), 'square');
  assert.equal(list.getPhotoOrientationTier({ imageWidth: '', imageHeight: '' }), 'square');
  assert.equal(list.getPhotoOrientationTier({ orientationTier: 'custom', imageWidth: 1, imageHeight: 2 }), 'custom');
});

test('Unicode search keeps original trimming, NFKC and Japanese-locale case normalization', () => {
  assert.deepEqual(fixture.searchInputs.map(value => ({ trimmed: list.normalizeSearchText(value), normalized: list.normalizeSearchQuery(value) })), baseline.searchResults);
  assert.equal(list.normalizeSearchQuery(' ｶﾌｪ '), 'カフェ');
  assert.equal(list.normalizeSearchQuery('ＣＡＦＥ'), 'cafe');
  assert.equal(list.normalizeSearchQuery('Cafe\u0301'), 'café');
  assert.equal(list.normalizeSearchQuery(123), '');
});

test('world, memo and print note scopes preserve target fields without cross-scope leakage', () => {
  for (const [scope, expected] of Object.entries(baseline.searchTargets)) {
    assert.deepEqual(fixture.photos.map(photo => list.getSearchTargetText(photo, scope)), expected, scope);
  }
  const photo = { worldName: 'manual world', rawWorldName: 'raw name', worldId: 'wrld_abc', memoText: 'memo only', printNoteText: 'print only' };
  assert(list.matchesPhoto(photo, { searchText: 'RAW NAME' }));
  assert(list.matchesPhoto(photo, { searchText: 'WRLD_ABC' }));
  assert(!list.matchesPhoto(photo, { searchText: 'memo only', searchScope: 'world' }));
  assert(list.matchesPhoto(photo, { searchText: 'memo only', searchScope: 'memo' }));
  assert(!list.matchesPhoto(photo, { searchText: 'print only', searchScope: 'memo' }));
  assert(list.matchesPhoto(photo, { searchText: 'print only', searchScope: 'printNote' }));
  assert.equal(list.getSearchTargetText(null), '');
});

test('label AND/OR uses exact normalized keys and ignores malformed or duplicate labels', () => {
  const photo = { photoLabels: [{ normalizedName: 'friend' }, { normalizedName: 'friend' }, {}, null] };
  assert(list.matchesPhoto(photo, { labelFilters: ['friend', 'missing'], labelMode: 'or' }));
  assert(!list.matchesPhoto(photo, { labelFilters: ['friend', 'missing'], labelMode: 'and' }));
  assert(list.matchesPhoto(photo, { labelFilters: ['friend', 'friend'], labelMode: 'and' }));
  assert(!list.matchesPhoto(photo, { labelFilters: ['FRIEND'] }));
  assert(!list.matchesPhoto({ photoLabels: 'invalid' }, { labelFilters: ['friend'] }));
  assert(list.matchesPhoto({}, { labelFilters: [], labelMode: 'and' }));
});

test('favorite and whitespace-only search preserve truthiness and empty-needle behavior', () => {
  assert(list.matchesPhoto({ isFavorite: 1 }, { favoriteOnly: true }));
  assert(list.matchesPhoto({ isFavorite: 'yes' }, { favoriteOnly: true }));
  assert(!list.matchesPhoto({ isFavorite: 0 }, { favoriteOnly: true }));
  assert(!list.matchesPhoto({}, { favoriteOnly: true }));
  assert(list.matchesPhoto({}, { searchText: ' \t ' }));
});

test('one prepared matcher captures query choices and still reads current photo fields', () => {
  const options = { favoriteOnly: true, labelFilters: ['friend'], searchText: 'ＣＡＦＥ', searchScope: 'world' };
  const matcher = list.createMatcher(options);
  options.labelFilters.push('missing');
  options.searchText = 'other';
  const photo = { isFavorite: true, worldName: 'Cafe', photoLabels: [{ normalizedName: 'friend' }] };
  assert(matcher(photo));
  photo.worldName = 'other';
  assert(!matcher(photo));
  photo.worldName = 'Cafe';
  photo.isFavorite = false;
  assert(!matcher(photo));
});

test('label catalog counts each photo once and preserves first label presentation', () => {
  assert.deepEqual(list.getLabelCatalog(fixture.photos), baseline.labelCatalog);
  const catalog = list.getLabelCatalog([
    { photoLabels: [{ normalizedName: 'a', name: 'First', colorHex: '#123456' }, { normalizedName: 'a', name: 'Duplicate' }] },
    { photoLabels: [{ normalizedName: 'a', name: 'Later', colorHex: '#FFFFFF' }, { normalizedName: 'b' }] },
    { photoLabels: null },
  ]);
  assert.deepEqual(catalog, [{ normalizedName: 'b', name: 'b', colorHex: '#6D5EF6', photoCount: 1 },
    { normalizedName: 'a', name: 'First', colorHex: '#123456', photoCount: 2 }]);
});

test('queries and catalog generation support frozen records without mutating inputs', () => {
  const photos = copy(fixture.photos);
  const freeze = value => {
    if (value && typeof value === 'object') { Object.values(value).forEach(freeze); Object.freeze(value); }
    return value;
  };
  freeze(photos);
  const options = freeze(fixture.createQueries()[122]);
  assert.doesNotThrow(() => list.queryPhotos(photos, options));
  assert.doesNotThrow(() => list.getLabelCatalog(photos));
});

test('browser UMD entry runs without DOM, Electron, UI globals or persistent application state', () => {
  const context = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/photo-list-query.js'), 'utf8'), context);
  assert.equal(typeof context.PhotoListQuery.queryPhotos, 'function');
  assert.equal(Object.isFrozen(context.PhotoListQuery), true);
  assert.equal(Object.isFrozen(list), true);
  const query = fixture.createQueries()[0];
  assert.deepEqual(copy(ids(context.PhotoListQuery.queryPhotos(fixture.photos, query))), baseline.queryResults[0]);
  assert.deepEqual(copy(context.PhotoListQuery.getLabelCatalog(fixture.photos)), baseline.labelCatalog);
});
