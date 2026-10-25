// Run with the Electron CLI wrapper to use the application's better-sqlite3 ABI:
// node node_modules/electron/cli.js scripts/benchmark-backup-restore.js --label=baseline
// No application startup, files containing photos, network, or user DB is used.
// All databases are :memory:. Only the benchmark report is written to out/.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { createHash } = require('node:crypto');
const { performance } = require('node:perf_hooks');
const { initDatabase } = require('../src/db');

const PHOTO_COUNTS = [5000, 20000];
const WARMUP_RUNS = 2;
const MEASURED_RUNS = 5;
const FIXED_TIME = '2026-09-09T00:00:00.000Z';
const label = (process.argv.find(value => value.startsWith('--label=')) || '--label=baseline').slice(8);
const runId = `${new Date().toISOString().replace(/[:.]/g, '-')}-${process.pid}`;
const outputDirectory = path.resolve(__dirname, '..', 'out', 'benchmark-backup-restore', runId);
const reportPath = path.join(outputDirectory, 'result.json');
const report = {
  ok: false,
  label,
  runId,
  methodology: {
    photoCounts: PHOTO_COUNTS,
    warmupRuns: WARMUP_RUNS,
    measuredRuns: MEASURED_RUNS,
    database: ':memory:',
    timing: 'Main-style preflight validation and DB restore are timed separately. Restore includes its own validation, deletion, and insertion transaction.',
    excluded: ['fixture generation', 'all-row verification', 'JSON parsing/serialization', 'backup file I/O', 'DB initialization'],
    verification: 'Every warmup and measured restore must match every field of every row in all five tables, including defaults for omitted columns.',
    workload: 'Repeated full replacement of an already populated DB after warmup; two labels per photo, one world cache row per 20 photos, 16 tracked folders.',
    scope: 'In-memory synthetic DB benchmark; these values do not estimate full UI restore time or disk-backed performance.',
  },
  environment: {
    platform: process.platform,
    arch: process.arch,
    osRelease: os.release(),
    cpuModel: os.cpus()[0]?.model || '',
    logicalCpus: os.cpus().length,
    versions: process.versions,
    appVersion: require('../package.json').version,
    dbSourceSha256: createHash('sha256').update(fs.readFileSync(path.join(__dirname, '..', 'src', 'db.js'))).digest('hex'),
  },
  results: [],
};
let loggingEnabled = true;
let finished = false;

function log(message) {
  if (!loggingEnabled) return;
  try { process.stdout.write(`${message}\n`); }
  catch (error) { handleOutputError(error); }
}

function finish(error = null) {
  if (finished) return;
  finished = true;
  if (error) {
    report.ok = false;
    report.error = error.stack || String(error);
  }
  try {
    fs.mkdirSync(outputDirectory, { recursive: true });
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
  } catch (writeError) {
    report.ok = false;
    log(`Cannot save benchmark report: ${writeError.stack || writeError}`);
  }
  if (report.error) log(report.error);
  log(`Backup restore benchmark report: ${reportPath}`);
  process.exit(report.ok ? 0 : 1);
}

function handleOutputError(error) {
  if (error?.code === 'EPIPE') loggingEnabled = false;
  else finish(error);
}
process.stdout.on('error', handleOutputError);
process.stderr.on('error', handleOutputError);
process.on('uncaughtException', finish);
process.on('unhandledRejection', reason => finish(reason instanceof Error ? reason : new Error(String(reason))));

const PHOTO_OPTIONAL_DEFAULTS = {
  world_id: null,
  world_name: null,
  world_name_manual: null,
  world_url: null,
  world_identity_source: 'legacy',
  thumbnail_path: null,
  image_width: null,
  image_height: null,
  resolution_tier: null,
  orientation_tier: null,
  print_note_text: null,
  memo_text: null,
  is_favorite: 0,
};
const WORLD_OPTIONAL_DEFAULTS = {
  source_url: null,
  world_name_official: null,
  world_description: null,
  world_tags_json: null,
  author_id: null,
  author_name: null,
  release_status: null,
  image_url: null,
  thumbnail_image_url: null,
  fetch_error: null,
  fetched_at: null,
};

function worldId(index) {
  return `wrld_00000000-0000-4000-8000-${String(index).padStart(12, '0')}`;
}

function countSummary(tables) {
  return {
    photoCount: tables.photos.length,
    trackedFolderCount: tables.tracked_folders.length,
    worldCacheCount: tables.world_metadata_cache.length,
    tagCount: tables.tags.length,
  };
}

function describeShapes(tables) {
  return Object.fromEntries(Object.entries(tables).map(([table, rows]) => {
    const shapes = new Map();
    for (const row of rows) {
      const key = JSON.stringify(Object.keys(row).sort());
      shapes.set(key, (shapes.get(key) || 0) + 1);
    }
    return [table, {
      rowCount: rows.length,
      distinctColumnSets: shapes.size,
      shapes: Array.from(shapes, ([columns, rowCount]) => ({ columns: JSON.parse(columns), rowCount })),
    }];
  }));
}

function createFixture(photoCount) {
  const tables = { photos: [], tracked_folders: [], world_metadata_cache: [], tags: [], photo_tags: [] };
  const tagCount = 80;
  const worldCount = Math.ceil(photoCount / 20);
  for (let id = 1; id <= tagCount; id += 1) {
    const tag = {
      id, name: `Label ${id}`, normalized_name: `label ${id}`,
      color_hex: id % 3 === 0 ? null : '#14B8A6',
      created_at: FIXED_TIME, updated_at: FIXED_TIME,
    };
    if (id % 3 === 1) delete tag.color_hex;
    tables.tags.push(tag);
  }
  for (let id = 1; id <= 16; id += 1) {
    tables.tracked_folders.push({
      id, folder_path: `synthetic/collection-${id}`,
      created_at: FIXED_TIME, updated_at: FIXED_TIME, last_imported_at: FIXED_TIME,
    });
  }
  for (let id = 1; id <= worldCount; id += 1) {
    const world = {
      world_id: worldId(id), source_url: `https://example.invalid/world/${id}`,
      world_name_official: `Synthetic world ${id}`,
      world_description: `Synthetic description ${id}: a calm landscape for a photo collection.`,
      world_tags_json: JSON.stringify(['synthetic', `theme-${id % 12}`]),
      author_id: `usr_synthetic_${id % 50}`, author_name: `Author ${id % 50}`,
      release_status: 'public', image_url: null, thumbnail_image_url: null,
      fetch_status: id % 2 === 0 ? 'success' : 'partial', fetch_error: null,
      fetched_at: FIXED_TIME, last_attempted_at: FIXED_TIME,
      created_at: FIXED_TIME, updated_at: FIXED_TIME,
    };
    if (id % 2 === 1) {
      delete world.world_description;
      delete world.image_url;
      delete world.thumbnail_image_url;
      delete world.fetched_at;
    }
    tables.world_metadata_cache.push(world);
  }
  for (let id = 1; id <= photoCount; id += 1) {
    const worldIndex = Math.floor((id - 1) / 20) + 1;
    let photo = {
      id,
      file_path: `synthetic/collection-${id % 16 + 1}/photo-${id}.png`,
      file_name: `photo-${id}.png`,
      file_hash: createHash('sha256').update(`synthetic-photo-${id}`).digest('hex'),
      taken_at: '2026/09/09 09:00:00', taken_at_timestamp: Date.parse(FIXED_TIME) + id * 1000,
      group_date: '2026-09-09', year: 2026, month: 9, day: 9,
      world_id: worldId(worldIndex), world_name: `Synthetic world ${worldIndex}`,
      world_name_manual: id % 11 === 0 ? `Manual name ${id}` : null,
      world_url: `https://example.invalid/world/${worldIndex}`,
      world_identity_source: id % 11 === 0 ? 'manual' : 'automatic',
      thumbnail_path: `synthetic/thumbnails/${id}.jpg`,
      image_width: id % 3 === 0 ? 2160 : 3840,
      image_height: id % 3 === 0 ? 3840 : 2160,
      resolution_tier: '4K', orientation_tier: id % 3 === 0 ? 'portrait' : 'landscape',
      print_note_text: id % 7 === 0 ? `Print note ${id}` : null,
      memo_text: `Memo ${id}: synthetic metadata for a repeatable restore benchmark.`,
      created_at: FIXED_TIME, updated_at: FIXED_TIME, is_favorite: id % 5 === 0 ? 1 : 0,
    };
    if (id % 4 === 1) {
      delete photo.memo_text;
      delete photo.thumbnail_path;
      photo.print_note_text = null;
    } else if (id % 4 === 2) {
      // Legacy/defaulted row shape must not inherit a preceding row's bindings.
      delete photo.is_favorite;
      delete photo.world_identity_source;
      delete photo.world_name_manual;
      delete photo.print_note_text;
      delete photo.memo_text;
    } else if (id % 4 === 3) {
      // Same column set as a full row, with a different JavaScript key order.
      photo = Object.fromEntries(Object.entries(photo).reverse());
      photo.memo_text = null;
    }
    tables.photos.push(photo);
    const labelIds = [(id - 1) % tagCount + 1, (id + 17) % tagCount + 1].sort((a, b) => a - b);
    for (const tagId of labelIds) {
      tables.photo_tags.push({ photo_id: id, tag_id: tagId, created_at: FIXED_TIME });
    }
  }
  const expectedTables = {
    photos: tables.photos.map(row => ({ ...PHOTO_OPTIONAL_DEFAULTS, ...row })),
    tracked_folders: tables.tracked_folders.map(row => ({ ...row })),
    world_metadata_cache: tables.world_metadata_cache.map(row => ({ ...WORLD_OPTIONAL_DEFAULTS, ...row })),
    tags: tables.tags.map(row => ({ color_hex: null, ...row })),
    photo_tags: tables.photo_tags.map(row => ({ ...row })),
  };
  return {
    input: { schemaVersion: 1, tables, counts: countSummary(tables) },
    expected: { schemaVersion: 1, tables: expectedTables, counts: countSummary(expectedTables) },
    shapes: describeShapes(tables),
  };
}

function verifyRestore(db, result, expected) {
  const actual = db.createBackupSnapshot();
  assert.deepEqual(actual, expected, 'Every restored row and optional/defaulted column must match');
  assert.deepEqual(result.afterCounts, expected.counts);
  assert.equal(result.restoredPhotoCount, expected.tables.photos.length);
  assert.equal(result.restoredTrackedFolderCount, expected.tables.tracked_folders.length);
  assert.equal(result.restoredWorldCacheCount, expected.tables.world_metadata_cache.length);
  assert.equal(result.restoredTagCount, expected.tables.tags.length);
  assert.equal(result.restoredPhotoTagCount, expected.tables.photo_tags.length);
}

function distribution(samples) {
  const sorted = [...samples].sort((a, b) => a - b);
  return { medianMs: sorted[Math.floor(sorted.length / 2)], minMs: sorted[0], maxMs: sorted.at(-1), samplesMs: samples };
}

function benchmark(photoCount) {
  const fixture = createFixture(photoCount);
  const inputBefore = JSON.stringify(fixture.input);
  const db = initDatabase(':memory:');
  const samples = [];
  try {
    for (let index = 0; index < WARMUP_RUNS + MEASURED_RUNS; index += 1) {
      const validationStart = performance.now();
      db.validateBackupSnapshot(fixture.input);
      const preflightMs = performance.now() - validationStart;
      const restoreStart = performance.now();
      const result = db.restoreBackupSnapshot(fixture.input);
      const restoreMs = performance.now() - restoreStart;
      const verificationStart = performance.now();
      verifyRestore(db, result, fixture.expected);
      assert.equal(JSON.stringify(fixture.input), inputBefore, 'Restore must not mutate the supplied backup');
      const verificationMs = performance.now() - verificationStart;
      if (index >= WARMUP_RUNS) samples.push({ preflightMs, restoreMs, verificationMs });
    }
    const result = {
      photoCount,
      totalRows: Object.values(fixture.input.tables).reduce((sum, rows) => sum + rows.length, 0),
      columnSets: fixture.shapes,
      verifiedRestores: WARMUP_RUNS + MEASURED_RUNS,
      verification: 'all rows and input immutability passed',
      preflightValidation: distribution(samples.map(sample => sample.preflightMs)),
      dbRestoreIncludingValidation: distribution(samples.map(sample => sample.restoreMs)),
      mainStyleCombinedDbWork: distribution(samples.map(sample => sample.preflightMs + sample.restoreMs)),
      verificationOutsideTimedRestore: distribution(samples.map(sample => sample.verificationMs)),
      samples,
      memoryUsageAfterCase: process.memoryUsage(),
    };
    report.results.push(result);
    log(`${photoCount} photos / ${result.totalRows} rows: restore median ${result.dbRestoreIncludingValidation.medianMs.toFixed(2)}ms, preflight median ${result.preflightValidation.medianMs.toFixed(2)}ms; all ${result.verifiedRestores} restores verified`);
  } finally {
    db.close();
  }
}

try {
  assert.ok(process.versions.electron, 'Run with the Electron CLI wrapper, not plain Node');
  const started = performance.now();
  for (const photoCount of PHOTO_COUNTS) benchmark(photoCount);
  report.elapsedMs = performance.now() - started;
  report.ok = true;
  finish();
} catch (error) {
  finish(error);
}
