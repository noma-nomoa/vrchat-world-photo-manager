'use strict';

// Run with: node scripts/benchmark-photo-editor-history.js --label=baseline
// Uses the real session controller with already-normalized synthetic snapshots.
// No Electron startup, image decoding, network, user files, or source changes.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { createHash } = require('node:crypto');
const { performance } = require('node:perf_hooks');
const { createController } = require('../src/photo-editor-session');
const { getDefaultPhotoEditorCurveState } = require('../src/photo-editor-color-core');

const MASK_SIZES = [0, 256 * 1024, 1024 * 1024];
const HISTORY_LIMIT = 40;
const OPERATIONS_PER_RUN = 64;
const WARMUP_RUNS = 3;
const MEASURED_RUNS = 7;
const label = (process.argv.find(value => value.startsWith('--label=')) || '--label=baseline').slice(8);
const runId = `${new Date().toISOString().replace(/[:.]/g, '-')}-${process.pid}`;
const outputDirectory = path.resolve(__dirname, '..', 'out', 'benchmark-photo-editor-history', runId);
const reportPath = path.join(outputDirectory, 'result.json');

const sha256 = value => createHash('sha256').update(value).digest('hex');
const jsonEqual = (left, right) => JSON.stringify(left || null) === JSON.stringify(right || null);

// Benchmark-only hypothesis, deliberately not wired into the session controller.
// The contract is normalized plain snapshots, without getters or toJSON hooks.
// Retaining the maskDataUrl key with an empty value preserves JSON key order.
function compareNormalizedSnapshotsWithoutEncodingMask(left, right) {
  const leftMask = left?.subjectMask?.maskDataUrl;
  const rightMask = right?.subjectMask?.maskDataUrl;
  if (typeof leftMask !== 'string' || typeof rightMask !== 'string') return jsonEqual(left, right);
  if (leftMask !== rightMask) return false;
  return jsonEqual(
    { ...left, subjectMask: { ...left.subjectMask, maskDataUrl: '' } },
    { ...right, subjectMask: { ...right.subjectMask, maskDataUrl: '' } }
  );
}

// Clone normalized containers while reusing immutable primitive strings.
// JSON/structuredClone would add large-string copying absent from mask normalization.
function cloneNormalizedSnapshot(value) {
  if (Array.isArray(value)) return value.map(cloneNormalizedSnapshot);
  if (!value || typeof value !== 'object') return value;
  return Object.fromEntries(Object.entries(value).map(([key, item]) => [
    key, cloneNormalizedSnapshot(item),
  ]));
}

function createMaskDataUrl(byteLength) {
  if (!byteLength) return '';
  const prefix = 'data:image/png;base64,';
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  const value = prefix + alphabet.repeat(Math.ceil(byteLength / alphabet.length)).slice(0, byteLength - prefix.length);
  assert.equal(Buffer.byteLength(value), byteLength);
  // Force any initial rope flattening outside the measured runs.
  sha256(value);
  return value;
}

function createSnapshot(maskDataUrl) {
  return {
    values: {
      brightness: 0, exposure: 0, contrast: 0, highlights: 0, whites: 0, shadows: 0,
      blacks: 0, gamma: 0, temperature: 0, tint: 0, saturation: 0, vibrance: 0,
      clarity: 0, texture: 0, sharpness: 0, denoise: 0, fade: 0, grain: 0, vignette: 0,
    },
    adjustmentTarget: 'whole',
    crop: { preset: 'original', zoomMode: 'offset', zoom: 0, offsetX: 0, offsetY: 0,
      rotation: 0, flipX: false, flipY: false, tilt: 0 },
    blur: { mode: 'none', amount: 0, centerX: 0.5, centerY: 0.5,
      radius: 0.3, outerRadius: 0.6, space: 'source', isConfirmed: false },
    curve: getDefaultPhotoEditorCurveState(),
    autoEnhance: { enabled: false, originalValuesBeforeAuto: null,
      generatedValues: null, strength: 100, presetKey: '' },
    textOverlays: [{ id: 'text-1', enabled: true, text: 'History benchmark', x: 0.5,
      y: 0.2, rotation: 0, fontKey: 'system', size: 64, color: '#ffffff', weight: '700',
      strokeType: 'outline', strokeWidth: 4, strokeColor: '#111827', maskMode: 'normal',
      fillTransparent: false, space: 'source', letterSpacing: 0 }],
    activeTextId: 'text-1',
    imageOverlays: [{ id: 'image-1', assetId: 'asset-1', enabled: true, x: 0.4, y: 0.5,
      width: 0.2, height: 0.1, rotation: 0, opacity: 1, blendMode: 'source-over',
      maskMode: 'normal', space: 'source' }],
    activeImageOverlayId: 'image-1',
    subjectMask: { enabled: Boolean(maskDataUrl), status: maskDataUrl ? 'ready' : 'none',
      source: maskDataUrl ? 'manual' : 'none', modelId: null, maskDataUrl,
      width: maskDataUrl ? 1024 : 0, height: maskDataUrl ? 1024 : 0,
      feather: 0, expand: 0, invert: false, opacity: 0.55, showOverlay: false,
      createdAt: null, updatedAt: null, errorMessage: '' },
    rulerGuides: { horizontal: [0.25, 0.75], vertical: [0.5] },
    exportSettings: { format: 'png', quality: 92, maxEdge: 0 },
    masks: [{ id: 'mask-1', type: 'fill', shape: 'polygon', strength: 40, color: '#111827',
      points: [{ x: 0.1, y: 0.1 }, { x: 0.2, y: 0.1 }, { x: 0.2, y: 0.2 }], space: 'source' }],
    maskShape: 'rect', maskStrengths: { blur: 45, mosaic: 45, fill: 100 },
    blurStrength: 45, fillColor: '#111827',
  };
}

function distribution(samplesMs) {
  const sorted = [...samplesMs].sort((left, right) => left - right);
  return { medianMs: sorted[Math.floor(sorted.length / 2)], minMs: sorted[0],
    maxMs: sorted.at(-1), samplesMs };
}

function verifyHistory(controller, owner, changed, maskDataUrl) {
  assert.deepEqual(controller.getHistoryStatus(), { canUndo: changed, canRedo: false });
  if (!changed) {
    controller.undo(); controller.redo();
    assert.equal(owner.snapshot.values.brightness, 0);
    return;
  }
  const minimumValue = OPERATIONS_PER_RUN - HISTORY_LIMIT;
  for (let value = OPERATIONS_PER_RUN - 1; value >= minimumValue; value -= 1) {
    controller.undo();
    assert.equal(owner.snapshot.values.brightness, value);
    assert.equal(owner.snapshot.subjectMask.maskDataUrl, maskDataUrl);
  }
  assert.deepEqual(controller.getHistoryStatus(), { canUndo: false, canRedo: true });
  controller.undo();
  assert.equal(owner.snapshot.values.brightness, minimumValue, 'History must stop at its 40-entry limit');
  controller.beginMutation(); controller.commitMutation();
  assert.equal(controller.getHistoryStatus().canRedo, true, 'No-op must preserve the redo branch');
  for (let value = minimumValue + 1; value <= OPERATIONS_PER_RUN; value += 1) {
    controller.redo();
    assert.equal(owner.snapshot.values.brightness, value);
    assert.equal(owner.snapshot.subjectMask.maskDataUrl, maskDataUrl);
  }
  assert.deepEqual(controller.getHistoryStatus(), { canUndo: true, canRedo: false });
  controller.redo();
  assert.equal(owner.snapshot.values.brightness, OPERATIONS_PER_RUN);
}

function benchmarkController(fixture, changed) {
  const samples = [];
  for (let run = 0; run < WARMUP_RUNS + MEASURED_RUNS; run += 1) {
    const owner = { snapshot: cloneNormalizedSnapshot(fixture) };
    let captures = 0;
    let notifications = 0;
    const controller = createController({
      historyLimit: HISTORY_LIMIT,
      captureSnapshot: () => { captures += 1; return cloneNormalizedSnapshot(owner.snapshot); },
      applySnapshot: snapshot => { owner.snapshot = cloneNormalizedSnapshot(snapshot); },
      onHistoryChanged: () => { notifications += 1; },
      setTimer: () => { throw new Error('Immediate commits must not create timers'); },
      clearTimer: () => { throw new Error('No timer should exist'); },
    });
    controller.activate(owner);
    const start = performance.now();
    for (let index = 0; index < OPERATIONS_PER_RUN; index += 1) {
      controller.beginMutation();
      owner.snapshot.values.brightness = changed ? index + 1 : 0;
      controller.commitMutation();
    }
    const elapsedMs = performance.now() - start;
    assert.equal(captures, OPERATIONS_PER_RUN * 2, 'Each begin/commit must capture both snapshots');
    assert.equal(notifications, OPERATIONS_PER_RUN);
    assert.equal(owner.snapshot.subjectMask.maskDataUrl, fixture.subjectMask.maskDataUrl);
    verifyHistory(controller, owner, changed, fixture.subjectMask.maskDataUrl);
    controller.deactivate();
    if (run >= WARMUP_RUNS) samples.push(elapsedMs / OPERATIONS_PER_RUN);
  }
  return { operation: changed ? 'brightness-change' : 'no-op',
    millisecondsPerBeginEditCommit: distribution(samples),
    verifiedRuns: WARMUP_RUNS + MEASURED_RUNS,
    historyVerification: changed ? '40-entry limit, 40 undos/redos, redo survives no-op' : 'no history created' };
}

function benchmarkOperation(operation, expected) {
  const samples = [];
  for (let run = 0; run < WARMUP_RUNS + MEASURED_RUNS; run += 1) {
    let successes = 0;
    const start = performance.now();
    for (let index = 0; index < OPERATIONS_PER_RUN; index += 1) {
      if (operation() === expected) successes += 1;
    }
    const elapsedMs = performance.now() - start;
    assert.equal(successes, OPERATIONS_PER_RUN);
    if (run >= WARMUP_RUNS) samples.push(elapsedMs / OPERATIONS_PER_RUN);
  }
  return distribution(samples);
}

function verifyCandidateSemantics(fixture) {
  const clone = () => cloneNormalizedSnapshot(fixture);
  const changed = clone(); changed.values.brightness = 1;
  const undefinedField = clone(); undefinedField.values.extra = undefined;
  const nullField = clone(); nullField.values.extra = null;
  const nanField = clone(); nanField.values.extra = NaN;
  const reordered = Object.fromEntries(Object.entries(clone()).reverse());
  const reorderedMask = clone();
  reorderedMask.subjectMask = Object.fromEntries(Object.entries(reorderedMask.subjectMask).reverse());
  const movedMaskKey = clone();
  delete movedMaskKey.subjectMask.maskDataUrl;
  movedMaskKey.subjectMask.maskDataUrl = fixture.subjectMask.maskDataUrl;
  const nullMask = clone(); nullMask.subjectMask.maskDataUrl = null;
  const missingMask = clone(); delete missingMask.subjectMask.maskDataUrl;
  const undefinedArray = clone(); undefinedArray.masks.push(undefined);
  const nullArray = clone(); nullArray.masks.push(null);
  const negativeZero = clone(); negativeZero.values.brightness = -0;
  const pairs = [
    ['same data', fixture, clone(), true],
    ['changed scalar', fixture, changed, false],
    ['undefined field omitted', fixture, undefinedField, true],
    ['NaN serializes as null', nullField, nanField, true],
    ['top-level key order', fixture, reordered, false],
    ['nested key order', fixture, reorderedMask, false],
    ['only mask key position moved', fixture, movedMaskKey, false],
    ['null mask is distinct', fixture, nullMask, false],
    ['missing and null are distinct', missingMask, nullMask, false],
    ['undefined array entry serializes as null', undefinedArray, nullArray, true],
    ['falsey roots normalize to null', null, undefined, true],
    ['negative zero serializes as zero', fixture, negativeZero, true],
  ];
  for (const [description, left, right, expected] of pairs) {
    assert.equal(jsonEqual(left, right), expected, `Existing JSON semantics: ${description}`);
    assert.equal(compareNormalizedSnapshotsWithoutEncodingMask(left, right), expected,
      `Candidate semantics: ${description}`);
  }
  return pairs.length;
}

const report = {
  ok: false, label, runId,
  methodology: {
    maskDataUrlBytes: MASK_SIZES, operationsPerRun: OPERATIONS_PER_RUN,
    warmupRuns: WARMUP_RUNS, measuredRuns: MEASURED_RUNS, historyLimit: HISTORY_LIMIT,
    realController: 'activate outside timing; beginMutation -> brightness edit/same-value assignment -> commitMutation timed',
    capture: 'Recursively clones normalized object/array containers, retaining primitive strings without JSON encoding or copying mask contents.',
    included: ['two captures', 'current controller equality and history updates', 'onHistoryChanged counter'],
    excluded: ['fixture generation', 'normalization math', 'DOM/rendering', 'timers', 'verification', 'getHistoryState JSON deep clone'],
    comparatorHypothesis: 'Compare maskDataUrl string values, retain its key with empty string, then use the same JSON equality for remaining fields; benchmark only, not injected into controller.',
    hypothesisScope: 'Normalized plain snapshots without getters/toJSON. String comparison may still scan distinct strings with equal contents.',
    caveat: 'Synthetic Node/V8 synchronous work, not full UI latency. Comparator-only timings are separate from actual controller timings.',
  },
  environment: { platform: process.platform, arch: process.arch, osRelease: os.release(),
    cpuModel: os.cpus()[0]?.model || '', logicalCpus: os.cpus().length, versions: process.versions,
    sessionSourceSha256: sha256(fs.readFileSync(path.join(__dirname, '..', 'src', 'photo-editor-session.js'))) },
  results: [],
};

try {
  const started = performance.now();
  for (const maskBytes of MASK_SIZES) {
    const fixture = createSnapshot(createMaskDataUrl(maskBytes));
    const equivalent = cloneNormalizedSnapshot(fixture);
    const changed = cloneNormalizedSnapshot(fixture); changed.values.brightness = 1;
    const comparatorCases = [['no-op-shared-mask', equivalent], ['brightness-change-shared-mask', changed]];
    if (maskBytes) {
      const equalNewString = cloneNormalizedSnapshot(fixture);
      equalNewString.subjectMask.maskDataUrl = Buffer.from(fixture.subjectMask.maskDataUrl, 'ascii').toString('ascii');
      const changedMask = cloneNormalizedSnapshot(fixture);
      changedMask.subjectMask.maskDataUrl = `${fixture.subjectMask.maskDataUrl.slice(0, -1)}!`;
      comparatorCases.push(['same-mask-content-reconstructed', equalNewString], ['mask-change-last-byte', changedMask]);
    }
    const result = {
      maskDataUrlBytes: maskBytes, snapshotJsonBytes: Buffer.byteLength(JSON.stringify(fixture)),
      inputSha256: sha256(JSON.stringify(fixture)), candidateSemanticCases: verifyCandidateSemantics(fixture),
      controller: [benchmarkController(fixture, false), benchmarkController(fixture, true)],
      captureOnly: benchmarkOperation(() => cloneNormalizedSnapshot(fixture).subjectMask.maskDataUrl, fixture.subjectMask.maskDataUrl),
      comparators: comparatorCases.map(([operation, right]) => {
        const expected = jsonEqual(fixture, right);
        const baseline = benchmarkOperation(() => jsonEqual(fixture, right), expected);
        const hypothesis = benchmarkOperation(() => compareNormalizedSnapshotsWithoutEncodingMask(fixture, right), expected);
        return { operation, expectedEqual: expected, currentJsonEquality: baseline,
          maskStringFirstHypothesis: hypothesis, medianRatio: baseline.medianMs / hypothesis.medianMs };
      }),
    };
    report.results.push(result);
    console.log(`${maskBytes / 1024} KiB mask: no-op ${result.controller[0].millisecondsPerBeginEditCommit.medianMs.toFixed(4)} ms, change ${result.controller[1].millisecondsPerBeginEditCommit.medianMs.toFixed(4)} ms per begin/edit/commit; history verified`);
  }
  report.ok = true;
  report.elapsedMs = performance.now() - started;
} catch (error) {
  report.error = error.stack || String(error);
  console.error(report.error);
  process.exitCode = 1;
} finally {
  fs.mkdirSync(outputDirectory, { recursive: true });
  fs.writeFileSync(reportPath, JSON.stringify(report, null, 2) + '\n');
  console.log(`Photo editor history benchmark report: ${reportPath}`);
}
