'use strict';

// Run with Node.js. No Electron, browser, image files, or user profile is needed.
// Expected values were captured from both implementations before extraction;
// this test never derives its oracle from renderer.js or photo-editor-worker.js.
const assert = require('node:assert/strict');
const { createHash } = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const cases = require('./fixtures/photo-editor-color-core-cases');
const baseline = require('./fixtures/photo-editor-color-core-baseline.json');

function hash(value) {
  return createHash('sha256').update(value).digest('hex');
}

function stableJson(value) {
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (ArrayBuffer.isView(value)) return stableJson(Array.from(value));
  if (Array.isArray(value)) return `[${value.map(stableJson).join(',')}]`;
  return `{${Object.keys(value).sort().map(key => (
    `${JSON.stringify(key)}:${stableJson(value[key])}`
  )).join(',')}}`;
}

function comparableLookups(lookups) {
  const { curve: _normalizedCurve, ...calculation } = lookups;
  return calculation;
}

function comparableParams(params) {
  return { ...params, curveLookups: comparableLookups(params.curveLookups) };
}

function verifyCore(core, { quiet = false } = {}) {
  const curveCases = cases.createCurveCases();
  const valueCases = cases.createValueCases();
  const input = cases.createInputRgba();
  assert.equal(baseline.schemaVersion, 1);
  assert.equal(baseline.width, cases.WIDTH);
  assert.equal(baseline.height, cases.HEIGHT);
  assert.equal(baseline.inputSha256, hash(input));
  assert.equal(baseline.curves.length, curveCases.length);
  assert.equal(baseline.cases.length, curveCases.length * valueCases.length);
  assert.equal(baseline.preExtractionVerification.rendererWorkerParity, true);

  const combinedPixels = createHash('sha256');
  let count = 0;
  for (const [curveIndex, { id: curveId, curve }] of curveCases.entries()) {
    const expectedCurve = baseline.curves[curveIndex];
    assert.equal(expectedCurve.id, curveId);
    const curveBefore = structuredClone(curve);
    const normalized = core.normalizePhotoEditorCurveState(curve);
    assert.equal(hash(stableJson(normalized)), expectedCurve.normalizedSha256,
      `${curveId}: normalized curve changed`);
    assert.equal(stableJson(normalized), stableJson(expectedCurve.normalized));
    const lookups = core.createPhotoEditorCurveLookups(curve);
    assert.equal(hash(stableJson(comparableLookups(lookups))), expectedCurve.lookupsSha256,
      `${curveId}: curve lookup tables changed`);
    for (const mode of ['rgb', 'hsv']) {
      for (const lookup of Object.values(lookups[mode])) {
        assert.equal(Object.prototype.toString.call(lookup), '[object Float32Array]');
        assert.equal(lookup.length, 256);
      }
    }

    for (const { id: valueId, values } of valueCases) {
      const id = `${curveId}/${valueId}`;
      const expected = baseline.cases[count];
      assert.equal(expected.id, id);
      const valuesBefore = structuredClone(values);
      const params = core.getPhotoEditorAdjustmentRenderParams(values, curve);
      assert.equal(hash(stableJson(comparableParams(params))), expected.paramsSha256,
        `${id}: adjustment parameters/LUT changed`);
      const imageData = { data: new Uint8ClampedArray(input) };
      const dataReference = imageData.data;
      const paramsBefore = stableJson(comparableParams(params));
      core.applyPhotoEditorPixelAdjustmentsToImageData(imageData, cases.WIDTH, params);
      assert.equal(imageData.data, dataReference, `${id}: must modify the existing pixel buffer`);
      assert.equal(hash(imageData.data), expected.pixelsSha256, `${id}: RGBA pixels changed`);
      for (let index = 3; index < input.length; index += 4) {
        assert.equal(imageData.data[index], input[index], `${id}: alpha changed at pixel ${index >> 2}`);
      }
      assert.deepEqual(values, valuesBefore, `${id}: input adjustment values were modified`);
      assert.deepEqual(curve, curveBefore, `${id}: input curve was modified`);
      assert.equal(stableJson(comparableParams(params)), paramsBefore,
        `${id}: render parameters/LUT were modified`);
      combinedPixels.update(imageData.data);
      count += 1;
    }
  }
  const aggregateSha256 = combinedPixels.digest('hex');
  assert.equal(aggregateSha256, baseline.aggregatePixelsSha256);
  if (!quiet) console.log(JSON.stringify({
    cases: count, curves: curveCases.length, valueCases: valueCases.length,
    normalizationAndLookupBaselines: 'passed', parameterAndPixelBaselines: 'passed',
    inputAndAlphaPreservation: 'passed', aggregateSha256,
  }, null, 2));
  return { count, aggregateSha256 };
}

function verifyScriptEntryPoints(core) {
  const moduleSource = fs.readFileSync(path.join(__dirname, '..', 'src', 'photo-editor-color-core.js'), 'utf8');
  const normalizedHash = baseline.curves[0].normalizedSha256;
  for (const environment of ['window', 'self']) {
    const context = vm.createContext({});
    vm.runInContext(`this.${environment} = this;`, context);
    vm.runInContext(moduleSource, context, { filename: 'photo-editor-color-core.js' });
    const scriptApi = context.PhotoEditorColorCore;
    assert(scriptApi, `${environment}: global script API was not created`);
    assert.deepEqual(Object.keys(scriptApi).sort(), Object.keys(core).sort());
    assert.equal(hash(stableJson(scriptApi.normalizePhotoEditorCurveState(null))), normalizedHash);
    const data = { data: cases.createInputRgba() };
    const params = scriptApi.getPhotoEditorAdjustmentRenderParams({}, null);
    scriptApi.applyPhotoEditorPixelAdjustmentsToImageData(data, cases.WIDTH, params);
    assert.equal(hash(data.data), baseline.cases[0].pixelsSha256);
  }
  console.log('CommonJS, browser script, and classic Worker script entry points: passed');
}

function verifyCurveContracts(core) {
  const defaults = core.getDefaultPhotoEditorCurveState();
  assert.deepEqual(defaults, baseline.curves[1].normalized);
  const secondDefaults = core.getDefaultPhotoEditorCurveState();
  defaults.points.rgb.r[1] = 0.9;
  assert.deepEqual(secondDefaults, baseline.curves[1].normalized,
    'Default curve calls must own their point arrays');
  assert.deepEqual(core.clonePhotoEditorCurvePoints([1, 0.5, 0]), [1, 0.75, 0.5, 0.25, 0]);
  assert.equal(core.isPhotoEditorCurvePointsIdentity([0.000999, 0.25, 0.5, 0.75, 1]), true);
  assert.equal(core.isPhotoEditorCurvePointsIdentity([0.001, 0.25, 0.5, 0.75, 1]), false);
  assert.equal(core.isPhotoEditorCurveStateIdentity(secondDefaults), true);
  for (const mode of ['rgb', 'hsv']) {
    for (const channel of Object.keys(secondDefaults.points[mode])) {
      const changed = structuredClone(secondDefaults);
      changed.points[mode][channel][2] = 0.6;
      assert.equal(core.isPhotoEditorCurveStateIdentity(changed), false, `${mode}/${channel}`);
    }
  }
  const lookups = core.createPhotoEditorCurveLookups(secondDefaults);
  assert.deepEqual(lookups.curve, secondDefaults,
    'Keep the renderer lookup result normalized curve field');
  assert.notEqual(lookups.curve, secondDefaults);
  assert.equal(core.getPhotoEditorCurveLookupIndex(0.5, 1), 128);
  assert.equal(core.getPhotoEditorCurveLookupIndex(127.5), 128);
  assert.equal(core.getPhotoEditorCurveLookupIndex(-100), 0);
  assert.equal(core.getPhotoEditorCurveLookupIndex(300), 255);
  console.log('Curve cloning, legacy points, identity threshold, and lookup rounding: passed');
}

if (require.main === module) {
  const core = require('../src/photo-editor-color-core');
  verifyCore(core);
  verifyCurveContracts(core);
  verifyScriptEntryPoints(core);
}

module.exports = { verifyCore, hash, stableJson, comparableLookups, comparableParams };
