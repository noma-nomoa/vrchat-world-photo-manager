'use strict';

// Stable inputs for the pre-extraction renderer/worker characterization.
// These fixtures contain no production color calculation or expected-value generator.
const WIDTH = 17;
const HEIGHT = 13;
const ADJUSTMENT_KEYS = Object.freeze([
  'brightness', 'exposure', 'contrast', 'highlights', 'whites', 'shadows',
  'blacks', 'gamma', 'temperature', 'tint', 'saturation', 'vibrance', 'fade', 'grain',
]);

function createValueCases() {
  const cases = [{ id: 'default', values: {} }];
  for (const key of ADJUSTMENT_KEYS) {
    for (const [label, value] of [
      ['min', -100], ['negative', -1], ['positive', 1], ['max', 100],
      ['nan', NaN], ['infinity', Infinity], ['string', '27'],
    ]) {
      cases.push({ id: `${key}-${label}`, values: { [key]: value } });
    }
  }
  for (let index = 0; index < 16; index += 1) {
    cases.push({
      id: `mixed-${index}`,
      values: Object.fromEntries(ADJUSTMENT_KEYS.map((key, keyIndex) => [
        key, ((index * 37 + keyIndex * 61) % 201) - 100,
      ])),
    });
  }
  return cases;
}

function createCurveCases() {
  const cases = [
    { id: 'null', curve: null },
    { id: 'default', curve: {} },
    { id: 'invalid', curve: {
      mode: 'bad', channel: 'bad', points: { rgb: { r: [NaN, Infinity, '0.6'] } },
    } },
  ];
  for (const mode of ['rgb', 'hsv']) {
    const channels = ['master', ...(mode === 'rgb' ? ['r', 'g', 'b'] : ['h', 's', 'v'])];
    for (const channel of channels) {
      for (const [label, points] of [
        ['five-point', [0, 0.14, 0.57, 0.9, 1]],
        ['legacy-three-point', [1, 0.5, 0]],
      ]) {
        cases.push({
          id: `${mode}-${channel}-${label}`,
          curve: { mode, channel, points: { [mode]: { [channel]: points } } },
        });
      }
    }
  }
  return cases;
}

function createInputRgba() {
  return Uint8ClampedArray.from({ length: WIDTH * HEIGHT * 4 }, (_, index) => (
    index % 4 === 3
      ? [0, 1, 127, 254, 255][(index >> 2) % 5]
      : ((index * 47 + (index >> 2) * 13) % 256)
  ));
}

module.exports = { WIDTH, HEIGHT, createValueCases, createCurveCases, createInputRgba };
