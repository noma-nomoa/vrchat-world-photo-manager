'use strict';

const assert = require('node:assert/strict');
const { test } = require('node:test');
const { mapWithConcurrencyLimit } = require('../src/concurrency');
const deferred = () => {
  let resolve;
  let reject;
  const promise = new Promise((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
};
const tick = () => new Promise(resolve => setImmediate(resolve));

test('empty and non-array inputs do not invoke mapper or progress', async () => {
  for (const input of [[], null, undefined, 'photo']) {
    assert.deepEqual(await mapWithConcurrencyLimit(input, 3,
      () => assert.fail('Unexpected mapping'), () => assert.fail('Unexpected progress')), []);
  }
});

test('concurrent results retain input order while progress follows actual completions', async () => {
  const gates = [deferred(), deferred(), deferred()];
  const progress = [];
  const inputs = Object.freeze(['a', 'b', 'c']);
  const pending = mapWithConcurrencyLimit(inputs, 3, (_item, index) => gates[index].promise, event => progress.push(event));
  gates[2].resolve('C');
  await tick();
  gates[0].resolve('A');
  await tick();
  gates[1].resolve('B');
  assert.deepEqual(await pending, ['A', 'B', 'C']);
  assert.deepEqual(progress, [
    { completedCount: 1, totalCount: 3, item: 'c', index: 2, result: 'C' },
    { completedCount: 2, totalCount: 3, item: 'a', index: 0, result: 'A' },
    { completedCount: 3, totalCount: 3, item: 'b', index: 1, result: 'B' },
  ]);
});

test('default, zero and negative concurrency preserve sequential processing', async () => {
  for (const limit of [undefined, 0, -2]) {
    const first = deferred();
    const started = [];
    const pending = mapWithConcurrencyLimit(['a', 'b'], limit, (item, index) => {
      started.push(item);
      return index ? 'B' : first.promise;
    });
    assert.deepEqual(started, ['a']);
    first.resolve('A');
    assert.deepEqual(await pending, ['A', 'B']);
    assert.deepEqual(started, ['a', 'b']);
  }
});

test('a large concurrency limit still invokes each item exactly once', async () => {
  const visited = [];
  assert.deepEqual(await mapWithConcurrencyLimit([2, 3], 100, (value, index) => {
    visited.push(index);
    return value * 2;
  }), [4, 6]);
  assert.deepEqual(visited, [0, 1]);
});

test('mapper rejection propagates the same error and does not report it as completed', async () => {
  const error = new Error('mapping failed');
  const progress = [];
  await assert.rejects(mapWithConcurrencyLimit(['a', 'b'], 1, item => {
    if (item === 'a') throw error;
    assert.fail('Sequential work must stop after failure');
  }, event => progress.push(event)), value => value === error);
  assert.deepEqual(progress, []);
});

test('progress observer rejection is not swallowed', async () => {
  const error = new Error('observer failed');
  await assert.rejects(mapWithConcurrencyLimit(['a'], 1, item => item,
    () => { throw error; }), value => value === error);
});

test('independent batches do not share pending indices or progress counters', async () => {
  const gate = deferred();
  const pending = mapWithConcurrencyLimit(['a'], 1, () => gate.promise);
  assert.deepEqual(await mapWithConcurrencyLimit(['b'], 1, () => 'B'), ['B']);
  gate.resolve('A');
  assert.deepEqual(await pending, ['A']);
});
