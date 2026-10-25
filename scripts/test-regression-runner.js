'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { test } = require('node:test');
const { runStep, runSteps } = require('./lib/regression-runner');

function fixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-runner-test-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  return { root, step: (name, code, options = {}) => ({ name, command: process.execPath,
    args: ['-e', code], cwd: root, logPath: path.join(root, `${name}.log`), timeoutMs: 10000, ...options }) };
}

test('successful child output is written to a regular file, including stderr', async t => {
  const { step } = fixture(t);
  const spec = step('success', "const fs=require('fs'); if(!fs.fstatSync(1).isFile()||!fs.fstatSync(2).isFile())process.exit(3); console.log('stdout'); console.error('stderr');");
  const result = await runStep(spec);
  assert.equal(result.ok, true);
  assert.equal(result.exitCode, 0);
  assert.match(fs.readFileSync(result.logPath, 'utf8'), /stdout[\s\S]*stderr/);
});

test('nonzero exit preserves diagnostics and stops the remaining checks', async t => {
  const { root, step } = fixture(t);
  const observed = [];
  const results = await runSteps([
    step('first', "console.log('first passed')"),
    step('failure', "console.error('deliberate failure'); process.exit(7)"),
    step('never', "require('fs').writeFileSync('should-not-exist', 'bad')"),
  ], { onResult: result => observed.push(result.name) });
  assert.deepEqual(observed, ['first', 'failure']);
  assert.equal(results[1].ok, false);
  assert.equal(results[1].exitCode, 7);
  assert.match(fs.readFileSync(results[1].logPath, 'utf8'), /deliberate failure/);
  assert.equal(fs.existsSync(path.join(root, 'should-not-exist')), false);
});

test('spawn errors are failures rather than successful missing checks', async t => {
  const { root, step } = fixture(t);
  const result = await runStep(step('missing', '', { command: path.join(root, 'missing-executable') }));
  assert.equal(result.ok, false);
  assert.match(result.error, /ENOENT/);
});

test('timeout stops a hung check and reports failure', async t => {
  const { step } = fixture(t);
  const result = await runStep(step('timeout', 'setInterval(() => {}, 1000)', { timeoutMs: 300 }));
  assert.equal(result.ok, false);
  assert.equal(result.timedOut, true);
  assert.equal(result.terminationError, undefined);
});

test('timeout also terminates the child process spawned by a check', async t => {
  const { root, step } = fixture(t);
  const spec = step('tree', `const {spawn}=require('child_process'); const fs=require('fs');
    const child=spawn(process.execPath,['-e','setInterval(() => {}, 1000)'],{stdio:'ignore',windowsHide:true});
    fs.writeFileSync('descendant.pid',String(child.pid)); setInterval(()=>{},1000);`, { timeoutMs: 1200 });
  const result = await runStep(spec);
  assert.equal(result.timedOut, true);
  assert.equal(result.terminationError, undefined);
  const pid = Number(fs.readFileSync(path.join(root, 'descendant.pid'), 'utf8'));
  const alive = () => { try { process.kill(pid, 0); return true; } catch (error) {
    if (error.code === 'ESRCH') return false; throw error;
  } };
  for (let i = 0; i < 50 && alive(); i += 1) await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(alive(), false, 'A timed-out check must not leave its descendant running');
});

test('cancellation before launch never starts the child', async t => {
  const { root, step } = fixture(t);
  const controller = new AbortController();
  controller.abort();
  const result = await runStep(step('aborted', "require('fs').writeFileSync('should-not-exist', 'bad')", { signal: controller.signal }));
  assert.equal(result.aborted, true);
  assert.equal(result.ok, false);
  assert.equal(fs.existsSync(path.join(root, 'should-not-exist')), false);
});

test('cancellation during execution terminates the check', async t => {
  const { step } = fixture(t);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 300);
  t.after(() => clearTimeout(timer));
  const result = await runStep(step('abort-running', 'setInterval(() => {}, 1000)', { signal: controller.signal }));
  assert.equal(result.aborted, true);
  assert.equal(result.timedOut, false);
  assert.equal(result.ok, false);
  assert.equal(result.terminationError, undefined);
});

test('an unbounded step is rejected before launching', t => {
  const { step } = fixture(t);
  for (const timeoutMs of [undefined, 0, -1, NaN, Infinity]) {
    assert.throws(() => runStep(step('invalid', '', { timeoutMs })), /positive step timeout/);
  }
});
