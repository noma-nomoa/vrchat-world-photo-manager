'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { spawn } = require('node:child_process');

// Keep child output attached to a file, even if the terminal that started the
// checks closes. In particular, Electron must not inherit a disappearing pipe.
function runStep({ name, command, args = [], cwd, logPath, timeoutMs, signal }) {
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
    throw new Error('A positive step timeout is required');
  }
  fs.mkdirSync(path.dirname(logPath), { recursive: true });
  const started = Date.now();
  const result = { name, command, args, logPath, startedAt: new Date(started).toISOString(),
    ok: false, exitCode: null, signal: null, timedOut: false, aborted: false };
  if (signal?.aborted) {
    fs.writeFileSync(logPath, 'Canceled before starting.\n');
    return Promise.resolve({ ...result, aborted: true, durationMs: 0 });
  }

  return new Promise(resolve => {
    const log = fs.openSync(logPath, 'w');
    let child;
    try {
      child = spawn(command, args, { cwd, stdio: ['ignore', log, log], windowsHide: true,
        // A private process group lets POSIX hosts stop descendants too.
        detached: process.platform !== 'win32' });
    } catch (error) {
      fs.closeSync(log);
      resolve({ ...result, error: error.message, durationMs: Date.now() - started });
      return;
    }
    fs.closeSync(log);
    result.pid = child.pid ?? null;
    let stopping = false;
    let timeout;

    const stop = reason => {
      if (stopping || child.exitCode !== null || child.signalCode !== null) return;
      stopping = true;
      result[reason] = true;
      if (!child.pid) return;
      if (process.platform === 'win32') {
        // Kill only this check's process tree, never Electron by image name.
        const killer = spawn('taskkill.exe', ['/pid', String(child.pid), '/t', '/f'],
          { stdio: 'ignore', windowsHide: true });
        killer.once('error', error => {
          result.terminationError = error.message;
          child.kill();
        });
        killer.once('exit', code => {
          if (code && child.exitCode === null && child.signalCode === null) {
            result.terminationError = `taskkill exited with ${code}`;
            child.kill();
          }
        });
      } else {
        try { process.kill(-child.pid, 'SIGKILL'); }
        catch (error) {
          if (error.code !== 'ESRCH') result.terminationError = error.message;
          child.kill('SIGKILL');
        }
      }
    };
    const abort = () => stop('aborted');
    signal?.addEventListener('abort', abort, { once: true });
    timeout = setTimeout(() => stop('timedOut'), timeoutMs);
    // Cover an abort that arrived between the initial check and registration.
    if (signal?.aborted) abort();
    child.once('error', error => { result.error = error.message; });
    child.once('close', (code, childSignal) => {
      clearTimeout(timeout);
      signal?.removeEventListener('abort', abort);
      result.exitCode = code;
      result.signal = childSignal;
      result.durationMs = Date.now() - started;
      result.ok = code === 0 && !result.error && !result.timedOut && !result.aborted
        && !result.terminationError;
      resolve(result);
    });
  });
}

async function runSteps(steps, { onResult = () => {}, signal } = {}) {
  const results = [];
  for (const step of steps) {
    const result = await runStep({ ...step, signal });
    results.push(result);
    await onResult(result, results);
    if (!result.ok) break;
  }
  return results;
}

module.exports = { runStep, runSteps };
