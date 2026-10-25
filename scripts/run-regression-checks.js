'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { runSteps } = require('./lib/regression-runner');
const { captureSourceHashes } = require('./lib/source-inventory');

const ROOT = path.resolve(__dirname, '..');
const SUITES = {
  unit: [{ name: 'check:syntax' }, { name: 'check:security' }, { name: 'test' }, { name: 'audit:i18n', args: ['--strict'] }],
  library: [{ name: 'smoke:photo-list' }, { name: 'smoke:thumbnail-store' }],
  retouch: [
    { name: 'smoke:subject-session' }, { name: 'smoke:save-session' },
    { name: 'smoke:overlay-session' }, { name: 'smoke:worker' },
    { name: 'smoke:resource-readiness' },
    { name: 'smoke:export', timeoutMs: 360000 },
    { name: 'smoke:coordinates', args: ['--skip-screenshots'] },
    { name: 'smoke:mask-brush', args: ['--skip-screenshots'] },
    { name: 'smoke:presets' },
  ],
  data: [
    { name: 'smoke:data' }, { name: 'smoke:backup-validation' },
    { name: 'smoke:backup-restore' }, { name: 'smoke:world-identity' },
  ],
  application: [{ name: 'smoke:app-update' }],
};

async function main() {
  // Only this runner's progress display may lose its pipe; diagnostics from every
  // check remain in their own files. Do not suppress any other stream error.
  for (const stream of [process.stdout, process.stderr]) {
    stream.on('error', error => { if (error.code !== 'EPIPE') throw error; });
  }
  const args = process.argv.slice(2);
  if (args.length > 1 || (args[0] && !/^--suite=(all|unit|library|retouch|data|application)$/.test(args[0]))) {
    throw new Error('Usage: npm run check -- [--suite=all|unit|library|retouch|data|application]');
  }
  const suite = args[0]?.slice('--suite='.length) || 'all';
  const npmCli = process.env.npm_execpath;
  if (!npmCli || !fs.existsSync(npmCli)) {
    throw new Error('Start this runner with npm run check so it uses the active npm CLI.');
  }
  const plan = suite === 'all' ? Object.values(SUITES).flat() : SUITES[suite];
  const scripts = require('../package.json').scripts;
  for (const step of plan) {
    if (!scripts[step.name]) throw new Error(`Missing package script: ${step.name}`);
  }
  const output = path.join(ROOT, 'out', 'regression-checks',
    `${new Date().toISOString().replace(/[:.]/g, '-')}-${process.pid}`);
  fs.mkdirSync(output, { recursive: true });
  const summaryPath = path.join(output, 'result.json');
  const report = { ok: false, status: 'running', suite, startedAt: new Date().toISOString(),
    sourceHashes: await captureSourceHashes(ROOT), sourceUnchanged: false,
    plannedSteps: plan.map(step => step.name), results: [] };
  const persist = () => {
    const temporaryPath = `${summaryPath}.tmp`;
    fs.writeFileSync(temporaryPath, JSON.stringify(report, null, 2));
    fs.renameSync(temporaryPath, summaryPath);
  };
  persist();
  console.log(`Regression checks (${suite}). Logs: ${output}`);
  const controller = new AbortController();
  const stop = () => controller.abort();
  process.once('SIGINT', stop);
  process.once('SIGTERM', stop);
  try {
    const steps = plan.map((step, index) => ({ ...step, command: process.execPath,
      args: [npmCli, 'run', step.name, ...(step.args ? ['--', ...step.args] : [])],
      cwd: ROOT, timeoutMs: step.timeoutMs || 180000,
      logPath: path.join(output, `${String(index + 1).padStart(2, '0')}-${step.name.replace(/:/g, '-')}.log`) }));
    await runSteps(steps, { signal: controller.signal, onResult(result) {
      report.results.push(result);
      persist();
      console.log(`${result.ok ? 'PASS' : 'FAIL'} ${result.name} (${(result.durationMs / 1000).toFixed(1)}s)`);
      if (!result.ok) console.error(`Check log: ${result.logPath}`);
    } });
    report.sourceUnchanged = JSON.stringify(report.sourceHashes) === JSON.stringify(await captureSourceHashes(ROOT));
    report.ok = report.sourceUnchanged && report.results.length === plan.length && report.results.every(result => result.ok);
    if (!report.sourceUnchanged) report.error = 'Build inputs changed during checks; rerun against the final source.';
    report.status = report.ok ? 'passed' : controller.signal.aborted ? 'aborted' : 'failed';
  } catch (error) {
    report.status = 'failed';
    report.error = error.stack || String(error);
  } finally {
    process.removeListener('SIGINT', stop);
    process.removeListener('SIGTERM', stop);
    report.finishedAt = new Date().toISOString();
    persist();
  }
  console.log(`Result: ${summaryPath}`);
  process.exitCode = report.ok ? 0 : controller.signal.aborted ? 130 : 1;
}

module.exports = { SUITES };
if (require.main === module) {
  main().catch(error => { console.error(error.stack || error); process.exitCode = 1; });
}
