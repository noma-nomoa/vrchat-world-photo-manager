'use strict';

// Build a local review candidate with the normal Forge configuration. This
// script never publishes, installs, changes version numbers, or reuses output.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { spawnSync } = require('node:child_process');
const asar = require('@electron/asar');
const yauzl = require('yauzl');
const { runStep } = require('./lib/regression-runner');
const ROOT = path.resolve(__dirname, '..');
const OUTPUT = path.join(ROOT, 'out');
const NATIVE = path.join(ROOT, 'node_modules', 'better-sqlite3', 'prebuilds', 'win32-x64.node');
const hashBytes = bytes => createHash('sha256').update(bytes).digest('hex');

function listFiles(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name)).flatMap(entry => {
    const file = path.join(directory, entry.name);
    if (entry.isSymbolicLink()) throw new Error(`Unexpected symbolic link: ${file}`);
    return entry.isDirectory() ? listFiles(file) : [file];
  });
}
function assertInside(directory, file) {
  const relative = path.relative(directory, file);
  assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative), `Path escaped candidate: ${file}`);
}
function inputSnapshot() {
  return [...listFiles(path.join(ROOT, 'src')).filter(file => /\.(?:js|html|css)$/.test(file)),
    ...['package.json', 'package-lock.json', 'forge.config.js'].map(file => path.join(ROOT, file))].map(file => ({
    file: path.relative(ROOT, file).replaceAll('\\', '/'), sha256: hashBytes(fs.readFileSync(file)), mtimeMs: fs.statSync(file).mtimeMs,
  }));
}
async function hashFile(file, algorithm = 'sha256') {
  const hash = createHash(algorithm);
  for await (const chunk of fs.createReadStream(file)) hash.update(chunk);
  return hash.digest('hex');
}
function parseReleases(text) {
  const lines = text.trim().split(/\r?\n/).filter(Boolean);
  assert(lines.length > 0, 'RELEASES is empty');
  return lines.map(line => {
    const match = /^([0-9a-f]{40})\s+(\S+)\s+(\d+)$/i.exec(line.trim());
    assert(match, `Malformed RELEASES line: ${line}`);
    const [, sha1, name, size] = match;
    assert(name === path.basename(name) && !/[\\/]/.test(name) && name.endsWith('.nupkg'), 'Unsafe RELEASES filename');
    assert(Number.isSafeInteger(Number(size)) && Number(size) > 0, 'Invalid RELEASES size');
    return { name, sha1: sha1.toLowerCase(), bytes: Number(size) };
  });
}
function validateRegressionReport(checks, sourceHashes, expectedPlan) {
  assert(checks?.ok && checks.status === 'passed' && checks.suite === 'all', 'All regression suites must have passed');
  assert.equal(checks.sourceUnchanged, true, 'Regression report did not verify unchanged source');
  assert(Array.isArray(expectedPlan) && expectedPlan.length > 0, 'Trusted regression plan is missing');
  assert.deepEqual(checks.plannedSteps, expectedPlan, 'Regression report does not contain the complete current plan');
  assert(Array.isArray(checks.results), 'Regression results are missing');
  assert.deepEqual(checks.results.map(result => result.name), expectedPlan, 'Regression results are incomplete or reordered');
  assert(checks.results.every(result => result.ok && result.exitCode === 0 && !result.signal
    && !result.timedOut && !result.aborted && !result.error && !result.terminationError), 'Regression result contains a failed check');
  assert.deepEqual(checks.sourceHashes, sourceHashes, 'Regression report does not match current sources/assets; rerun npm run check');
}
async function verifyExecutionStub(bytes, vendorBytes, applicationBytes) {
  const pe = await import('pe-library');
  const actual = pe.NtExecutable.from(bytes);
  const vendor = pe.NtExecutable.from(vendorBytes);
  const application = pe.NtExecutable.from(applicationBytes);
  assert.equal(actual.is32bit(), vendor.is32bit(), 'Execution stub architecture changed');
  const headerFields = ['magic', 'addressOfEntryPoint', 'baseOfCode', 'imageBase', 'sectionAlignment',
    'fileAlignment', 'subsystem', 'dllCharacteristics', 'sizeOfCode', 'sizeOfStackReserve', 'sizeOfHeapReserve'];
  for (const field of headerFields) assert.equal(actual.newHeader.optionalHeader[field], vendor.newHeader.optionalHeader[field], `Execution stub ${field} changed`);
  const sections = exe => exe.getAllSections().filter(section => section.info.name !== '.rsrc').map(section => ({
    name: section.info.name, bytes: section.data?.byteLength || 0, sha256: hashBytes(Buffer.from(section.data || [])),
    // The relocation section moves when resource size changes; its bytes stay identical.
    virtualAddress: section.info.name === '.reloc' ? null : section.info.virtualAddress,
    characteristics: section.info.characteristics,
  }));
  const resourceEntries = exe => pe.NtExecutableResource.from(exe).entries.map(entry => ({
    type: entry.type, id: entry.id, lang: entry.lang, bytes: entry.bin.byteLength, sha256: hashBytes(Buffer.from(entry.bin)),
  })).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
  assert.deepEqual(sections(actual), sections(vendor), 'Execution stub code differs from Squirrel vendor template');
  assert.deepEqual(resourceEntries(actual), resourceEntries(application), 'Execution stub resources differ from packaged application');
  assert.equal(actual.getExtraData()?.byteLength || 0, 0, 'Execution stub has unexpected trailing data');
  return { bytes: bytes.length, sha256: hashBytes(bytes), vendorSha256: hashBytes(vendorBytes),
    applicationSha256: hashBytes(applicationBytes), comparedCodeSections: sections(actual),
    comparedResourceEntries: resourceEntries(actual).length,
    transformation: 'Squirrel --releasify: vendor StubExecutable.exe code with resources copied from the packaged application EXE' };
}
async function prepareInstallerHelpers(candidateRoot, appDirectory) {
  const vendor = path.join(ROOT, 'node_modules', 'electron-winstaller', 'vendor');
  const integrity = fs.mkdtempSync(path.join(candidateRoot, 'installer-integrity-'));
  const squirrel = path.join(integrity, 'squirrel.exe');
  fs.copyFileSync(path.join(vendor, 'Squirrel.exe'), squirrel);
  // electron-winstaller/lib/index.js copies Squirrel.exe and applies setupIcon.
  const transform = await runStep({ name: 'verify-squirrel-icon', command: path.join(vendor, 'rcedit.exe'),
    args: [squirrel, '--set-icon', path.join(ROOT, 'img', 'logo.ico')], cwd: ROOT,
    logPath: path.join(integrity, 'rcedit.log'), timeoutMs: 30000 });
  assert(transform.ok, 'Could not reconstruct the standard Squirrel icon transform');
  const template = fs.readFileSync(path.join(vendor, 'StubExecutable.exe'));
  const application = fs.readFileSync(path.join(appDirectory, 'WorldShotLog.exe'));
  return {
    'squirrel.exe': { bytes: fs.statSync(squirrel).size, sha256: await hashFile(squirrel),
      vendorSha256: await hashFile(path.join(vendor, 'Squirrel.exe')), iconSha256: await hashFile(path.join(ROOT, 'img', 'logo.ico')),
      transformation: 'electron-winstaller: vendor/Squirrel.exe + rcedit --set-icon img/logo.ico', transform },
    'WorldShotLog_ExecutionStub.exe': { maxBytes: 5 * 1024 * 1024,
      verifyBuffer: bytes => verifyExecutionStub(bytes, template, application) },
  };
}
async function verifyNupkgPayload(nupkg, appDirectory, helpers = {}) {
  const expected = new Map();
  for (const file of listFiles(appDirectory)) {
    const name = path.relative(appDirectory, file).replaceAll('\\', '/');
    expected.set(name, { bytes: fs.statSync(file).size, sha256: await hashFile(file) });
  }
  const applicationFiles = expected.size;
  for (const [name, helper] of Object.entries(helpers)) {
    assert(!expected.has(name), 'Installer helper collides with application payload');
    expected.set(name, { ...helper });
  }
  const checked = [];
  await new Promise((resolve, reject) => {
    yauzl.open(nupkg, { lazyEntries: true, validateEntrySizes: true }, (error, zip) => {
      if (error) return reject(error);
      const fail = error => { zip.close(); reject(error); };
      zip.on('error', fail);
      zip.on('end', resolve);
      zip.on('entry', entry => {
        if (!entry.fileName.startsWith('lib/net45/') || entry.fileName.endsWith('/')) { zip.readEntry(); return; }
        const name = entry.fileName.slice('lib/net45/'.length);
        try {
          assert(expected.has(name), `Unexpected installer payload: ${name}`);
          assert(!checked.includes(name), `Duplicate installer payload: ${name}`);
          const specification = expected.get(name);
          if (specification.verifyBuffer) assert(entry.uncompressedSize <= specification.maxBytes, `Installer helper too large: ${name}`);
          else assert.equal(entry.uncompressedSize, specification.bytes, `Installer payload size differs: ${name}`);
        } catch (error) { fail(error); return; }
        zip.openReadStream(entry, (error, stream) => {
          if (error) return fail(error);
          const hash = createHash('sha256');
          const specification = expected.get(name);
          const chunks = [];
          stream.on('error', fail);
          stream.on('data', bytes => { hash.update(bytes); if (specification.verifyBuffer) chunks.push(bytes); });
          stream.on('end', async () => {
            try {
              if (specification.verifyBuffer) expected.set(name, await specification.verifyBuffer(Buffer.concat(chunks)));
              else assert.equal(hash.digest('hex'), specification.sha256, `Installer payload differs: ${name}`);
            }
            catch (error) { fail(error); return; }
            checked.push(name); zip.readEntry();
          });
        });
      });
      zip.readEntry();
    });
  });
  assert.equal(checked.length, expected.size, 'Installer omitted packaged application files');
  return { comparedFiles: checked.length, applicationFiles, installerHelperFiles: Object.keys(helpers), files: Object.fromEntries(expected) };
}
async function verifyArtifacts(candidateRoot, version) {
  const appDirectory = path.join(candidateRoot, 'build', 'WorldShot Log-win32-x64');
  const makeDirectory = path.join(candidateRoot, 'build', 'make', 'squirrel.windows', 'x64');
  const releasePath = path.join(makeDirectory, 'RELEASES');
  const releases = parseReleases(fs.readFileSync(releasePath, 'utf8'));
  const expectedPackage = `worldshot_log-${require('electron-winstaller').convertVersion(version)}-full.nupkg`;
  assert.equal(releases.length, 1, 'Fresh full build must contain exactly one RELEASES entry');
  assert.equal(releases[0].name, expectedPackage);
  const nupkg = path.join(makeDirectory, expectedPackage);
  assertInside(makeDirectory, nupkg);
  assert.equal(fs.statSync(nupkg).size, releases[0].bytes, 'RELEASES package size mismatch');
  assert.equal(await hashFile(nupkg, 'sha1'), releases[0].sha1, 'RELEASES SHA1 mismatch');
  const setup = path.join(makeDirectory, 'WorldShotLogSetup.exe');
  const handle = fs.openSync(setup, 'r');
  try { const magic = Buffer.alloc(2); fs.readSync(handle, magic, 0, 2, 0); assert.equal(magic.toString(), 'MZ', 'Installer is not a Windows executable'); }
  finally { fs.closeSync(handle); }
  const payload = await verifyNupkgPayload(nupkg, appDirectory, await prepareInstallerHelpers(candidateRoot, appDirectory));
  const artifacts = [];
  for (const file of [setup, nupkg, releasePath, path.join(appDirectory, 'WorldShotLog.exe')]) {
    artifacts.push({ file: path.relative(candidateRoot, file).replaceAll('\\', '/'), bytes: fs.statSync(file).size, sha256: await hashFile(file) });
  }
  return { releases, artifacts, payload };
}
function writeReviewLauncher(candidateRoot) {
  const script = [
    '$ErrorActionPreference = "Stop"',
    '$candidateRoot = [IO.Path]::GetFullPath($PSScriptRoot)',
    '$reviewProfileRoot = Join-Path $candidateRoot "ReviewProfile"',
    '$seed = Get-Content -LiteralPath (Join-Path $candidateRoot "review-profile.json") -Raw -Encoding UTF8 | ConvertFrom-Json',
    'if ([IO.Path]::GetFullPath($seed.profile) -ne $reviewProfileRoot) { throw "Keep this local review folder in its original location, or rebuild its profile." }',
    '$exe = Join-Path $candidateRoot "build\\WorldShot Log-win32-x64\\WorldShotLog.exe"',
    '$candidate = Get-Content -LiteralPath (Join-Path $candidateRoot "candidate.json") -Raw -Encoding UTF8 | ConvertFrom-Json',
    'if (-not $candidate.ok) { throw "Candidate checks have not passed." }',
    '$paths = @{',
    '  WORLDSHOT_HOME_ROOT = $reviewProfileRoot; USERPROFILE = $reviewProfileRoot;',
    '  APPDATA = (Join-Path $reviewProfileRoot "AppData\\Roaming"); WORLDSHOT_APPDATA_ROOT = (Join-Path $reviewProfileRoot "AppData\\Roaming");',
    '  LOCALAPPDATA = (Join-Path $reviewProfileRoot "AppData\\Local"); WORLDSHOT_LOCALAPPDATA_ROOT = (Join-Path $reviewProfileRoot "AppData\\Local");',
    '  TEMP = (Join-Path $reviewProfileRoot "Temp"); TMP = (Join-Path $reviewProfileRoot "Temp")',
    '}',
    '$previous = @{}',
    'try {',
    '  foreach ($key in $paths.Keys) { $previous[$key] = [Environment]::GetEnvironmentVariable($key, "Process"); [Environment]::SetEnvironmentVariable($key, $paths[$key], "Process") }',
    '  $stamp = Get-Date -Format "yyyyMMdd-HHmmss-fff"',
    '  $stdout = Join-Path $candidateRoot "review-$stamp.stdout.log"',
    '  $stderr = Join-Path $candidateRoot "review-$stamp.stderr.log"',
    '  $process = Start-Process -FilePath $exe -ArgumentList "--squirrel-firstrun" -PassThru -Wait -WindowStyle Normal -RedirectStandardOutput $stdout -RedirectStandardError $stderr',
    '  if ($process.ExitCode -ne 0) { throw "Review app exited with code $($process.ExitCode). See $stderr" }',
    '} finally {',
    '  foreach ($key in $previous.Keys) { [Environment]::SetEnvironmentVariable($key, $previous[$key], "Process") }',
    '}',
  ].join('\r\n') + '\r\n';
  fs.writeFileSync(path.join(candidateRoot, 'Open-review.ps1'), '\uFEFF' + script);
  fs.writeFileSync(path.join(candidateRoot, 'Open-review.cmd'), [
    '@echo off',
    'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Open-review.ps1"',
    'if errorlevel 1 (',
    '  echo Review launch failed. See the message and review logs above.',
    '  pause',
    '  exit /b 1',
    ')', '',
  ].join('\r\n'));
}
async function makeStage(candidateRoot) {
  assertInside(OUTPUT, candidateRoot);
  assert(/^review-candidate-[^\\/]+$/.test(path.relative(OUTPUT, candidateRoot)), 'Invalid candidate directory');
  const manifest = JSON.parse(fs.readFileSync(path.join(candidateRoot, 'candidate.json'), 'utf8'));
  assert.equal(manifest.status, 'building');
  const output = path.join(candidateRoot, 'build');
  assert(!fs.existsSync(output), 'Refusing to replace existing build output');
  const results = await require('@electron-forge/core').api.make({ dir: ROOT, outDir: output,
    platform: 'win32', arch: 'x64', interactive: false });
  fs.writeFileSync(path.join(candidateRoot, 'make-results.json'), JSON.stringify(results, null, 2));
}
async function main() {
  const args = process.argv.slice(2);
  const stage = args.find(arg => arg.startsWith('--make-stage='));
  if (stage) { assert.equal(args.length, 1); await makeStage(path.resolve(stage.slice('--make-stage='.length))); return; }
  assert(process.platform === 'win32', 'Review candidate builder supports Windows only');
  assert.equal(args.length, 1, 'Usage: node scripts/build-review-candidate.js --checks-result=<passed full check result.json>');
  assert(/^--checks(?:-result)?=/.test(args[0]), 'A passed full regression report is required');
  const checksPath = path.resolve(args[0].slice(args[0].indexOf('=') + 1));
  const checks = JSON.parse(fs.readFileSync(checksPath, 'utf8'));
  const { captureSourceHashes } = require('./lib/source-inventory');
  const sourceHashes = await captureSourceHashes(ROOT);
  const { SUITES } = require('./run-regression-checks');
  validateRegressionReport(checks, sourceHashes, Object.values(SUITES).flat().map(step => step.name));
  const sourceInputs = inputSnapshot();
  fs.mkdirSync(OUTPUT, { recursive: true });
  const candidateRoot = fs.mkdtempSync(path.join(OUTPUT, 'review-candidate-'));
  const manifestPath = path.join(candidateRoot, 'candidate.json');
  const version = require('../package.json').version;
  const gitHead = spawnSync('git', ['rev-parse', 'HEAD'], { cwd: ROOT, encoding: 'utf8', windowsHide: true });
  const report = { ok: false, status: 'building', reviewOnly: true, releaseReady: false,
    version, versionUnchanged: true, note: 'Version is unchanged from the current checkout; this is an unpublished local review candidate, not a release.',
    candidateRoot, startedAt: new Date().toISOString(), gitHead: gitHead.status === 0 ? gitHead.stdout.trim() : null,
    sourceInputs, sourceHashes, checksPath, copiedChecksResult: 'regression-result.json', checksSha256: await hashFile(checksPath), steps: [],
    nativeBefore: await hashFile(NATIVE), limits: [
      'Installer construction and payload integrity are checked; installation, installed upgrade, rollback and update-service publication are not executed.',
      'The six review images are generated fixtures. The supplied profile uses absolute local paths and should remain in this directory.',
      'Use Open-review.ps1 for isolated review. Its first-run flag suppresses the automatic startup update check; normal font resources may use the network.',
    ] };
  const persist = () => fs.writeFileSync(manifestPath, JSON.stringify(report, null, 2));
  fs.copyFileSync(checksPath, path.join(candidateRoot, 'regression-result.json'));
  persist();
  console.log(`Review candidate: ${candidateRoot}`);
  const controller = new AbortController();
  const stop = () => controller.abort();
  process.once('SIGINT', stop); process.once('SIGTERM', stop);
  async function step(name, args, timeoutMs, command = process.execPath) {
    const result = await runStep({ name, command, args, cwd: ROOT, timeoutMs,
      logPath: path.join(candidateRoot, 'logs', `${name}.log`), signal: controller.signal });
    report.steps.push(result); persist();
    console.log(`${result.ok ? 'PASS' : 'FAIL'} ${name}: ${result.logPath}`);
    assert(result.ok, `${name} failed; see ${result.logPath}`);
  }
  try {
    await step('git-diff-check', ['diff', '--check'], 30000, 'git');
    await step('patch-installer', [path.join(ROOT, 'scripts', 'patch-electron-winstaller.js')], 30000);
    await step('make', [__filename, `--make-stage=${candidateRoot}`], 1200000);
    const archive = path.join(candidateRoot, 'build', 'WorldShot Log-win32-x64', 'resources', 'app.asar');
    for (const input of sourceInputs.filter(input => input.file.startsWith('src/'))) {
      assert.equal(hashBytes(asar.extractFile(archive, path.normalize(input.file))), input.sha256, `ASAR source differs: ${input.file}`);
    }
    const exe = path.join(candidateRoot, 'build', 'WorldShot Log-win32-x64', 'WorldShotLog.exe');
    await step('packaged-smoke', [path.join(ROOT, 'scripts', 'smoke-packaged-app.js'), `--exe=${exe}`], 180000);
    const smokeOutput = fs.readFileSync(path.join(candidateRoot, 'logs', 'packaged-smoke.log'), 'utf8')
      .match(/Packaged smoke artifacts: (.+)/)?.[1]?.trim();
    assert(smokeOutput, 'Packaged smoke did not identify its report');
    assertInside(path.join(OUTPUT, 'smoke-packaged'), path.resolve(smokeOutput));
    const packagedSmoke = JSON.parse(fs.readFileSync(path.join(smokeOutput, 'result.json'), 'utf8'));
    assert(packagedSmoke.ok && packagedSmoke.exePath === exe, 'Packaged smoke report does not match candidate');
    report.packagedSmokeResult = 'packaged-smoke-result.json';
    fs.copyFileSync(path.join(smokeOutput, 'result.json'), path.join(candidateRoot, report.packagedSmokeResult));
    report.distribution = await verifyArtifacts(candidateRoot, version);
    await step('seed-review-profile', [path.join(ROOT, 'node_modules', 'electron', 'cli.js'),
      path.join(ROOT, 'scripts', 'seed-review-profile.js'), `--candidate=${candidateRoot}`], 60000);
    await step('review-profile-smoke', [path.join(ROOT, 'scripts', 'smoke-review-profile.js'), `--candidate=${candidateRoot}`], 120000);
    report.reviewProfileSmokeResult = 'review-profile-smoke/result.json';
    writeReviewLauncher(candidateRoot);
    assert.deepEqual(inputSnapshot().map(({ file, sha256 }) => ({ file, sha256 })),
      sourceInputs.map(({ file, sha256 }) => ({ file, sha256 })), 'Source changed while building the candidate');
    assert.deepEqual(await captureSourceHashes(ROOT), sourceHashes, 'Source or assets changed while building the candidate');
    report.nativeAfter = await hashFile(NATIVE);
    assert.equal(report.nativeAfter, report.nativeBefore, 'Workspace native SQLite changed during build');
    report.reviewLauncher = path.join(candidateRoot, 'Open-review.cmd');
    report.profileManifest = path.join(candidateRoot, 'review-profile.json');
    report.ok = true; report.status = 'ready-for-user-review';
  } catch (error) { report.status = 'failed'; report.error = error.stack || String(error); }
  finally {
    process.removeListener('SIGINT', stop); process.removeListener('SIGTERM', stop);
    try {
      report.nativeAfter = await hashFile(NATIVE);
      assert.equal(report.nativeAfter, report.nativeBefore, 'Workspace native SQLite changed during build');
    } catch (error) { report.ok = false; report.status = 'failed'; report.error ||= error.stack || String(error); }
    report.finishedAt = new Date().toISOString(); persist();
  }
  console.log(`Candidate result: ${manifestPath}`);
  if (!report.ok) { console.error(report.error); process.exitCode = 1; }
}
module.exports = { writeReviewLauncher, parseReleases, verifyNupkgPayload, verifyArtifacts, verifyExecutionStub, assertInside, validateRegressionReport };
if (require.main === module) main().catch(error => { console.error(error.stack || error); process.exitCode = 1; });
