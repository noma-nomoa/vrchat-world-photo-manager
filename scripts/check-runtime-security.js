'use strict';

// Run with Node: the installed Electron executable is checked in run-as-Node
// mode, without importing the application or opening an application profile.
// This is a runtime/provenance gate, not an exploit or malformed-input test.
const fs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { spawnSync } = require('node:child_process');

const EXPECTED = Object.freeze({
  electron: '45.0.0',
  chrome: '156.0.8078.48',
  node: '24.22.0',
});

// Verified against the exact release sources on 2026-10-25. Electron links
// Node's Brotli calls to Chromium's decoder; process.versions.node/brotli alone
// therefore cannot establish whether this Electron build includes the fix.
const PROVENANCE = Object.freeze({
  cve: 'CVE-2026-111497',
  verifiedOn: '2026-10-25',
  fixedCommit: '0ed4268a4401abcb4472096340477cf4af7d1a2b',
  chromiumBrotliRevision: 'cfa184a997ac19454d5a3ae97f6d600aac2445a6',
  ancestry: { status: 'ahead', aheadBy: 5, behindBy: 0 },
  sources: {
    electronDependencies: 'https://github.com/electron/electron/blob/v45.0.0/DEPS',
    sharedDecoderBuild: 'https://github.com/electron/electron/blob/v45.0.0/patches/node/build_add_gn_build_files.patch',
    chromiumBrotliRevision: 'https://chromium.googlesource.com/chromium/src/+/156.0.8078.48/third_party/brotli/README.chromium',
    chromiumDecoder: 'https://chromium.googlesource.com/chromium/src/+/156.0.8078.48/third_party/brotli/dec/decode.c',
    fixAncestry: 'https://github.com/google/brotli/compare/0ed4268a4401abcb4472096340477cf4af7d1a2b...cfa184a997ac19454d5a3ae97f6d600aac2445a6',
    securityRelease: 'https://github.com/google/brotli/releases/tag/v1.0.10',
  },
});

function inspectRuntime() {
  const actual = { ...process.versions };
  const checks = Object.entries(EXPECTED).map(([component, expected]) => ({
    component, expected, actual: actual[component] || null,
    ok: actual[component] === expected,
  }));
  const result = {
    ok: checks.every(check => check.ok),
    scope: 'Installed Electron runtime; exact versions and verified release-source provenance',
    executable: process.execPath,
    platform: process.platform,
    arch: process.arch,
    versions: actual,
    checks,
    provenance: PROVENANCE,
    validBrotliRoundTrip: 'skipped',
    malformedInputTested: false,
    patchedBehaviorTested: false,
  };

  // Do not decode anything on an unreviewed runtime. For the approved build,
  // exercise only a valid stream generated locally by its own compressor.
  if (result.ok) {
    const { brotliCompressSync, brotliDecompressSync } = require('node:zlib');
    const input = Buffer.from('WorldShot Log runtime verification: valid Brotli data.\n');
    const decoded = brotliDecompressSync(brotliCompressSync(input));
    result.validBrotliRoundTrip = decoded.equals(input) ? 'passed' : 'failed';
    result.ok = result.validBrotliRoundTrip === 'passed';
  }

  result.executableSha256 = createHash('sha256')
    .update(fs.readFileSync(process.execPath)).digest('hex');
  if (!result.ok) {
    result.error = 'Runtime does not match the reviewed release or the valid Brotli roundtrip failed. Re-review release source provenance before changing EXPECTED.';
  }
  return result;
}

function main() {
  if (process.argv.includes('--runtime-child')) {
    const result = inspectRuntime();
    process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
    process.exitCode = result.ok ? 0 : 1;
    return;
  }

  if (process.versions.electron) {
    throw new Error('Run this checker with Node, or with ELECTRON_RUN_AS_NODE=1 and --runtime-child.');
  }
  // Electron 45's package entry point can download a missing runtime. Read the
  // installed executable location directly so this check remains offline.
  const electronRoot = path.dirname(require.resolve('electron/package.json'));
  const pathFile = path.join(electronRoot, 'path.txt');
  if (!fs.existsSync(pathFile)) {
    throw new Error('Electron runtime is not installed; install dependencies before running this check.');
  }
  const electronExecutable = path.join(electronRoot, 'dist', fs.readFileSync(pathFile, 'utf8').trim());
  if (!fs.existsSync(electronExecutable)) {
    throw new Error(`Installed Electron executable is missing: ${electronExecutable}`);
  }
  const child = spawnSync(electronExecutable, [__filename, '--runtime-child'], {
    cwd: path.resolve(__dirname, '..'),
    env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
    encoding: 'utf8', windowsHide: true, timeout: 30000,
    maxBuffer: 1024 * 1024,
  });
  if (child.stdout) process.stdout.write(child.stdout);
  if (child.stderr) process.stderr.write(child.stderr);
  if (child.error) throw child.error;
  if (child.signal) throw new Error(`Runtime checker terminated with signal ${child.signal}`);
  process.exitCode = child.status === 0 ? 0 : 1;
}

module.exports = { EXPECTED, PROVENANCE };
if (require.main === module) {
  try {
    main();
  } catch (error) {
    process.stdout.write(`${JSON.stringify({ ok: false, error: error.message }, null, 2)}\n`);
    process.exitCode = 1;
  }
}
