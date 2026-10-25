'use strict';

// Execute the generated Windows entry points, including Windows PowerShell 5.1
// decoding, cmd quoting, Start-Process and its environment/exit-code handling.
// The small fixture EXE never opens Electron or any real application profile.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { writeReviewLauncher } = require('./build-review-candidate');
const { runStep } = require('./lib/regression-runner');

const windows = process.platform === 'win32';
const windowsRoot = process.env.SystemRoot || 'C:\\Windows';
const powershell = path.join(windowsRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
const cmd = path.join(windowsRoot, 'System32', 'cmd.exe');
const keys = ['WORLDSHOT_HOME_ROOT', 'USERPROFILE', 'APPDATA', 'WORLDSHOT_APPDATA_ROOT',
  'LOCALAPPDATA', 'WORLDSHOT_LOCALAPPDATA_ROOT', 'TEMP', 'TMP'];
let output;
let fixtureExe;
let sequence = 0;
const psString = value => `'${value.replaceAll("'", "''")}'`;
const json = file => JSON.parse(fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, ''));

test.before(async () => {
  if (!windows) return;
  output = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-review-launcher-'));
  fixtureExe = path.join(output, 'LauncherFixture.exe');
  const source = path.join(output, 'LauncherFixture.cs');
  fs.writeFileSync(source, `using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
public static class LauncherFixture {
  public static int Main(string[] args) {
    string root = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", ".."));
    var rows = new List<string>();
    for (int i = 0; i < args.Length; i++) rows.Add("arg." + i + "\\t" + Convert.ToBase64String(Encoding.UTF8.GetBytes(args[i])));
    foreach (string key in new string[] { ${keys.map(key => `"${key}"`).join(', ')} })
      rows.Add(key + "\\t" + Convert.ToBase64String(Encoding.UTF8.GetBytes(Environment.GetEnvironmentVariable(key) ?? "<unset>")));
    File.WriteAllLines(Path.Combine(root, "fixture-record.txt"), rows.ToArray(), new UTF8Encoding(false));
    Console.Out.WriteLine("fixture stdout: ready");
    Console.Error.WriteLine("fixture stderr: diagnostic");
    return Int32.Parse(File.ReadAllText(Path.Combine(root, "fixture-exit-code.txt")));
  }
}
`);
  const compiler = ['Framework64', 'Framework'].map(directory => path.join(windowsRoot, 'Microsoft.NET', directory, 'v4.0.30319', 'csc.exe'))
    .find(file => fs.existsSync(file));
  assert(compiler, 'Windows .NET Framework C# compiler is required for the launcher fixture');
  const result = await runStep({ name: 'compile-launcher-fixture', command: compiler,
    args: ['/nologo', '/target:winexe', `/out:${fixtureExe}`, source], cwd: output,
    logPath: path.join(output, 'compile.log'), timeoutMs: 30000 });
  assert(result.ok, fs.readFileSync(result.logPath, 'utf8'));
});

function prepare({ ready = true, moved = false, exitCode = 0 } = {}) {
  // A trailing Japanese character followed by a JSON quote reproduces the
  // original UTF-8-as-Shift-JIS parsing failure on Japanese Windows.
  const root = path.join(output, `確認 用 ${++sequence}`);
  const profile = path.join(root, 'ReviewProfile');
  const app = path.join(root, 'build', 'WorldShot Log-win32-x64');
  fs.mkdirSync(app, { recursive: true });
  const expected = {
    WORLDSHOT_HOME_ROOT: profile, USERPROFILE: profile,
    APPDATA: path.join(profile, 'AppData', 'Roaming'), WORLDSHOT_APPDATA_ROOT: path.join(profile, 'AppData', 'Roaming'),
    LOCALAPPDATA: path.join(profile, 'AppData', 'Local'), WORLDSHOT_LOCALAPPDATA_ROOT: path.join(profile, 'AppData', 'Local'),
    TEMP: path.join(profile, 'Temp'), TMP: path.join(profile, 'Temp'),
  };
  for (const directory of new Set(Object.values(expected))) fs.mkdirSync(directory, { recursive: true });
  fs.copyFileSync(fixtureExe, path.join(app, 'WorldShotLog.exe'));
  fs.writeFileSync(path.join(root, 'fixture-exit-code.txt'), String(exitCode));
  fs.writeFileSync(path.join(root, 'review-profile.json'), JSON.stringify({ ok: true,
    profile: moved ? path.join(root, 'Previously Located Profile') : profile,
    photos: [{ name: 'プリント / 文字と画像' }] }, null, 2));
  fs.writeFileSync(path.join(root, 'candidate.json'), JSON.stringify({ ok: ready, note: '確認用 / 月切り替え' }, null, 2));
  writeReviewLauncher(root);
  return { root, profile, expected };
}

async function launch(fixture, entry, scriptPath) {
  const result = await runStep({ name: `launcher-${entry}`, command: entry === 'cmd' ? cmd : powershell,
    args: entry === 'cmd' ? ['/d', '/c', 'Open-review.cmd']
      : ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath || path.join(fixture.root, 'Open-review.ps1')],
    cwd: fixture.root, logPath: path.join(fixture.root, `${entry}.log`), timeoutMs: 15000 });
  assert.equal(result.timedOut, false, `Launcher timed out: ${result.logPath}`);
  assert.equal(result.signal, null);
  assert.equal(result.error, undefined);
  assert.equal(result.terminationError, undefined);
  return { ...result, log: fs.readFileSync(result.logPath, 'utf8') };
}

function assertFixtureLaunched(fixture) {
  const rows = fs.readFileSync(path.join(fixture.root, 'fixture-record.txt'), 'utf8').trim().split(/\r?\n/);
  const actual = Object.fromEntries(rows.map(row => {
    const [key, value] = row.split('\t'); return [key, Buffer.from(value, 'base64').toString('utf8')];
  }));
  assert.deepEqual(actual, { 'arg.0': '--squirrel-firstrun', ...fixture.expected });
  for (const [stream, message] of [['stdout', 'fixture stdout: ready'], ['stderr', 'fixture stderr: diagnostic']]) {
    const logs = fs.readdirSync(fixture.root).filter(name => new RegExp(`^review-.*\\.${stream}\\.log$`).test(name));
    assert.equal(logs.length, 1, `Expected one ${stream} log`);
    assert.match(fs.readFileSync(path.join(fixture.root, logs[0]), 'utf8'), new RegExp(message));
  }
}

for (const entry of ['powershell', 'cmd']) {
  test(`${entry} launches from a Japanese path with UTF-8 JSON and an isolated profile`, { skip: !windows }, async () => {
    const fixture = prepare();
    const result = await launch(fixture, entry);
    assert.equal(result.exitCode, 0, result.log);
    assert.equal(result.ok, true);
    assertFixtureLaunched(fixture);
  });

  test(`${entry} propagates a failing child exit and preserves diagnostic logs`, { skip: !windows }, async () => {
    const fixture = prepare({ exitCode: 23 });
    const result = await launch(fixture, entry);
    assert.equal(result.exitCode, 1, result.log);
    assert.equal(result.ok, false);
    assert.match(result.log, /Review app exited with code 23/);
    if (entry === 'cmd') assert.match(result.log, /Review launch failed/);
    assertFixtureLaunched(fixture);
  });
}

for (const [label, options, entry, message] of [
  ['unverified candidate', { ready: false }, 'powershell', /Candidate checks have not passed/],
  ['moved profile', { moved: true }, 'cmd', /Keep this local review folder/],
]) {
  test(`launcher refuses ${label} before starting any executable`, { skip: !windows }, async () => {
    const fixture = prepare(options);
    const result = await launch(fixture, entry);
    assert.equal(result.exitCode, 1, result.log);
    assert.match(result.log, message);
    assert.equal(fs.existsSync(path.join(fixture.root, 'fixture-record.txt')), false);
    assert.equal(fs.readdirSync(fixture.root).some(name => /^review-.*\.(stdout|stderr)\.log$/.test(name)), false);
  });
}

for (const exitCode of [0, 23]) {
  test(`PowerShell restores its caller's environment after child exit ${exitCode}`, { skip: !windows }, async () => {
    const fixture = prepare({ exitCode });
    const wrapper = path.join(fixture.root, 'check-environment.ps1');
    // Mix pre-existing values with absent variables to exercise both restore
    // branches. All changes occur in a new PowerShell process, never Node.
    const original = Object.fromEntries(keys.map((key, index) => [key, index % 2 ? null : path.join(fixture.root, `元 の 環境 ${key}`)]));
    fs.writeFileSync(wrapper, '\uFEFF' + [
      '$ErrorActionPreference = "Stop"',
      ...Object.entries(original).map(([key, value]) => `[Environment]::SetEnvironmentVariable(${psString(key)}, ${value === null ? '$null' : psString(value)}, "Process")`),
      '$caught = $null',
      `try { & ${psString(path.join(fixture.root, 'Open-review.ps1'))} } catch { $caught = $_.Exception.Message }`,
      '$restored = @{}',
      ...keys.map(key => `$restored[${psString(key)}] = [Environment]::GetEnvironmentVariable(${psString(key)}, "Process")`),
      `@{ restored = $restored; error = $caught } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath ${psString(path.join(fixture.root, 'environment-result.json'))} -Encoding UTF8`,
    ].join('\r\n') + '\r\n');
    const result = await launch(fixture, 'powershell', wrapper);
    assert.equal(result.exitCode, 0, result.log);
    const restored = json(path.join(fixture.root, 'environment-result.json'));
    assert.deepEqual(restored.restored, original);
    if (exitCode) assert.match(restored.error, /Review app exited with code 23/);
    else assert.equal(restored.error, null);
    assertFixtureLaunched(fixture);
  });
}
