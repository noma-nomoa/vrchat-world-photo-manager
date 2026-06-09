const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn } = require('node:child_process');

const REPO_ROOT = path.resolve(__dirname, '..');
const DEFAULT_EXE_PATH = path.join(
  REPO_ROOT,
  'out',
  'WorldShot Log-win32-x64',
  'WorldShotLog.exe'
);
const APP_DISPLAY_NAME = 'WorldShot Log';

function getArgValue(name) {
  const prefix = `${name}=`;
  const match = process.argv.slice(2).find((arg) => arg.startsWith(prefix));
  return match ? match.slice(prefix.length) : '';
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function main() {
  const exePath = path.resolve(getArgValue('--exe') || DEFAULT_EXE_PATH);
  const waitMs = Number.parseInt(getArgValue('--wait-ms') || '9000', 10);

  if (!fs.existsSync(exePath)) {
    throw new Error(`Packaged app was not found: ${exePath}`);
  }

  const runRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'worldshot-packaged-smoke-'));
  const appData = path.join(runRoot, 'AppData', 'Roaming');
  const localAppData = path.join(runRoot, 'AppData', 'Local');
  ensureDir(appData);
  ensureDir(localAppData);

  const child = spawn(exePath, [], {
    env: {
      ...process.env,
      APPDATA: appData,
      LOCALAPPDATA: localAppData,
      WORLDSHOT_APPDATA_ROOT: appData,
      WORLDSHOT_LOCALAPPDATA_ROOT: localAppData,
    },
    stdio: 'ignore',
    windowsHide: true,
  });

  let exitInfo = null;
  child.once('exit', (code, signal) => {
    exitInfo = { code, signal };
  });

  await wait(waitMs);

  if (exitInfo && exitInfo.code !== 0) {
    throw new Error(
      `Packaged app exited during smoke test: ${JSON.stringify(exitInfo)}`
    );
  }

  const expectedDbPath = path.join(appData, APP_DISPLAY_NAME, 'data', 'app.sqlite');
  const dbExists = fs.existsSync(expectedDbPath);

  if (!exitInfo) {
    child.kill();
    await wait(1000);
  }

  if (!dbExists) {
    throw new Error(`Packaged app did not create the expected DB: ${expectedDbPath}`);
  }

  console.log(`Packaged app smoke passed: ${exePath}`);
  console.log(`Isolated app data: ${runRoot}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
