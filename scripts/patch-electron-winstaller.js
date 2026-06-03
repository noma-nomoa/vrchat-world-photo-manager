const fs = require('fs');
const path = require('path');

const signPath = path.join(
  __dirname,
  '..',
  'node_modules',
  'electron-winstaller',
  'lib',
  'sign.js'
);

const target =
  'if (!fs_extra_1.default.existsSync(BACKUP_SIGN_TOOL_PATH)) return [3 /*break*/, 3];';
const replacement =
  'if (!BACKUP_SIGN_TOOL_PATH || !fs_extra_1.default.existsSync(BACKUP_SIGN_TOOL_PATH)) return [3 /*break*/, 3];';

if (!fs.existsSync(signPath)) {
  console.warn('[patch-electron-winstaller] electron-winstaller sign.js was not found.');
  process.exit(0);
}

const source = fs.readFileSync(signPath, 'utf8');

if (source.includes(replacement)) {
  console.log('[patch-electron-winstaller] Patch already applied.');
  process.exit(0);
}

if (!source.includes(target)) {
  if (source.includes('existsSync(BACKUP_SIGN_TOOL_PATH)')) {
    console.error(
      '[patch-electron-winstaller] Unexpected resetSignTool implementation.'
    );
    process.exit(1);
  }

  console.log('[patch-electron-winstaller] Patch is not needed.');
  process.exit(0);
}

fs.writeFileSync(signPath, source.replace(target, replacement));
console.log('[patch-electron-winstaller] Patch applied.');
