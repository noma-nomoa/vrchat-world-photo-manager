'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const root = path.resolve(__dirname, '..');
const files = ['forge.config.js'];
for (const directory of ['src', 'scripts']) {
  for (const entry of fs.readdirSync(path.join(root, directory), { recursive: true, withFileTypes: true })) {
    if (entry.isFile() && entry.name.endsWith('.js')) {
      files.push(path.relative(root, path.join(entry.parentPath, entry.name)));
    }
  }
}
let failed = 0;
for (const file of files.sort()) {
  const result = spawnSync(process.execPath, ['--check', file], {
    cwd: root, encoding: 'utf8', windowsHide: true, timeout: 10000,
  });
  if (result.error || result.status !== 0) {
    failed += 1;
    console.error(`${file}\n${result.error?.message || result.stderr || result.stdout}`);
  }
}
console.log(`JavaScript syntax: ${files.length - failed}/${files.length} passed`);
process.exitCode = failed ? 1 : 0;
