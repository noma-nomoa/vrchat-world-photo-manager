'use strict';

const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const { createHash } = require('node:crypto');

// Tie a successful regression run to the exact build inputs, including fixtures
// and bundled resources. Output/profile directories and review prose are excluded.
async function captureSourceHashes(root) {
  const files = ['package.json', 'package-lock.json', 'forge.config.js'];
  async function walk(relative) {
    for (const entry of await fsp.readdir(path.join(root, relative), { withFileTypes: true })) {
      const name = `${relative}/${entry.name}`;
      if (entry.isDirectory()) await walk(name);
      else if (entry.isFile()) files.push(name);
      else throw new Error(`Unsupported build input: ${name}`);
    }
  }
  for (const directory of ['src', 'scripts', 'resources', 'img']) await walk(directory);
  const hashes = {};
  for (const file of files.sort()) {
    const hash = createHash('sha256');
    for await (const chunk of fs.createReadStream(path.join(root, file))) hash.update(chunk);
    hashes[file] = hash.digest('hex');
  }
  return hashes;
}

module.exports = { captureSourceHashes };
