'use strict';

// Run in a child: the previous WriteStream implementation emitted an uncaught
// error before reading its first response chunk. No Electron or real model is used.
const fs = require('node:fs/promises');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { createAiSubjectModelStore } = require(process.argv[2]);

async function main() {
  const root = process.argv[3];
  const bytes = Buffer.from('synthetic model');
  const model = { id: 'fixture', directoryName: 'fixture', fileName: 'model.onnx',
    metadataFileName: 'model.json', downloadUrl: 'https://fixture.invalid/model',
    sha256: createHash('sha256').update(bytes).digest('hex') };
  const store = createAiSubjectModelStore({ bundledRootPath: root, managedRootPath: root,
    legacyRootPath: root, wasmRootPath: root, models: { fixture: model }, openPath: async () => '',
    fetch: async () => {
      // Reproduce a real filesystem error after response acquisition and before
      // the first network chunk: another operation has occupied the output path.
      await fs.mkdir(path.join(root, 'fixture', 'model.onnx.download'));
      return new Response(new ReadableStream({ async pull(c) {
        await new Promise(resolve => setTimeout(resolve, 30));
        c.enqueue(bytes); c.close();
      } }));
    } });
  try {
    await store.downloadAiSubjectModel('fixture');
    throw new Error('Expected a filesystem rejection');
  } catch (error) {
    if (!['EISDIR', 'EPERM', 'EACCES'].includes(error.code)) throw error;
    process.stdout.write(JSON.stringify({ ok: true, rejected: error.code }));
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
