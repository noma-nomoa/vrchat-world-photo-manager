(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PhotoEditorSave = api;
})(globalThis, function () {
  'use strict';

  // Captures the save input before any asynchronous rendering. Persisted results
  // still synchronize app data after closing; editor UI completion belongs only
  // to the session that started the operation.
  function createController({ session, render, persist, synchronize, onStart, onFinish }) {
    async function save({ state, sourceMetadata, exportSettings, transparentSubjectBackground = false }) {
      if (!session.isActive(state) || state.isSaving) return { kind: 'ignored' };
      const request = session.beginRequest('save');
      const isCurrent = () => session.isRequestCurrent(request);
      const source = state.sourcePhoto;
      const settings = { ...exportSettings };
      const payload = {
        sourceFilePath: source.filePath,
        sourceFileName: source.fileName,
        sourceTakenAtTimestamp: source.takenAtTimestamp,
        sourcePhotoMetadata: JSON.parse(JSON.stringify(sourceMetadata || null)),
        outputFormat: settings.format,
        outputQuality: settings.quality,
      };
      state.isSaving = true;
      try {
        onStart(state);
        const rendered = await render(settings, { transparentSubjectBackground });
        if (!isCurrent()) return { kind: 'superseded' };
        if (!rendered?.dataUrl) throw new Error('編集結果を描画できませんでした');
        const result = await persist({ ...payload, dataUrl: rendered.dataUrl });
        if (result?.canceled) return { kind: 'canceled', result };
        if (!result?.ok) return { kind: 'failed', result };
        if (result.importResult && !result.importFailed) {
          await synchronize(result, { isCurrent });
        }
        return { kind: 'saved', result };
      } finally {
        state.isSaving = false;
        try {
          if (isCurrent()) onFinish(state);
        } finally {
          session.finishRequest(request);
        }
      }
    }
    return { save };
  }
  return Object.freeze({ createController });
});
