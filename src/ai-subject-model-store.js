'use strict';

const path = require('node:path');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const { createHash } = require('node:crypto');
const { pathToFileURL } = require('node:url');
const { AI_MODEL_TIERS: defaultTiers, AI_SUBJECT_MODELS: defaultModels } = require('./ai-subject-model-catalog');

// Filesystem and download lifecycle. Electron provides roots, fetch and folder opening.
function createAiSubjectModelStore({
  bundledRootPath: AI_SUBJECT_MODEL_ROOT_PATH,
  managedRootPath: AI_SUBJECT_MANAGED_MODEL_ROOT_PATH,
  legacyRootPath: AI_SUBJECT_LEGACY_MODEL_ROOT_PATH,
  wasmRootPath: AI_SUBJECT_WASM_ROOT_PATH,
  fetch: fetchResponse,
  openPath,
  models: AI_SUBJECT_MODELS = defaultModels,
  tiers: AI_MODEL_TIERS = defaultTiers,
}) {
  const aiSubjectModelVerificationCache = new Map();
  const net = { fetch: fetchResponse };
  const shell = { openPath };
  const ensureDir = directoryPath => fs.mkdir(directoryPath, { recursive: true });
  async function calculateFileSha256(filePath) {
    const hash = createHash('sha256');
    for await (const chunk of fsSync.createReadStream(filePath)) {
      hash.update(chunk);
    }
    return hash.digest('hex');
  }

  function getAiSubjectModelDefinition(modelId) {
    return Object.values(AI_SUBJECT_MODELS).find((model) => model.id === modelId);
  }

  function getAiSubjectModelDirectoryPath(modelDefinition) {
    const rootPath = modelDefinition.bundled
      ? AI_SUBJECT_MODEL_ROOT_PATH
      : modelDefinition.storageRoot === 'legacy'
        ? AI_SUBJECT_LEGACY_MODEL_ROOT_PATH
        : AI_SUBJECT_MANAGED_MODEL_ROOT_PATH;
    return path.join(rootPath, modelDefinition.directoryName);
  }

  function getAiSubjectModelFileDefinitions(modelDefinition) {
    if (Array.isArray(modelDefinition.files) && modelDefinition.files.length > 0) {
      return modelDefinition.files.map((fileDefinition) => ({
        ...fileDefinition,
        role: fileDefinition.role || '',
        sha256: fileDefinition.sha256 || '',
        sizeBytes: Number(fileDefinition.sizeBytes) || 0,
      }));
    }

    return [
      {
        role: modelDefinition.role || 'model',
        fileName: modelDefinition.fileName,
        downloadUrl: modelDefinition.downloadUrl || '',
        sha256: modelDefinition.sha256 || '',
        sizeBytes: Number(modelDefinition.sizeBytes) || 0,
        sizeLabel: modelDefinition.sizeLabel || '',
      },
    ].filter((fileDefinition) => fileDefinition.fileName);
  }

  function getAiSubjectModelFilePath(modelDefinition, fileDefinition = null) {
    const targetFile = fileDefinition || getAiSubjectModelFileDefinitions(modelDefinition)[0];

    return path.join(
      getAiSubjectModelDirectoryPath(modelDefinition),
      targetFile.fileName
    );
  }

  function getAiSubjectModelMetadataFilePath(modelDefinition) {
    return path.join(
      getAiSubjectModelDirectoryPath(modelDefinition),
      modelDefinition.metadataFileName
    );
  }

  function createAiSubjectModelMetadata(modelDefinition, files = []) {
    const normalizedFiles = Array.isArray(files) ? files : [];

    return {
      id: modelDefinition.id,
      tier: modelDefinition.tier,
      modelName: modelDefinition.modelName,
      displayName: modelDefinition.displayName,
      version: modelDefinition.version || '',
      license: modelDefinition.license,
      provider: modelDefinition.provider || '',
      source: 'official',
      sourceUrl: modelDefinition.sourceUrl,
      sha256: modelDefinition.sha256 || '',
      size: normalizedFiles.reduce(
        (total, file) => total + (Number(file.size) || 0),
        0
      ),
      inputSize: modelDefinition.inputSize,
      description: modelDefinition.description,
      localOnly: modelDefinition.localOnly !== false,
      usesCloudApi: Boolean(modelDefinition.usesCloudApi),
      downloadedAt: new Date().toISOString(),
      files: normalizedFiles,
    };
  }

  async function verifyAiSubjectModelFile(
    modelDefinition,
    fileDefinition,
    modelFilePath,
    modelStat
  ) {
    const cacheKey = [
      modelDefinition.id,
      fileDefinition.fileName,
      modelFilePath,
      modelStat.size,
      Math.round(modelStat.mtimeMs),
      fileDefinition.sha256,
    ].join(':');
    const cached = aiSubjectModelVerificationCache.get(cacheKey);

    if (cached) {
      return cached;
    }

    const actualSha256 = await calculateFileSha256(modelFilePath);
    const result = {
      sha256: actualSha256,
      verified:
        actualSha256.toLowerCase() === fileDefinition.sha256.toLowerCase(),
    };
    aiSubjectModelVerificationCache.set(cacheKey, result);
    return result;
  }

  async function runAiSubjectModelInMain(payload = {}) {
    const modelDefinition = getAiSubjectModelDefinition(payload.modelId);

    if (!modelDefinition) {
      throw new Error('AIモデルが見つかりません');
    }

    if (
      modelDefinition.future ||
      modelDefinition.licenseRestricted ||
      modelDefinition.runtimeSupported === false
    ) {
      throw new Error('このAIモデルは現在の実行環境では利用できません');
    }

    throw new Error('このAIモデルは現在の実行環境では利用できません');
  }

  async function getAiSubjectModelStatus() {
    const models = [];

    for (const modelDefinition of Object.values(AI_SUBJECT_MODELS)) {
      const metadataFilePath = getAiSubjectModelMetadataFilePath(modelDefinition);
      const fileDefinitions = getAiSubjectModelFileDefinitions(modelDefinition);
      let metadata = {};
      let installed = false;
      let verified = false;
      let sizeBytes = 0;
      let message = '';
      const files = [];

      try {
        const metadataText = await fs.readFile(metadataFilePath, 'utf8').catch(() => '');
        metadata = metadataText ? JSON.parse(metadataText) : {};
      } catch {
        metadata = {};
      }

      for (const fileDefinition of fileDefinitions) {
        const filePath = getAiSubjectModelFilePath(modelDefinition, fileDefinition);
        let fileInstalled = false;
        let fileVerified = false;
        let fileSize = 0;
        let fileSha256 = '';

        try {
          const fileStat = await fs.stat(filePath);
          fileInstalled = fileStat.isFile();
          fileSize = fileInstalled ? fileStat.size : 0;

          if (fileInstalled && fileDefinition.sha256) {
            const verification = await verifyAiSubjectModelFile(
              modelDefinition,
              fileDefinition,
              filePath,
              fileStat
            );
            fileVerified = verification.verified;
            fileSha256 = verification.sha256;
          } else if (fileInstalled && !fileDefinition.sha256) {
            fileVerified = true;
          }
        } catch {
          fileInstalled = false;
          fileVerified = false;
        }

        sizeBytes += fileSize;
        files.push({
          role: fileDefinition.role || '',
          fileName: fileDefinition.fileName,
          sizeBytes: fileSize,
          expectedSizeBytes: Number(fileDefinition.sizeBytes) || 0,
          sizeLabel: fileDefinition.sizeLabel || '',
          sha256: fileSha256,
          expectedSha256: fileDefinition.sha256 || '',
          installed: fileInstalled,
          verified: fileVerified,
          fileUrl: fileInstalled ? pathToFileURL(filePath).href : '',
        });
      }

      installed = files.length > 0 && files.every((file) => file.installed);
      const partialInstalled =
        files.some((file) => file.installed) && !installed;
      verified =
        installed && files.every((file) => file.verified || !file.expectedSha256);

      if (partialInstalled) {
        message = 'モデルファイルが一部不足しています';
      } else if (installed && !verified) {
        message = 'モデルファイルの検証に失敗しました';
      } else if (!installed) {
        message = 'モデルファイルが見つかりません';
      }

      const modelFileMap = {};
      for (const file of files) {
        if (file.role) {
          modelFileMap[file.role] = file.fileUrl;
        }
      }

      models.push({
        id: modelDefinition.id,
        tier: modelDefinition.tier,
        label: modelDefinition.label,
        displayName: modelDefinition.displayName,
        modelName: modelDefinition.modelName,
        license: metadata.license || modelDefinition.license,
        sizeLabel: modelDefinition.sizeLabel,
        sizeBytes,
        bundled: Boolean(modelDefinition.bundled),
        visible: Boolean(modelDefinition.visible),
        recommended: Boolean(modelDefinition.recommended),
        deprecated: Boolean(modelDefinition.deprecated),
        licenseRestricted: Boolean(modelDefinition.licenseRestricted),
        future: Boolean(modelDefinition.future),
        runtime: modelDefinition.runtime || 'web',
        runtimeSupported: modelDefinition.runtimeSupported !== false,
        provider: metadata.provider || modelDefinition.provider || '',
        execution: metadata.execution || modelDefinition.execution || '',
        localOnly: modelDefinition.localOnly !== false,
        usesCloudApi: Boolean(modelDefinition.usesCloudApi),
        pipeline: modelDefinition.pipeline || '',
        partial: partialInstalled,
        files,
        canDownload: Boolean(
          !modelDefinition.bundled &&
            !modelDefinition.future &&
            !modelDefinition.licenseRestricted &&
            modelDefinition.runtimeSupported !== false &&
            fileDefinitions.every(
              (fileDefinition) =>
                fileDefinition.downloadUrl && fileDefinition.sha256
            )
        ),
        canDelete: Boolean(!modelDefinition.bundled && (installed || partialInstalled)),
        installed,
        verified,
        ready:
          installed &&
          verified &&
          !modelDefinition.future &&
          !modelDefinition.licenseRestricted &&
          modelDefinition.runtimeSupported !== false,
        inputSize: Number(metadata.inputSize || modelDefinition.inputSize) || 320,
        keepAspectRatio: Boolean(modelDefinition.keepAspectRatio),
        sourceUrl: metadata.sourceUrl || modelDefinition.sourceUrl,
        description: metadata.description || modelDefinition.description,
        modelUrl: installed && files[0] ? files[0].fileUrl : '',
        modelFiles: modelFileMap,
        modelDirectoryPath:
          installed || partialInstalled
            ? getAiSubjectModelDirectoryPath(modelDefinition)
            : '',
        wasmBaseUrl: pathToFileURL(`${AI_SUBJECT_WASM_ROOT_PATH}${path.sep}`).href,
        message:
          message ||
          (modelDefinition.runtimeSupported === false
            ? '現在の実行環境では検証中です'
            : modelDefinition.licenseRestricted
              ? 'ライセンス確認により利用停止中です'
            : modelDefinition.future
              ? '将来実装予定です'
            : '') ||
          (!modelDefinition.bundled && !installed
            ? '設定からダウンロードできます'
            : ''),
      });
    }

    return {
      ok: true,
      tiers: AI_MODEL_TIERS,
      models,
    };
  }

  function sendAiSubjectModelDownloadProgress(webContents, payload) {
    try {
      webContents?.send?.('ai-subject-model-download-progress', payload);
    } catch {
      // Progress events are best-effort; the download result is still returned.
    }
  }

  async function writeResponseBodyToFileWithProgress(
    response,
    destinationPath,
    {
      modelId,
      webContents,
      totalBytes = 0,
      receivedOffsetBytes = 0,
      totalExpectedBytes = 0,
      fileName = '',
    } = {}
  ) {
    const reader = response.body.getReader();
    let output;
    let receivedBytes = 0;

    try {
      // Await each write so open/write errors reject the download instead of
      // escaping as an unhandled WriteStream error while the network is idle.
      output = await fs.open(destinationPath, 'w');
      while (true) {
        const { done, value } = await reader.read();

        if (done) {
          break;
        }

        const chunk = Buffer.from(value);
        await output.writeFile(chunk);
        receivedBytes += chunk.length;

        sendAiSubjectModelDownloadProgress(webContents, {
          modelId,
          fileName,
          receivedBytes: receivedOffsetBytes + receivedBytes,
          currentFileReceivedBytes: receivedBytes,
          currentFileTotalBytes: totalBytes,
          totalBytes: totalExpectedBytes || totalBytes,
        });
      }
    } finally {
      try {
        await output?.close();
      } finally {
        await reader.cancel().catch(() => {});
        reader.releaseLock();
      }
    }
  }

  async function downloadTextFileToPath(downloadUrl, destinationPath) {
    if (!downloadUrl) {
      return false;
    }

    try {
      const response = await net.fetch(downloadUrl);

      if (!response.ok) {
        return false;
      }

      const text = await response.text();
      await fs.writeFile(destinationPath, text, 'utf8');
      return true;
    } catch {
      return false;
    }
  }

  async function downloadAiSubjectModel(modelId, webContents = null) {
    const modelDefinition = getAiSubjectModelDefinition(modelId);

    if (!modelDefinition) {
      throw new Error('AIモデルが見つかりません');
    }

    if (
      modelDefinition.bundled ||
      modelDefinition.future ||
      modelDefinition.licenseRestricted ||
      modelDefinition.runtimeSupported === false ||
      getAiSubjectModelFileDefinitions(modelDefinition).some(
        (fileDefinition) => !fileDefinition.downloadUrl || !fileDefinition.sha256
      )
    ) {
      throw new Error('このAIモデルはダウンロード対象ではありません');
    }

    const modelDirectoryPath = getAiSubjectModelDirectoryPath(modelDefinition);
    const metadataFilePath = getAiSubjectModelMetadataFilePath(modelDefinition);
    const fileDefinitions = getAiSubjectModelFileDefinitions(modelDefinition);
    const totalExpectedBytes = fileDefinitions.reduce(
      (total, fileDefinition) => total + (Number(fileDefinition.sizeBytes) || 0),
      0
    );
    let receivedOffsetBytes = 0;
    const downloadedFiles = [];

    await ensureDir(modelDirectoryPath);

    try {
      for (const fileDefinition of fileDefinitions) {
        const modelFilePath = getAiSubjectModelFilePath(
          modelDefinition,
          fileDefinition
        );
        const tempFilePath = `${modelFilePath}.download`;

        await fs.rm(tempFilePath, { force: true }).catch(() => {});

        sendAiSubjectModelDownloadProgress(webContents, {
          modelId: modelDefinition.id,
          fileName: fileDefinition.fileName,
          receivedBytes: receivedOffsetBytes,
          totalBytes: totalExpectedBytes,
        });

        const response = await net.fetch(fileDefinition.downloadUrl);

        if (!response.ok || !response.body) {
          throw new Error(
            `AIモデルをダウンロードできませんでした (${response.status || 'unknown'})`
          );
        }

        const totalBytes =
          Number(response.headers.get('content-length')) ||
          Number(fileDefinition.sizeBytes) ||
          0;
        await writeResponseBodyToFileWithProgress(response, tempFilePath, {
          modelId: modelDefinition.id,
          webContents,
          totalBytes,
          receivedOffsetBytes,
          totalExpectedBytes,
          fileName: fileDefinition.fileName,
        });

        const actualSha256 = await calculateFileSha256(tempFilePath);

        if (actualSha256.toLowerCase() !== fileDefinition.sha256.toLowerCase()) {
          await fs.rm(tempFilePath, { force: true }).catch(() => {});
          throw new Error('AIモデルの検証に失敗しました');
        }

        const modelStat = await fs.stat(tempFilePath);
        await fs.rm(modelFilePath, { force: true }).catch(() => {});
        await fs.rename(tempFilePath, modelFilePath);
        receivedOffsetBytes += modelStat.size;
        downloadedFiles.push({
          role: fileDefinition.role || '',
          fileName: fileDefinition.fileName,
          sha256: actualSha256,
          size: modelStat.size,
        });
      }
    } catch (error) {
      for (const fileDefinition of fileDefinitions) {
        await fs
          .rm(`${getAiSubjectModelFilePath(modelDefinition, fileDefinition)}.download`, {
            force: true,
          })
          .catch(() => {});
      }
      throw error;
    }

    if (Array.isArray(modelDefinition.supplementalFiles)) {
      for (const supplementalFile of modelDefinition.supplementalFiles) {
        const targetPath = path.join(modelDirectoryPath, supplementalFile.fileName);
        await downloadTextFileToPath(supplementalFile.downloadUrl, targetPath);
      }
    }

    await fs.writeFile(
      metadataFilePath,
      `${JSON.stringify(createAiSubjectModelMetadata(modelDefinition, downloadedFiles), null, 2)}\n`,
      'utf8'
    );

    return getAiSubjectModelStatus();
  }

  async function deleteAiSubjectModel(modelId) {
    const modelDefinition = getAiSubjectModelDefinition(modelId);

    if (!modelDefinition) {
      throw new Error('AIモデルが見つかりません');
    }

    if (modelDefinition.bundled) {
      throw new Error('同梱AIモデルは削除できません');
    }

    const modelDirectoryPath = path.resolve(
      getAiSubjectModelDirectoryPath(modelDefinition)
    );
    const managedRootPath = path.resolve(
      modelDefinition.storageRoot === 'legacy'
        ? AI_SUBJECT_LEGACY_MODEL_ROOT_PATH
        : AI_SUBJECT_MANAGED_MODEL_ROOT_PATH
    );

    if (
      modelDirectoryPath !== managedRootPath &&
      !modelDirectoryPath.startsWith(`${managedRootPath}${path.sep}`)
    ) {
      throw new Error('AIモデルの削除先が不正です');
    }

    for (const cacheKey of aiSubjectModelVerificationCache.keys()) {
      if (cacheKey.startsWith(`${modelDefinition.id}:`)) {
        aiSubjectModelVerificationCache.delete(cacheKey);
      }
    }

    await fs.rm(modelDirectoryPath, { recursive: true, force: true });
    return getAiSubjectModelStatus();
  }

  async function openAiSubjectModelFolder(modelId) {
    const modelDefinition = getAiSubjectModelDefinition(modelId);

    if (!modelDefinition) {
      throw new Error('AIモデルが見つかりません');
    }

    const modelDirectoryPath = getAiSubjectModelDirectoryPath(modelDefinition);
    await ensureDir(modelDirectoryPath);
    const result = await shell.openPath(modelDirectoryPath);

    if (result) {
      throw new Error(result);
    }

    return {
      ok: true,
    };
  }


  return { getAiSubjectModelStatus, downloadAiSubjectModel, deleteAiSubjectModel, openAiSubjectModelFolder, runAiSubjectModelInMain };
}

module.exports = { createAiSubjectModelStore };
