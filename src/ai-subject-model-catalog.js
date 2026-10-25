'use strict';

const AI_MODEL_TIERS = Object.freeze({
  lightweight: {
    id: 'lightweight',
    label: '軽量AI',
    defaultModelId: 'u2netp',
    visible: true,
  },
  standard: {
    id: 'standard',
    label: '標準AI',
    defaultModelId: 'withoutbg-snap',
    visible: true,
  },
  highQuality: {
    id: 'high-quality',
    label: '高精度AI',
    defaultModelId: 'withoutbg-focus',
    visible: true,
  },
  legacy: {
    id: 'legacy',
    label: '旧モデル',
    defaultModelId: 'u2net-standard',
    visible: false,
  },
});
const AI_SUBJECT_MODELS = Object.freeze({
  u2netp: {
    id: 'u2netp',
    tier: 'lightweight',
    label: 'U-2-NetP',
    displayName: '軽量AI',
    modelName: 'U-2-Netp',
    license: 'Apache-2.0',
    sizeLabel: '約5MB',
    directoryName: 'u2netp',
    fileName: 'model.onnx',
    metadataFileName: 'model.json',
    sourceUrl: 'https://huggingface.co/BritishWerewolf/U-2-Netp',
    sha256: '309c8469258dda742793dce0ebea8e6dd393174f89934733ecc8b14c76f4ddd8',
    inputSize: 320,
    keepAspectRatio: true,
    bundled: true,
    visible: true,
    recommended: true,
    description:
      'すぐ使える高速モデルです。複雑な髪型・衣装・羽・尻尾などでは精度が落ちる場合があります。',
  },
  'withoutbg-snap': {
    id: 'withoutbg-snap',
    tier: 'standard',
    label: 'withoutBG Snap',
    displayName: '標準AI',
    modelName: 'withoutBG Snap',
    license: 'Apache-2.0',
    sizeLabel: '約140MB',
    directoryName: 'withoutbg-snap',
    metadataFileName: 'model.json',
    sourceUrl: 'https://huggingface.co/withoutbg/snap',
    provider: 'withoutBG',
    execution: 'onnxruntime-web',
    localOnly: true,
    usesCloudApi: false,
    version: '0.1.0',
    inputSize: 518,
    keepAspectRatio: false,
    bundled: false,
    visible: true,
    recommended: true,
    pipeline: 'withoutbg-snap',
    files: [
      {
        role: 'depth',
        fileName: 'depth_anything_v2_vits_slim.onnx',
        downloadUrl:
          'https://huggingface.co/withoutbg/snap/resolve/main/depth_anything_v2_vits_slim.onnx',
        sha256:
          '396bc234301510f59fd45ada24cbb72cdc7cf201f6578dfce76f42b67bc609f7',
        sizeBytes: 98985353,
        sizeLabel: '約99MB',
      },
      {
        role: 'matting',
        fileName: 'snap_matting_0.1.0.onnx',
        downloadUrl:
          'https://huggingface.co/withoutbg/snap/resolve/main/snap_matting_0.1.0.onnx',
        sha256:
          '094d9b674939cf0f75edc1ddb768a786345205564e01a727f017a5457335197d',
        sizeBytes: 26831092,
        sizeLabel: '約26.8MB',
      },
      {
        role: 'refiner',
        fileName: 'snap_refiner_0.1.0.onnx',
        downloadUrl:
          'https://huggingface.co/withoutbg/snap/resolve/main/snap_refiner_0.1.0.onnx',
        sha256:
          '9c10a55fcb01f871b35acc5db03dc30a8c92a15c0e445d4d8ba3c8f4da3c3a80',
        sizeBytes: 14462962,
        sizeLabel: '約14.5MB',
      },
    ],
    supplementalFiles: [
      {
        fileName: 'LICENSE.txt',
        downloadUrl:
          'https://raw.githubusercontent.com/withoutbg/withoutbg/main/LICENSE',
      },
      {
        fileName: 'THIRD_PARTY_LICENSES.md',
        downloadUrl:
          'https://raw.githubusercontent.com/withoutbg/withoutbg/main/THIRD_PARTY_LICENSES.md',
      },
      {
        fileName: 'README.md',
        downloadUrl: 'https://huggingface.co/withoutbg/snap/raw/main/README.md',
      },
    ],
    description:
      'VRChatアバター向けの標準モデルです。軽量AIよりも、髪型・衣装・装飾・羽・尻尾などの検出精度向上を狙います。',
  },
  'withoutbg-focus': {
    id: 'withoutbg-focus',
    tier: 'high-quality',
    label: 'withoutBG Focus OSS',
    displayName: '高精度AI',
    modelName: 'withoutBG Focus OSS',
    license: 'Apache-2.0',
    sizeLabel: '約320MB',
    directoryName: 'withoutbg-focus',
    metadataFileName: 'model.json',
    sourceUrl: 'https://huggingface.co/withoutbg/focus',
    provider: 'withoutBG',
    execution: 'onnxruntime-web',
    localOnly: true,
    usesCloudApi: false,
    version: '1.0.0',
    inputSize: 518,
    keepAspectRatio: false,
    bundled: false,
    visible: true,
    recommended: false,
    pipeline: 'withoutbg-focus',
    files: [
      {
        role: 'isnet',
        fileName: 'isnet.onnx',
        downloadUrl:
          'https://huggingface.co/withoutbg/focus/resolve/main/isnet.onnx',
        sha256:
          'c3335f04d3aa1f586974cee6310834680498f2f0469319e3821c3995ed61f64e',
        sizeBytes: 177268929,
        sizeLabel: '約177MB',
      },
      {
        role: 'depth',
        fileName: 'depth_anything_v2_vits_slim.onnx',
        downloadUrl:
          'https://huggingface.co/withoutbg/focus/resolve/main/depth_anything_v2_vits_slim.onnx',
        sha256:
          '396bc234301510f59fd45ada24cbb72cdc7cf201f6578dfce76f42b67bc609f7',
        sizeBytes: 98985353,
        sizeLabel: '約99MB',
      },
      {
        role: 'matting',
        fileName: 'focus_matting_1.0.0.onnx',
        downloadUrl:
          'https://huggingface.co/withoutbg/focus/resolve/main/focus_matting_1.0.0.onnx',
        sha256:
          '2ac4d48e0c9d670f1146375dffe6cd5e68c386e41bc333753360f63cc70f7ee7',
        sizeBytes: 26831672,
        sizeLabel: '約26.8MB',
      },
      {
        role: 'refiner',
        fileName: 'focus_refiner_1.0.0.onnx',
        downloadUrl:
          'https://huggingface.co/withoutbg/focus/resolve/main/focus_refiner_1.0.0.onnx',
        sha256:
          'ca5989d91539cad5eac1d6380f80811844488a988380f0541f67082af3a7847b',
        sizeBytes: 14462382,
        sizeLabel: '約14.5MB',
      },
    ],
    supplementalFiles: [
      {
        fileName: 'LICENSE.txt',
        downloadUrl:
          'https://raw.githubusercontent.com/withoutbg/withoutbg/main/LICENSE',
      },
      {
        fileName: 'THIRD_PARTY_LICENSES.md',
        downloadUrl:
          'https://raw.githubusercontent.com/withoutbg/withoutbg/main/THIRD_PARTY_LICENSES.md',
      },
      {
        fileName: 'README.md',
        downloadUrl: 'https://huggingface.co/withoutbg/focus/raw/main/README.md',
      },
    ],
    description:
      '境界をよりきれいに検出する上位モデルです。背景ぼかし、被写体だけ補正、文字や画像を被写体に合わせる加工に向いています。処理時間やメモリ使用量が大きくなる場合があります。',
  },
  'birefnet-lite-onnx': {
    id: 'birefnet-lite-onnx',
    tier: 'standard',
    label: 'BiRefNet_lite-ONNX',
    displayName: '標準AI',
    modelName: 'BiRefNet_lite-ONNX',
    license: 'MIT',
    sizeLabel: '約115MB',
    directoryName: 'birefnet-lite-onnx',
    fileName: 'model.onnx',
    metadataFileName: 'model.json',
    sourceUrl: 'https://huggingface.co/onnx-community/BiRefNet_lite-ONNX',
    downloadUrl: '',
    sha256: 'd39b897ceb16ae654c1731f3dba0cf9b368d9cae74b5a57459b455cc8bfec402',
    inputSize: 1024,
    keepAspectRatio: false,
    runtime: 'node',
    bundled: false,
    visible: false,
    recommended: false,
    deprecated: true,
    licenseRestricted: true,
    description:
      '学習データセットの商用利用条件を確認した結果、現在の推奨構成では使用しません。必要に応じて削除できます。',
  },
  'ben2-onnx': {
    id: 'ben2-onnx',
    tier: 'high-quality',
    label: 'BEN2-ONNX',
    displayName: '高精度AI',
    modelName: 'BEN2-ONNX',
    license: 'MIT',
    sizeLabel: '約219MB',
    directoryName: 'ben2-onnx',
    fileName: 'model.onnx',
    metadataFileName: 'model.json',
    sourceUrl: 'https://huggingface.co/onnx-community/BEN2-ONNX',
    downloadUrl: '',
    sha256: 'dfdc25f421f32a0d1268e0f2ff2153d340e8f1d52d3dd16f5dc33c1ce85cedf1',
    inputSize: 1024,
    keepAspectRatio: false,
    runtime: 'node',
    bundled: false,
    visible: false,
    recommended: false,
    deprecated: true,
    licenseRestricted: true,
    description:
      '学習データセットの商用利用条件を確認した結果、現在の推奨構成では使用しません。必要に応じて削除できます。',
  },
  'birefnet-onnx': {
    id: 'birefnet-onnx',
    tier: 'high-quality',
    label: 'BiRefNet-ONNX',
    displayName: '高精度AI候補',
    modelName: 'BiRefNet-ONNX',
    license: 'MIT',
    sizeLabel: '約490MB(fp16)',
    directoryName: 'birefnet-onnx',
    fileName: 'model.onnx',
    metadataFileName: 'model.json',
    sourceUrl: 'https://huggingface.co/onnx-community/BiRefNet-ONNX',
    downloadUrl: '',
    sha256: '3654c741eb80bd926ada8fed1713b506ccf8d30eb1f6487e87eb9f234f33df09',
    inputSize: 1024,
    keepAspectRatio: false,
    bundled: false,
    visible: false,
    recommended: false,
    deprecated: true,
    licenseRestricted: true,
    description:
      '学習データセットの商用利用条件を確認した結果、現在の推奨構成では使用しません。',
  },
  'u2net-standard': {
    id: 'u2net-standard',
    tier: 'legacy',
    label: 'U-2-Net standard',
    displayName: '旧モデル',
    modelName: 'U-2-Net',
    license: 'Apache-2.0',
    sizeLabel: '約176MB',
    directoryName: 'u2net',
    fileName: 'model.onnx',
    metadataFileName: 'model.json',
    sourceUrl: 'https://huggingface.co/BritishWerewolf/U-2-Net',
    downloadUrl: '',
    sha256: '8d10d2f3bb75ae3b6d527c77944fc5e7dcd94b29809d47a739a7a728a912b491',
    inputSize: 320,
    keepAspectRatio: true,
    bundled: false,
    visible: false,
    recommended: false,
    deprecated: true,
    storageRoot: 'legacy',
    description:
      '旧構成で高精度AIとして扱っていたモデルです。現在の推奨構成では使用しません。',
  },
});

module.exports = { AI_MODEL_TIERS, AI_SUBJECT_MODELS };
