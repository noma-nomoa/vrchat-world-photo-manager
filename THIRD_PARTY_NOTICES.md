# Third-Party Notices

WorldShot Log uses third-party open source software and AI model assets. This
file summarizes the AI subject selection models that are bundled with, or are
referenced by, the app.

This notice is informational and is not legal advice. Check the linked upstream
licenses and terms before changing the model set.

## Bundled AI Model

### U-2-NetP

- App label: Lightweight AI
- Upstream model: BritishWerewolf/U-2-Netp
- Source: https://huggingface.co/BritishWerewolf/U-2-Netp
- Original project: https://github.com/xuebinqin/U-2-Net
- ONNX model credit: rembg, https://github.com/danielgatis/rembg
- License: Apache-2.0, matching the original U-2-Net project
- Local model path: `resources/models/u2netp/model.onnx`
- Local license copy: `resources/models/u2netp/LICENSE.txt`

## Optional Download AI Models

These models are not bundled with the app. They are downloaded only when the
user explicitly chooses to add them. They run locally with ONNX Runtime Web; the
withoutBG cloud API is not used and user images are not sent to withoutBG.

### withoutBG Snap

- App label: Standard AI
- Source: https://huggingface.co/withoutbg/snap
- Official project: https://github.com/withoutbg/withoutbg
- License: Apache-2.0
- Distribution: first-run download to the app's managed model folder
- Files: `depth_anything_v2_vits_slim.onnx`,
  `snap_matting_0.1.0.onnx`, `snap_refiner_0.1.0.onnx`

### withoutBG Focus OSS

- App label: High Quality AI
- Source: https://huggingface.co/withoutbg/focus
- Official project: https://github.com/withoutbg/withoutbg
- License: Apache-2.0
- Distribution: first-run download to the app's managed model folder
- Files: `isnet.onnx`, `depth_anything_v2_vits_slim.onnx`,
  `focus_matting_1.0.0.onnx`, `focus_refiner_1.0.0.onnx`
- Note: this model is larger and may use substantially more memory during
  local inference.

The app stores withoutBG's `LICENSE` and `THIRD_PARTY_LICENSES.md` files in the
managed model folder after download.

## Disabled AI Models

The following models were previously evaluated for standard or high quality
subject selection, but are not used in the current recommended configuration.
They are not offered for download by the app. If they already exist in a user's
managed model folder, the settings screen shows them as disabled legacy models
so they can be deleted.

### BiRefNet_lite-ONNX

- Former app label: Standard AI
- Source: https://huggingface.co/onnx-community/BiRefNet_lite-ONNX
- Base model: https://huggingface.co/ZhengPeng7/BiRefNet_lite
- Code repository: https://github.com/ZhengPeng7/BiRefNet
- Reason disabled: the model card traces the model to BiRefNet_lite, and the
  upstream BiRefNet_lite card states that it was trained on DIS-TR. DIS-TR is
  part of DIS5K, whose Terms of Use restrict the dataset to non-commercial
  research or educational use and prohibit commercial use without permission.

### BEN2-ONNX

- Former app label: High quality AI
- Source: https://huggingface.co/onnx-community/BEN2-ONNX
- Base model: https://huggingface.co/PramaLLC/BEN2
- Code repository: https://github.com/PramaLLC/BEN2
- Reason disabled: the model card traces the model to BEN2, and the upstream
  BEN2 card states that BEN2 was trained on DIS5K plus a proprietary dataset.
  DIS5K's Terms of Use restrict the dataset to non-commercial research or
  educational use and prohibit commercial use without permission.

## Review Notes

The license shown on a Hugging Face model card is not enough by itself. For AI
models, check the model repository, base model, training dataset, and any
separate terms of use before adding a model to the app. The DIS5K terms checked
for this review are published at:
https://raw.githubusercontent.com/xuebinqin/DIS/main/DIS5K-Dataset-Terms-of-Use.pdf
