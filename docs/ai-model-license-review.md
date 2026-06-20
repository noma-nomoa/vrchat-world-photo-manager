# AI Model License Review

Date: 2026-06-20

This review checks the AI subject selection models referenced by WorldShot Log.
It follows the practical rule that a Hugging Face license tag is only one signal:
the base model, training datasets, and separate dataset terms also need to be
checked.

This document is not legal advice.

## Current Decision

- Keep U-2-NetP as the bundled Lightweight AI model.
- Add withoutBG Snap as the first-run download Standard AI model.
- Add withoutBG Focus OSS as the first-run download High Quality AI model.
- Do not offer BiRefNet_lite-ONNX, BEN2-ONNX, or BiRefNet-ONNX as downloads.
- If these models are already installed locally, show them only as disabled
  legacy models that can be deleted.

## Model Findings

### U-2-NetP

- App role: Lightweight AI
- Bundled: yes
- Source: https://huggingface.co/BritishWerewolf/U-2-Netp
- Original project: https://github.com/xuebinqin/U-2-Net
- Local license copy: `resources/models/u2netp/LICENSE.txt`

Finding: acceptable to keep. The Hugging Face model card identifies the model as
Apache-2.0 and credits rembg for the ONNX model plus the original U-2-Net
authors. The original U-2-Net repository is Apache-2.0.

Required action: keep the upstream credit and Apache-2.0 license copy in the
repository/distribution.

### withoutBG Snap

- App role: Standard AI
- Distribution: first-run download
- Source: https://huggingface.co/withoutbg/snap
- Official project: https://github.com/withoutbg/withoutbg
- License: Apache-2.0
- Files:
  - `depth_anything_v2_vits_slim.onnx`
  - `snap_matting_0.1.0.onnx`
  - `snap_refiner_0.1.0.onnx`

Finding: acceptable to offer as an optional local download. The Hugging Face
model card is Apache-2.0 and its README states that the Snap tier models are
free for commercial and non-commercial use. The README identifies Depth Anything
V2 as the depth component under Apache-2.0. The app downloads model files from
the official Hugging Face repository and verifies the Hugging Face LFS SHA-256
for each ONNX file before use.

Required action: keep local-only wording in the UI, do not use the withoutBG
cloud API, and store the upstream Apache-2.0 license and third-party notices in
the managed model folder.

### withoutBG Focus OSS

- App role: High Quality AI
- Distribution: first-run download
- Source: https://huggingface.co/withoutbg/focus
- Official project: https://github.com/withoutbg/withoutbg
- License: Apache-2.0
- Files:
  - `isnet.onnx`
  - `depth_anything_v2_vits_slim.onnx`
  - `focus_matting_1.0.0.onnx`
  - `focus_refiner_1.0.0.onnx`

Finding: acceptable to offer as an optional local download with clear resource
usage warnings. The Hugging Face model repository is tagged Apache-2.0. The
official withoutBG GitHub repository is Apache-2.0 and describes the Focus model
as local, private, and offline-capable. Its third-party notice lists Depth
Anything V2 and ISNet code components as Apache-2.0. Because ISNet is associated
with the DIS project, this review treats the model as acceptable only as
withoutBG's Apache-2.0 published model weights and does not separately include
or redistribute DIS5K dataset contents.

Required action: keep Focus as an explicit high-quality action, warn about
larger memory use, do not use the withoutBG cloud API, and store upstream
license/notice files in the managed model folder.

### BiRefNet_lite-ONNX

- Former app role: Standard AI
- Source: https://huggingface.co/onnx-community/BiRefNet_lite-ONNX
- Base model: https://huggingface.co/ZhengPeng7/BiRefNet_lite
- Code repository: https://github.com/ZhengPeng7/BiRefNet
- Hugging Face tag: MIT

Finding: do not use in the current app. The ONNX model card lists
ZhengPeng7/BiRefNet_lite as the base model. The upstream BiRefNet_lite model
card states that this model was trained on DIS-TR and validated on DIS-TEs and
DIS-VD. DIS-TR is part of DIS5K. The official DIS repository says the code and
evaluation metric are Apache-2.0, but the DIS5K dataset has a separate Terms of
Use PDF. That PDF states the dataset is available for non-commercial research or
educational use, and that commercial use is prohibited without permission.

Required action: remove download and automatic-use paths. Existing local copies
should be deletable as disabled legacy models.

### BEN2-ONNX

- Former app role: High quality AI
- Source: https://huggingface.co/onnx-community/BEN2-ONNX
- Base model: https://huggingface.co/PramaLLC/BEN2
- Code repository: https://github.com/PramaLLC/BEN2
- Hugging Face tag: MIT

Finding: do not use in the current app. The ONNX model card lists PramaLLC/BEN2
as the base model. The upstream BEN2 model card states that BEN2 was trained on
DIS5K and a proprietary segmentation dataset. Because DIS5K's separate Terms of
Use prohibit commercial use without permission, this is not a good fit for this
app's distributed default model set.

Required action: remove download and high-quality detection paths. Existing
local copies should be deletable as disabled legacy models.

### BiRefNet-ONNX

- App role: hidden comparison candidate
- Source: https://huggingface.co/onnx-community/BiRefNet-ONNX
- Base model: https://huggingface.co/ZhengPeng7/BiRefNet
- Hugging Face tag: MIT

Finding: do not expose or download. It has the same BiRefNet/DIS5K concern as
BiRefNet_lite-ONNX.

## Sources Checked

- OkojoAI license checklist:
  https://okojo.ai/blog/huggingface-license-commercial-use-guide/
- U-2-NetP model card:
  https://huggingface.co/BritishWerewolf/U-2-Netp
- U-2-Net original repository:
  https://github.com/xuebinqin/U-2-Net
- rembg repository:
  https://github.com/danielgatis/rembg
- withoutBG Snap model card:
  https://huggingface.co/withoutbg/snap
- withoutBG Focus model card:
  https://huggingface.co/withoutbg/focus
- withoutBG repository:
  https://github.com/withoutbg/withoutbg
- withoutBG third-party license notice:
  https://github.com/withoutbg/withoutbg/blob/main/THIRD_PARTY_LICENSES.md
- BiRefNet_lite-ONNX model card:
  https://huggingface.co/onnx-community/BiRefNet_lite-ONNX
- BiRefNet_lite model card:
  https://huggingface.co/ZhengPeng7/BiRefNet_lite
- BiRefNet repository:
  https://github.com/ZhengPeng7/BiRefNet
- BEN2-ONNX model card:
  https://huggingface.co/onnx-community/BEN2-ONNX
- BEN2 model card:
  https://huggingface.co/PramaLLC/BEN2
- BEN2 repository:
  https://github.com/PramaLLC/BEN2
- DIS repository and dataset terms:
  https://github.com/xuebinqin/DIS
- DIS5K Terms of Use PDF:
  https://raw.githubusercontent.com/xuebinqin/DIS/main/DIS5K-Dataset-Terms-of-Use.pdf
- DIS project page:
  https://xuebinqin.github.io/dis/index.html
