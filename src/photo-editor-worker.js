'use strict';

importScripts('photo-editor-color-core.js');

const PHOTO_EDITOR_RADIAL_BLUR_MIN_RADIUS = 0.08;
const PHOTO_EDITOR_RADIAL_BLUR_MAX_RADIUS = 1.12;
const PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER = 0.04;
const PHOTO_EDITOR_MASK_OUTSIDE_MARGIN = 1.25;

function clampNumber(value, min, max, fallback = 0) {
  const numericValue = Number(value);
  return Number.isFinite(numericValue)
    ? Math.max(min, Math.min(max, numericValue))
    : fallback;
}

function clampColorChannel(value) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function smoothstep(edge0, edge1, value) {
  if (edge0 === edge1) {
    return value < edge0 ? 0 : 1;
  }

  const t = clampNumber((value - edge0) / (edge1 - edge0), 0, 1, 0);
  return t * t * (3 - 2 * t);
}

function getDeterministicNoise(x, y) {
  return globalThis.PhotoEditorColorCore.getDeterministicNoise(x, y);
}

function clonePhotoEditorCurvePoints(points) {
  return globalThis.PhotoEditorColorCore.clonePhotoEditorCurvePoints(points);
}

function normalizePhotoEditorCurveState(curve = {}) {
  return globalThis.PhotoEditorColorCore.normalizePhotoEditorCurveState(curve);
}

function getPhotoEditorCurveValue(points, inputValue) {
  return globalThis.PhotoEditorColorCore.getPhotoEditorCurveValue(points, inputValue);
}

function buildPhotoEditorCurveLookup(points, { outputScale = 255 } = {}) {
  return globalThis.PhotoEditorColorCore.buildPhotoEditorCurveLookup(points, { outputScale });
}

function getPhotoEditorCurveLookupIndex(value, scale = 255) {
  return globalThis.PhotoEditorColorCore.getPhotoEditorCurveLookupIndex(value, scale);
}

function isPhotoEditorCurvePointsIdentity(points) {
  return globalThis.PhotoEditorColorCore.isPhotoEditorCurvePointsIdentity(points);
}

function createPhotoEditorCurveLookups(curveState) {
  return globalThis.PhotoEditorColorCore.createPhotoEditorCurveLookups(curveState);
}

function convertRgbToHsv(red, green, blue) {
  return globalThis.PhotoEditorColorCore.convertRgbToHsv(red, green, blue);
}

function convertHsvToRgb(hue, saturation, value) {
  return globalThis.PhotoEditorColorCore.convertHsvToRgb(hue, saturation, value);
}

function applyPhotoEditorCurveStateToColor(red, green, blue, curveLookups) {
  return globalThis.PhotoEditorColorCore.applyPhotoEditorCurveStateToColor(red, green, blue, curveLookups);
}

function getPhotoEditorAdjustmentRenderParams(values, curveState = null) {
  return globalThis.PhotoEditorColorCore.getPhotoEditorAdjustmentRenderParams(values || {}, curveState);
}

function applyPhotoEditorPixelAdjustmentsToImageData(imageData, width, params) {
  return globalThis.PhotoEditorColorCore.applyPhotoEditorPixelAdjustmentsToImageData(imageData, width, params);
}

function applyPhotoEditorDenoiseToCanvas(ctx, width, height, value) {
  const denoise = clampNumber(value, 0, 100, 0);

  if (denoise <= 0) {
    return;
  }

  const tempCanvas = new OffscreenCanvas(width, height);
  const tempCtx = tempCanvas.getContext('2d');

  if (!tempCtx) {
    return;
  }

  tempCtx.filter = 'blur(1.4px)';
  tempCtx.drawImage(ctx.canvas, 0, 0);
  tempCtx.filter = 'none';
  ctx.save();
  ctx.globalAlpha = Math.min(0.55, denoise / 170);
  ctx.drawImage(tempCanvas, 0, 0);
  ctx.restore();
}

function applyPhotoEditorClarityToCanvas(ctx, width, height, value) {
  const clarity = clampNumber(value, -100, 100, 0);

  if (clarity === 0) {
    return;
  }

  const tempCanvas = new OffscreenCanvas(width, height);
  const tempCtx = tempCanvas.getContext('2d', { willReadFrequently: true });

  if (!tempCtx) {
    return;
  }

  tempCtx.filter = 'blur(2px)';
  tempCtx.drawImage(ctx.canvas, 0, 0);
  tempCtx.filter = 'none';

  if (clarity < 0) {
    ctx.save();
    ctx.globalAlpha = Math.min(0.82, Math.abs(clarity) / 100);
    ctx.drawImage(tempCanvas, 0, 0);
    ctx.restore();
    return;
  }

  let sourceData;
  let blurredData;

  try {
    sourceData = ctx.getImageData(0, 0, width, height);
    blurredData = tempCtx.getImageData(0, 0, width, height);
  } catch {
    return;
  }

  const source = sourceData.data;
  const blurred = blurredData.data;
  const amount = (clarity / 100) * 1.65;

  for (let index = 0; index < source.length; index += 4) {
    source[index] = clampColorChannel(
      source[index] + (source[index] - blurred[index]) * amount
    );
    source[index + 1] = clampColorChannel(
      source[index + 1] + (source[index + 1] - blurred[index + 1]) * amount
    );
    source[index + 2] = clampColorChannel(
      source[index + 2] + (source[index + 2] - blurred[index + 2]) * amount
    );
  }

  ctx.putImageData(sourceData, 0, 0);
}

function applyPhotoEditorTextureToCanvas(ctx, width, height, value) {
  const texture = clampNumber(value, -100, 100, 0);

  if (texture === 0) {
    return;
  }

  const tempCanvas = new OffscreenCanvas(width, height);
  const tempCtx = tempCanvas.getContext('2d', { willReadFrequently: true });

  if (!tempCtx) {
    return;
  }

  tempCtx.filter = 'blur(0.85px)';
  tempCtx.drawImage(ctx.canvas, 0, 0);
  tempCtx.filter = 'none';

  if (texture < 0) {
    ctx.save();
    ctx.globalAlpha = Math.min(0.64, Math.abs(texture) / 145);
    ctx.drawImage(tempCanvas, 0, 0);
    ctx.restore();
    return;
  }

  let sourceData;
  let blurredData;

  try {
    sourceData = ctx.getImageData(0, 0, width, height);
    blurredData = tempCtx.getImageData(0, 0, width, height);
  } catch {
    return;
  }

  const source = sourceData.data;
  const blurred = blurredData.data;
  const amount = (texture / 100) * 1.05;

  for (let index = 0; index < source.length; index += 4) {
    const detailRed = source[index] - blurred[index];
    const detailGreen = source[index + 1] - blurred[index + 1];
    const detailBlue = source[index + 2] - blurred[index + 2];
    const detailLuma = 0.2126 * detailRed + 0.7152 * detailGreen + 0.0722 * detailBlue;

    source[index] = clampColorChannel(source[index] + detailLuma * amount);
    source[index + 1] = clampColorChannel(source[index + 1] + detailLuma * amount);
    source[index + 2] = clampColorChannel(source[index + 2] + detailLuma * amount);
  }

  ctx.putImageData(sourceData, 0, 0);
}

function applyPhotoEditorSharpnessToCanvas(ctx, width, height, value) {
  const sharpness = clampNumber(value, 0, 100, 0);

  if (sharpness <= 0) {
    return;
  }

  const tempCanvas = new OffscreenCanvas(width, height);
  const tempCtx = tempCanvas.getContext('2d', { willReadFrequently: true });

  if (!tempCtx) {
    return;
  }

  tempCtx.filter = 'blur(1.2px)';
  tempCtx.drawImage(ctx.canvas, 0, 0);
  tempCtx.filter = 'none';

  let sourceData;
  let blurredData;

  try {
    sourceData = ctx.getImageData(0, 0, width, height);
    blurredData = tempCtx.getImageData(0, 0, width, height);
  } catch {
    return;
  }

  const source = sourceData.data;
  const blurred = blurredData.data;
  const amount = (sharpness / 100) * 1.15;

  for (let index = 0; index < source.length; index += 4) {
    source[index] = clampColorChannel(
      source[index] + (source[index] - blurred[index]) * amount
    );
    source[index + 1] = clampColorChannel(
      source[index + 1] + (source[index + 1] - blurred[index + 1]) * amount
    );
    source[index + 2] = clampColorChannel(
      source[index + 2] + (source[index + 2] - blurred[index + 2]) * amount
    );
  }

  ctx.putImageData(sourceData, 0, 0);
}

function applyPhotoEditorVignette(ctx, width, height, value) {
  const vignette = clampNumber(value, -100, 100, 0);

  if (vignette === 0) {
    return;
  }

  const centerX = width / 2;
  const centerY = height / 2;
  const radius = Math.sqrt(centerX * centerX + centerY * centerY);
  const gradient = ctx.createRadialGradient(
    centerX,
    centerY,
    radius * 0.28,
    centerX,
    centerY,
    radius
  );
  const alpha = Math.min(0.72, Math.abs(vignette) / 100 * 0.72);
  const color = vignette > 0
    ? `rgba(255, 255, 255, ${alpha})`
    : `rgba(0, 0, 0, ${alpha})`;

  gradient.addColorStop(0, 'rgba(0, 0, 0, 0)');
  gradient.addColorStop(1, color);
  ctx.save();
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, width, height);
  ctx.restore();
}

function applyPhotoEditorAdjustmentsToCanvas(ctx, width, height, values, curveState) {
  const renderParams = getPhotoEditorAdjustmentRenderParams(values || {}, curveState);

  if (renderParams.hasPixelWork) {
    let imageData;

    try {
      imageData = ctx.getImageData(0, 0, width, height);
    } catch {
      return;
    }

    applyPhotoEditorPixelAdjustmentsToImageData(imageData, width, renderParams);
    ctx.putImageData(imageData, 0, 0);
  }

  applyPhotoEditorDenoiseToCanvas(ctx, width, height, values?.denoise);
  applyPhotoEditorClarityToCanvas(ctx, width, height, values?.clarity);
  applyPhotoEditorTextureToCanvas(ctx, width, height, values?.texture);
  applyPhotoEditorSharpnessToCanvas(ctx, width, height, values?.sharpness);
  applyPhotoEditorVignette(ctx, width, height, values?.vignette);
}

function normalizePhotoEditorBlurState(blur = {}) {
  const radius = clampNumber(
    blur?.radius,
    PHOTO_EDITOR_RADIAL_BLUR_MIN_RADIUS,
    PHOTO_EDITOR_RADIAL_BLUR_MAX_RADIUS - PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER,
    0.34
  );
  const outerRadius = clampNumber(
    blur?.outerRadius,
    radius + PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER,
    PHOTO_EDITOR_RADIAL_BLUR_MAX_RADIUS,
    Math.max(radius + 0.18, 0.52)
  );

  return {
    mode: blur?.mode === 'radial' ? 'radial' : 'full',
    amount: clampNumber(blur?.amount, 0, 100, 0),
    centerX: clampNumber(blur?.centerX, 0, 1, 0.5),
    centerY: clampNumber(blur?.centerY, 0, 1, 0.5),
    radius,
    outerRadius,
  };
}

function drawPhotoEditorBlurPaddingSource(targetCtx, sourceCanvas, padding, width, height) {
  targetCtx.drawImage(sourceCanvas, padding, padding);
  targetCtx.drawImage(sourceCanvas, 0, 0, width, 1, padding, 0, width, padding);
  targetCtx.drawImage(sourceCanvas, 0, height - 1, width, 1, padding, padding + height, width, padding);
  targetCtx.drawImage(sourceCanvas, 0, 0, 1, height, 0, padding, padding, height);
  targetCtx.drawImage(sourceCanvas, width - 1, 0, 1, height, padding + width, padding, padding, height);
  targetCtx.drawImage(sourceCanvas, 0, 0, 1, 1, 0, 0, padding, padding);
  targetCtx.drawImage(sourceCanvas, width - 1, 0, 1, 1, padding + width, 0, padding, padding);
  targetCtx.drawImage(sourceCanvas, 0, height - 1, 1, 1, 0, padding + height, padding, padding);
  targetCtx.drawImage(sourceCanvas, width - 1, height - 1, 1, 1, padding + width, padding + height, padding, padding);
}

function createPhotoEditorBlurredCanvas(
  ctx,
  width,
  height,
  amount,
  { isInteractivePreview = false } = {}
) {
  const blurAmount = clampNumber(amount, 0, 100, 0);

  if (blurAmount <= 0) {
    return null;
  }

  const blurRadius = Math.max(
    1,
    Math.min(
      isInteractivePreview ? 24 : 42,
      Math.round(blurAmount * (isInteractivePreview ? 0.3 : 0.42))
    )
  );
  const padding = Math.ceil(blurRadius * 2);
  const paddedWidth = width + padding * 2;
  const paddedHeight = height + padding * 2;
  const paddedCanvas = new OffscreenCanvas(paddedWidth, paddedHeight);
  const paddedCtx = paddedCanvas.getContext('2d');

  if (!paddedCtx) {
    return null;
  }

  drawPhotoEditorBlurPaddingSource(paddedCtx, ctx.canvas, padding, width, height);

  const paddedBlurCanvas = new OffscreenCanvas(paddedWidth, paddedHeight);
  const paddedBlurCtx = paddedBlurCanvas.getContext('2d');

  if (!paddedBlurCtx) {
    return null;
  }

  paddedBlurCtx.filter = `blur(${blurRadius}px)`;
  paddedBlurCtx.drawImage(paddedCanvas, 0, 0);
  paddedBlurCtx.filter = 'none';

  const tempCanvas = new OffscreenCanvas(width, height);
  const tempCtx = tempCanvas.getContext('2d');

  if (!tempCtx) {
    return null;
  }

  tempCtx.drawImage(paddedBlurCanvas, padding, padding, width, height, 0, 0, width, height);

  return {
    canvas: tempCanvas,
    ctx: tempCtx,
  };
}

function getPhotoEditorBlurGeometry(width, height, blurState) {
  const blur = normalizePhotoEditorBlurState(blurState);
  const minEdge = Math.min(width, height);
  const centerX = clampNumber(blur.centerX, 0, 1, 0.5) * width;
  const centerY = clampNumber(blur.centerY, 0, 1, 0.5) * height;
  const radius = clampNumber(
    blur.radius,
    PHOTO_EDITOR_RADIAL_BLUR_MIN_RADIUS,
    PHOTO_EDITOR_RADIAL_BLUR_MAX_RADIUS - PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER,
    0.34
  ) * minEdge;
  const outerRadius = clampNumber(
    blur.outerRadius,
    blur.radius + PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER,
    PHOTO_EDITOR_RADIAL_BLUR_MAX_RADIUS,
    Math.max(blur.radius + 0.18, 0.52)
  ) * minEdge;

  return {
    centerX,
    centerY,
    radius,
    outerRadius: Math.max(radius + minEdge * PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER, outerRadius),
  };
}

function applyPhotoEditorBlurEffectToCanvas(
  ctx,
  width,
  height,
  blurState = null,
  { isInteractivePreview = false } = {}
) {
  const blur = normalizePhotoEditorBlurState(blurState);
  const blurred = createPhotoEditorBlurredCanvas(ctx, width, height, blur.amount, {
    isInteractivePreview,
  });

  if (!blurred) {
    return;
  }

  if (blur.mode === 'full') {
    ctx.drawImage(blurred.canvas, 0, 0);
    return;
  }

  const { centerX, centerY, radius, outerRadius } = getPhotoEditorBlurGeometry(
    width,
    height,
    blur
  );
  const maskCanvas = new OffscreenCanvas(width, height);
  const maskCtx = maskCanvas.getContext('2d');

  if (!maskCtx) {
    return;
  }

  const gradient = maskCtx.createRadialGradient(centerX, centerY, radius, centerX, centerY, outerRadius);
  gradient.addColorStop(0, 'rgba(0, 0, 0, 0)');
  gradient.addColorStop(1, 'rgba(0, 0, 0, 1)');
  maskCtx.fillStyle = gradient;
  maskCtx.fillRect(0, 0, width, height);
  blurred.ctx.globalCompositeOperation = 'destination-in';
  blurred.ctx.drawImage(maskCanvas, 0, 0);
  blurred.ctx.globalCompositeOperation = 'source-over';
  ctx.drawImage(blurred.canvas, 0, 0);
}

function getPhotoEditorMaskCoordinateBounds(allowOutside = false) {
  return allowOutside
    ? {
        min: -PHOTO_EDITOR_MASK_OUTSIDE_MARGIN,
        max: 1 + PHOTO_EDITOR_MASK_OUTSIDE_MARGIN,
      }
    : { min: 0, max: 1 };
}

function canPhotoEditorMaskExtendOutside(maskOrShape) {
  const shape =
    typeof maskOrShape === 'string' ? maskOrShape : maskOrShape?.shape;
  return shape === 'ellipse' || Boolean(maskOrShape?.allowOutside);
}

function normalizePhotoEditorMaskRotation(rotation) {
  const numericRotation = Number(rotation) || 0;
  let normalizedRotation = ((numericRotation % 360) + 360) % 360;

  if (normalizedRotation > 180) {
    normalizedRotation -= 360;
  }

  return normalizedRotation;
}

function getPhotoEditorMaskRotation(mask) {
  return normalizePhotoEditorMaskRotation(mask?.rotation);
}

function rotatePointAroundCenter(point, center, rotationRadians) {
  const cos = Math.cos(rotationRadians);
  const sin = Math.sin(rotationRadians);
  const deltaX = point.x - center.x;
  const deltaY = point.y - center.y;

  return {
    x: center.x + deltaX * cos - deltaY * sin,
    y: center.y + deltaX * sin + deltaY * cos,
  };
}

function rotateNormalizedPointAroundCenterForSize(point, center, rotationRadians, width, height) {
  const rotatedPoint = rotatePointAroundCenter(
    {
      x: point.x * width,
      y: point.y * height,
    },
    {
      x: center.x * width,
      y: center.y * height,
    },
    rotationRadians
  );

  return {
    x: rotatedPoint.x / Math.max(1, width),
    y: rotatedPoint.y / Math.max(1, height),
  };
}

function getPhotoEditorMaskNormalizedRect(mask) {
  const rect = mask?.rect || {};

  return {
    x: Number(rect.x) || 0,
    y: Number(rect.y) || 0,
    width: Math.max(0, Number(rect.width) || 0),
    height: Math.max(0, Number(rect.height) || 0),
  };
}

function getPhotoEditorMaskRectCenter(rect) {
  return {
    x: (Number(rect?.x) || 0) + (Number(rect?.width) || 0) / 2,
    y: (Number(rect?.y) || 0) + (Number(rect?.height) || 0) / 2,
  };
}

function getPhotoEditorRotatedMaskRectCorners(mask, width = 1, height = 1) {
  const rect = getPhotoEditorMaskNormalizedRect(mask);
  const center = getPhotoEditorMaskRectCenter(rect);
  const rotationRadians = getPhotoEditorMaskRotation(mask) * Math.PI / 180;
  const corners = [
    { x: rect.x, y: rect.y },
    { x: rect.x + rect.width, y: rect.y },
    { x: rect.x + rect.width, y: rect.y + rect.height },
    { x: rect.x, y: rect.y + rect.height },
  ];

  return rotationRadians === 0
    ? corners
    : corners.map((corner) =>
        rotateNormalizedPointAroundCenterForSize(corner, center, rotationRadians, width, height)
      );
}

function normalizePhotoEditorCropRotation(value) {
  const numericValue = Number(value);
  const roundedValue = Number.isFinite(numericValue)
    ? Math.round(numericValue / 90) * 90
    : 0;
  return ((roundedValue % 360) + 360) % 360;
}

function isPhotoEditorCropRotationSideways(crop = {}) {
  return normalizePhotoEditorCropRotation(crop?.rotation) % 180 !== 0;
}

function getPhotoEditorCropRotationDegrees(crop = {}) {
  return (
    normalizePhotoEditorCropRotation(crop?.rotation) +
    clampNumber(crop?.tilt, -45, 45, 0)
  );
}

function normalizePhotoEditorCropZoomValue(crop = {}) {
  const rawZoom = Number(crop?.zoom);

  if (!Number.isFinite(rawZoom)) {
    return 0;
  }

  const nextZoom = crop?.zoomMode === 'offset'
    ? rawZoom
    : rawZoom >= 100
      ? rawZoom - 100
      : rawZoom;

  return clampNumber(nextZoom, 0, 300, 0);
}

function getPhotoEditorCropZoomScale(crop = {}) {
  return 1 + normalizePhotoEditorCropZoomValue(crop) / 100;
}

function isPhotoEditorTransparentPaddingCrop(crop = {}) {
  return ['vrcGallery', 'vrcSticker'].includes(String(crop?.preset || ''));
}

function getPhotoEditorCropRenderGeometry(sourceRect, width, height, crop = {}) {
  const shouldUseTransparentPadding = isPhotoEditorTransparentPaddingCrop(crop);
  const isSideways = isPhotoEditorCropRotationSideways(crop);
  const baseWidth = isSideways ? sourceRect.height : sourceRect.width;
  const baseHeight = isSideways ? sourceRect.width : sourceRect.height;
  const outputScale = Math.min(
    width / Math.max(1, baseWidth),
    height / Math.max(1, baseHeight)
  );
  const drawWidth = Math.max(1, sourceRect.width * outputScale);
  const drawHeight = Math.max(1, sourceRect.height * outputScale);
  const rotationRadians = getPhotoEditorCropRotationDegrees(crop) * Math.PI / 180;
  const absCos = Math.abs(Math.cos(rotationRadians));
  const absSin = Math.abs(Math.sin(rotationRadians));
  const coverScale = shouldUseTransparentPadding
    ? getPhotoEditorCropZoomScale(crop)
    : Math.max(
        1,
        (width * absCos + height * absSin) / drawWidth,
        (width * absSin + height * absCos) / drawHeight
      );
  const scaledDrawWidth = drawWidth * coverScale;
  const scaledDrawHeight = drawHeight * coverScale;
  const rotatedWidth = scaledDrawWidth * absCos + scaledDrawHeight * absSin;
  const rotatedHeight = scaledDrawWidth * absSin + scaledDrawHeight * absCos;
  const overflowX = Math.max(0, (rotatedWidth - width) / 2);
  const overflowY = Math.max(0, (rotatedHeight - height) / 2);
  const translateX = shouldUseTransparentPadding
    ? overflowX * clampNumber(crop?.offsetX, -100, 100, 0) / 100
    : 0;
  const translateY = shouldUseTransparentPadding
    ? overflowY * clampNumber(crop?.offsetY, -100, 100, 0) / 100
    : 0;

  return {
    drawWidth,
    drawHeight,
    rotationRadians,
    coverScale,
    overflowX,
    overflowY,
    translateX,
    translateY,
    flipX: Boolean(crop?.flipX),
    flipY: Boolean(crop?.flipY),
  };
}

function mapPhotoEditorMaskPointToCanvas(point, mask, width, height, sourceRect, sourceImageSize, crop) {
  const allowOutside = canPhotoEditorMaskExtendOutside(mask);
  const bounds = getPhotoEditorMaskCoordinateBounds(allowOutside);
  const x = clampNumber(point?.x, bounds.min, bounds.max, 0);
  const y = clampNumber(point?.y, bounds.min, bounds.max, 0);
  const imageWidth = Number(sourceImageSize?.width) || 0;
  const imageHeight = Number(sourceImageSize?.height) || 0;
  const canUseSourceSpace =
    mask?.space === 'source' &&
    sourceRect &&
    sourceRect.width > 0 &&
    sourceRect.height > 0 &&
    imageWidth > 0 &&
    imageHeight > 0;

  if (!canUseSourceSpace) {
    return {
      x: x * width,
      y: y * height,
    };
  }

  const geometry = getPhotoEditorCropRenderGeometry(sourceRect, width, height, crop);
  const localX =
    ((x * imageWidth - sourceRect.x) / sourceRect.width - 0.5) *
    geometry.drawWidth *
    geometry.coverScale;
  const localY =
    ((y * imageHeight - sourceRect.y) / sourceRect.height - 0.5) *
    geometry.drawHeight *
    geometry.coverScale;
  const flippedX = geometry.flipX ? -localX : localX;
  const flippedY = geometry.flipY ? -localY : localY;
  const cos = Math.cos(geometry.rotationRadians);
  const sin = Math.sin(geometry.rotationRadians);

  return {
    x: width / 2 + geometry.translateX + flippedX * cos - flippedY * sin,
    y: height / 2 + geometry.translateY + flippedX * sin + flippedY * cos,
  };
}

function getPhotoEditorMaskCanvasPoints(mask, width, height, sourceRect, sourceImageSize, crop) {
  return (Array.isArray(mask?.points) ? mask.points : [])
    .filter((point) => Number.isFinite(point?.x) && Number.isFinite(point?.y))
    .map((point) =>
      mapPhotoEditorMaskPointToCanvas(
        point,
        mask,
        width,
        height,
        sourceRect,
        sourceImageSize,
        crop
      )
    );
}

function getRawCanvasRectFromMask(mask, width, height, sourceRect, sourceImageSize, crop) {
  if (mask?.shape === 'freehand') {
    const points = getPhotoEditorMaskCanvasPoints(mask, width, height, sourceRect, sourceImageSize, crop);

    if (points.length > 0) {
      const xs = points.map((point) => point.x);
      const ys = points.map((point) => point.y);
      const left = Math.min(...xs);
      const top = Math.min(...ys);
      const right = Math.max(...xs);
      const bottom = Math.max(...ys);

      return {
        x: left,
        y: top,
        width: right - left,
        height: bottom - top,
      };
    }
  }

  const corners = getPhotoEditorRotatedMaskRectCorners(mask, width, height).map((corner) =>
    mapPhotoEditorMaskPointToCanvas(
      corner,
      mask,
      width,
      height,
      sourceRect,
      sourceImageSize,
      crop
    )
  );
  const xs = corners.map((corner) => corner.x);
  const ys = corners.map((corner) => corner.y);
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  const right = Math.max(...xs);
  const bottom = Math.max(...ys);

  return {
    x: left,
    y: top,
    width: right - left,
    height: bottom - top,
  };
}

function getVisibleCanvasRectFromMask(mask, width, height, sourceRect, sourceImageSize, crop) {
  const maskRect = getRawCanvasRectFromMask(mask, width, height, sourceRect, sourceImageSize, crop);

  if (!maskRect) {
    return null;
  }

  const left = clampNumber(maskRect.x, 0, width, 0);
  const top = clampNumber(maskRect.y, 0, height, 0);
  const right = clampNumber(maskRect.x + maskRect.width, 0, width, 0);
  const bottom = clampNumber(maskRect.y + maskRect.height, 0, height, 0);

  if (right <= left || bottom <= top) {
    return null;
  }

  return {
    x: Math.round(left),
    y: Math.round(top),
    width: Math.max(1, Math.round(right - left)),
    height: Math.max(1, Math.round(bottom - top)),
  };
}

function addPhotoEditorMaskPath(ctx, mask, width, height, sourceRect, sourceImageSize, crop) {
  const maskRect = getRawCanvasRectFromMask(mask, width, height, sourceRect, sourceImageSize, crop);

  if (!maskRect || maskRect.width <= 0 || maskRect.height <= 0) {
    return null;
  }

  ctx.beginPath();

  if (
    mask?.space !== 'source' &&
    ['rect', 'ellipse'].includes(mask?.shape) &&
    getPhotoEditorMaskRotation(mask) !== 0
  ) {
    const rect = getPhotoEditorMaskNormalizedRect(mask);
    const center = getPhotoEditorMaskRectCenter(rect);
    const centerX = center.x * width;
    const centerY = center.y * height;
    const rectWidth = rect.width * width;
    const rectHeight = rect.height * height;

    ctx.save();
    ctx.translate(centerX, centerY);
    ctx.rotate(getPhotoEditorMaskRotation(mask) * Math.PI / 180);

    if (mask.shape === 'ellipse') {
      ctx.ellipse(0, 0, rectWidth / 2, rectHeight / 2, 0, 0, Math.PI * 2);
    } else {
      ctx.rect(-rectWidth / 2, -rectHeight / 2, rectWidth, rectHeight);
    }

    ctx.restore();
    return maskRect;
  }

  if (mask?.shape === 'ellipse') {
    ctx.ellipse(
      maskRect.x + maskRect.width / 2,
      maskRect.y + maskRect.height / 2,
      maskRect.width / 2,
      maskRect.height / 2,
      0,
      0,
      Math.PI * 2
    );
    return maskRect;
  }

  if (mask?.shape === 'freehand' && Array.isArray(mask.points) && mask.points.length > 1) {
    const points = getPhotoEditorMaskCanvasPoints(mask, width, height, sourceRect, sourceImageSize, crop);
    points.forEach((point, index) => {
      if (index === 0) {
        ctx.moveTo(point.x, point.y);
      } else {
        ctx.lineTo(point.x, point.y);
      }
    });
    ctx.closePath();
    return maskRect;
  }

  ctx.rect(maskRect.x, maskRect.y, maskRect.width, maskRect.height);
  return maskRect;
}

function withPhotoEditorMaskClip(ctx, mask, width, height, sourceRect, sourceImageSize, crop, draw) {
  const maskRect = getVisibleCanvasRectFromMask(mask, width, height, sourceRect, sourceImageSize, crop);

  if (!maskRect) {
    return;
  }

  ctx.save();
  const pathRect = addPhotoEditorMaskPath(ctx, mask, width, height, sourceRect, sourceImageSize, crop);

  if (!pathRect) {
    ctx.restore();
    return;
  }

  ctx.clip();
  draw(maskRect);
  ctx.restore();
}

function applyPhotoEditorFillMask(ctx, mask, width, height, sourceRect, sourceImageSize, crop) {
  const strength = clampNumber(mask?.strength, 0, 100, 100);

  if (strength <= 0) {
    return;
  }

  withPhotoEditorMaskClip(ctx, mask, width, height, sourceRect, sourceImageSize, crop, (maskRect) => {
    ctx.fillStyle = mask.color || '#111827';
    ctx.globalAlpha = strength / 100;
    ctx.fillRect(maskRect.x, maskRect.y, maskRect.width, maskRect.height);
    ctx.globalAlpha = 1;
  });
}

function applyPhotoEditorMosaicMask(ctx, mask, width, height, sourceRect, sourceImageSize, crop) {
  const strength = clampNumber(mask?.strength, 0, 100, 45);

  if (strength <= 0) {
    return;
  }

  const maskRect = getVisibleCanvasRectFromMask(mask, width, height, sourceRect, sourceImageSize, crop);

  if (!maskRect) {
    return;
  }

  const mosaicDivisor = Math.max(8, 30 - strength * 0.22);
  const cellSize = Math.max(
    3,
    Math.round(Math.min(maskRect.width, maskRect.height) / mosaicDivisor)
  );
  const tempCanvas = new OffscreenCanvas(
    Math.max(1, Math.ceil(maskRect.width / cellSize)),
    Math.max(1, Math.ceil(maskRect.height / cellSize))
  );
  const tempCtx = tempCanvas.getContext('2d');

  if (!tempCtx) {
    return;
  }

  tempCtx.imageSmoothingEnabled = true;
  tempCtx.drawImage(
    ctx.canvas,
    maskRect.x,
    maskRect.y,
    maskRect.width,
    maskRect.height,
    0,
    0,
    tempCanvas.width,
    tempCanvas.height
  );

  ctx.save();
  const pathRect = addPhotoEditorMaskPath(ctx, mask, width, height, sourceRect, sourceImageSize, crop);

  if (!pathRect) {
    ctx.restore();
    return;
  }

  ctx.clip();
  ctx.imageSmoothingEnabled = false;
  ctx.drawImage(tempCanvas, 0, 0, tempCanvas.width, tempCanvas.height, maskRect.x, maskRect.y, maskRect.width, maskRect.height);
  ctx.restore();
}

function applyPhotoEditorBlurMask(
  ctx,
  mask,
  width,
  height,
  sourceRect,
  sourceImageSize,
  crop,
  { isInteractivePreview = false } = {}
) {
  const strength = clampNumber(mask?.strength, 0, 100, 45);

  if (strength <= 0) {
    return;
  }

  const maskRect = getVisibleCanvasRectFromMask(mask, width, height, sourceRect, sourceImageSize, crop);

  if (!maskRect) {
    return;
  }

  const blurRadius = Math.max(
    1,
    Math.min(
      isInteractivePreview ? 28 : 46,
      Math.round(
        (Math.min(maskRect.width, maskRect.height) * (isInteractivePreview ? 0.075 : 0.1)) *
          (strength / 100)
      )
    )
  );
  const padding = Math.ceil(blurRadius * 2);
  const sourceX = Math.max(0, maskRect.x - padding);
  const sourceY = Math.max(0, maskRect.y - padding);
  const sourceRight = Math.min(ctx.canvas.width, maskRect.x + maskRect.width + padding);
  const sourceBottom = Math.min(ctx.canvas.height, maskRect.y + maskRect.height + padding);
  const sourceWidth = Math.max(1, sourceRight - sourceX);
  const sourceHeight = Math.max(1, sourceBottom - sourceY);
  const tempCanvas = new OffscreenCanvas(sourceWidth, sourceHeight);
  const tempCtx = tempCanvas.getContext('2d');

  if (!tempCtx) {
    return;
  }

  tempCtx.drawImage(ctx.canvas, sourceX, sourceY, sourceWidth, sourceHeight, 0, 0, sourceWidth, sourceHeight);
  ctx.save();
  const pathRect = addPhotoEditorMaskPath(ctx, mask, width, height, sourceRect, sourceImageSize, crop);

  if (!pathRect) {
    ctx.restore();
    ctx.filter = 'none';
    return;
  }

  ctx.clip();
  ctx.filter = `blur(${blurRadius}px)`;
  ctx.drawImage(tempCanvas, sourceX, sourceY, sourceWidth, sourceHeight);
  ctx.restore();
  ctx.filter = 'none';
}

function applyPhotoEditorMasksToCanvas(
  ctx,
  masks,
  width,
  height,
  sourceRect,
  sourceImageSize,
  crop,
  { isInteractivePreview = false } = {}
) {
  for (const mask of Array.isArray(masks) ? masks : []) {
    if (mask.type === 'blur') {
      applyPhotoEditorBlurMask(ctx, mask, width, height, sourceRect, sourceImageSize, crop, {
        isInteractivePreview,
      });
      continue;
    }

    if (mask.type === 'mosaic') {
      applyPhotoEditorMosaicMask(ctx, mask, width, height, sourceRect, sourceImageSize, crop);
      continue;
    }

    applyPhotoEditorFillMask(ctx, mask, width, height, sourceRect, sourceImageSize, crop);
  }
}

function normalizePhotoEditorTextColor(value, fallback = '#ffffff') {
  const color = String(value || '').trim();
  return /^#[0-9a-f]{6}$/i.test(color) ? color : fallback;
}

function normalizePhotoEditorTextState(textState = {}) {
  const strokeTypes = ['none', 'outline', 'shadow', 'glow'];
  const weights = ['400', '500', '700', '900'];
  const text = String(textState?.text || '').slice(0, 160);
  const fontFamily =
    typeof textState?.fontFamily === 'string' && textState.fontFamily.trim()
      ? textState.fontFamily
      : '"Segoe UI", "Yu Gothic UI", sans-serif';
  const strokeType = strokeTypes.includes(textState?.strokeType)
    ? textState.strokeType
    : 'outline';
  const weight = weights.includes(String(textState?.weight))
    ? String(textState.weight)
    : '700';

  return {
    enabled: Boolean(textState?.enabled || text.trim()),
    text,
    x: clampNumber(textState?.x, 0, 1, 0.5),
    y: clampNumber(textState?.y, 0, 1, 0.5),
    fontFamily,
    size: clampNumber(textState?.size, 12, 220, 64),
    color: normalizePhotoEditorTextColor(textState?.color, '#ffffff'),
    weight,
    strokeType,
    strokeWidth: clampNumber(textState?.strokeWidth, 0, 28, 4),
    strokeColor: normalizePhotoEditorTextColor(textState?.strokeColor, '#111827'),
    fillTransparent: Boolean(textState?.fillTransparent),
    letterSpacing: clampNumber(textState?.letterSpacing, -10, 40, 0),
  };
}

function getPhotoEditorTextLines(text) {
  return String(text || '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.trimEnd())
    .filter((line) => line.length > 0)
    .slice(0, 6);
}

function measurePhotoEditorTextLineWidth(ctx, line, letterSpacing = 0) {
  const characters = Array.from(String(line || ''));

  if (characters.length === 0) {
    return 0;
  }

  return characters.reduce((width, character, index) => {
    const spacing = index < characters.length - 1 ? letterSpacing : 0;
    return width + ctx.measureText(character).width + spacing;
  }, 0);
}

function drawPhotoEditorTextLine(ctx, line, x, y, textState) {
  const characters = Array.from(String(line || ''));
  const letterSpacing = clampNumber(textState.letterSpacing, -10, 40, 0);
  let currentX =
    x - measurePhotoEditorTextLineWidth(ctx, line, letterSpacing) / 2;

  for (const character of characters) {
    if (textState.strokeType !== 'none' && textState.strokeWidth > 0) {
      ctx.strokeText(character, currentX, y);
    }

    if (!textState.fillTransparent) {
      ctx.fillText(character, currentX, y);
    }

    currentX += ctx.measureText(character).width + letterSpacing;
  }
}

function applyPhotoEditorTextOverlayToCanvas(ctx, width, height, textOverlay) {
  const textState = normalizePhotoEditorTextState(textOverlay);
  const lines = textState.enabled ? getPhotoEditorTextLines(textState.text) : [];

  if (lines.length === 0) {
    return;
  }

  const x = clampNumber(textState.x, 0, 1, 0.5) * width;
  const y = clampNumber(textState.y, 0, 1, 0.5) * height;
  const size = clampNumber(textState.size, 12, 220, 64);
  const lineHeight = size * 1.18;
  const startY = y - (lineHeight * (lines.length - 1)) / 2;
  const strokeWidth = clampNumber(textState.strokeWidth, 0, 28, 4);

  ctx.save();
  ctx.font = `${textState.weight} ${size}px ${textState.fontFamily}`;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  ctx.fillStyle = textState.color;
  ctx.strokeStyle = textState.strokeColor;
  ctx.lineWidth =
    textState.strokeType === 'none' ? 0 : Math.max(1, strokeWidth);
  ctx.lineJoin = 'round';
  ctx.miterLimit = 2;

  if (textState.strokeType === 'shadow') {
    ctx.shadowColor = textState.strokeColor;
    ctx.shadowBlur = Math.max(2, strokeWidth * 1.4);
    ctx.shadowOffsetX = Math.max(1, strokeWidth * 0.45);
    ctx.shadowOffsetY = Math.max(2, strokeWidth * 0.8);
  } else if (textState.strokeType === 'glow') {
    ctx.shadowColor = textState.strokeColor;
    ctx.shadowBlur = Math.max(6, strokeWidth * 2.8);
    ctx.shadowOffsetX = 0;
    ctx.shadowOffsetY = 0;
  }

  for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
    drawPhotoEditorTextLine(
      ctx,
      lines[lineIndex],
      x,
      startY + lineHeight * lineIndex,
      textState
    );
  }

  ctx.restore();
}

function applyPhotoEditorEffectsToCanvas(ctx, payload) {
  const width = payload.width;
  const height = payload.height;
  const isInteractivePreview = Boolean(
    payload.isInteractivePreview && payload.responseType !== 'blob'
  );

  applyPhotoEditorAdjustmentsToCanvas(
    ctx,
    width,
    height,
    payload.values || {},
    payload.curve || {}
  );
  applyPhotoEditorBlurEffectToCanvas(ctx, width, height, payload.blur || {}, {
    isInteractivePreview,
  });
  applyPhotoEditorMasksToCanvas(
    ctx,
    payload.masks || [],
    width,
    height,
    payload.sourceRect,
    payload.sourceImageSize,
    payload.crop || {},
    { isInteractivePreview }
  );

  if (payload.includeDraft && payload.draftMask) {
    applyPhotoEditorMasksToCanvas(
      ctx,
      [payload.draftMask],
      width,
      height,
      payload.sourceRect,
      payload.sourceImageSize,
      payload.crop || {},
      { isInteractivePreview }
    );
  }

  applyPhotoEditorTextOverlayToCanvas(
    ctx,
    width,
    height,
    payload.textOverlay || {}
  );
}

async function renderPhotoEditorPayload(payload) {
  if (!payload?.sourceBitmap || !payload.width || !payload.height) {
    throw new Error('Invalid render payload');
  }

  const canvas = new OffscreenCanvas(payload.width, payload.height);
  const ctx = canvas.getContext('2d', { willReadFrequently: true });

  if (!ctx) {
    throw new Error('OffscreenCanvas 2D context is unavailable');
  }

  ctx.drawImage(payload.sourceBitmap, 0, 0, payload.width, payload.height);
  payload.sourceBitmap.close?.();
  applyPhotoEditorEffectsToCanvas(ctx, payload);

  if (payload.responseType === 'blob') {
    const blob = await canvas.convertToBlob({
      type: payload.mimeType || 'image/png',
      quality: payload.quality,
    });
    const buffer = await blob.arrayBuffer();

    return {
      buffer,
      mimeType: blob.type || payload.mimeType || 'image/png',
      width: payload.width,
      height: payload.height,
    };
  }

  return {
    bitmap: canvas.transferToImageBitmap(),
    width: payload.width,
    height: payload.height,
  };
}

self.addEventListener('message', async (event) => {
  const message = event.data || {};

  if (message.type !== 'render') {
    return;
  }

  try {
    const result = await renderPhotoEditorPayload(message.payload || {});

    if (result.bitmap) {
      self.postMessage(
        {
          type: 'render-result',
          requestId: message.requestId,
          ok: true,
          width: result.width,
          height: result.height,
          bitmap: result.bitmap,
        },
        [result.bitmap]
      );
      return;
    }

    self.postMessage(
      {
        type: 'render-result',
        requestId: message.requestId,
        ok: true,
        width: result.width,
        height: result.height,
        mimeType: result.mimeType,
        buffer: result.buffer,
      },
      [result.buffer]
    );
  } catch (error) {
    self.postMessage({
      type: 'render-result',
      requestId: message.requestId,
      ok: false,
      error: error?.message || String(error),
    });
  }
});
