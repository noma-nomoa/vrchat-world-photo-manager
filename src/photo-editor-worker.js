'use strict';

const PHOTO_EDITOR_RADIAL_BLUR_MIN_RADIUS = 0.08;
const PHOTO_EDITOR_RADIAL_BLUR_MAX_RADIUS = 1.12;
const PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER = 0.04;
const PHOTO_EDITOR_MASK_OUTSIDE_MARGIN = 1.25;
const PHOTO_EDITOR_CURVE_DEFAULT_POINTS = Object.freeze([0, 0.25, 0.5, 0.75, 1]);
const PHOTO_EDITOR_CURVE_MODES = Object.freeze(['rgb', 'hsv']);
const PHOTO_EDITOR_CURVE_CHANNELS = Object.freeze({
  rgb: ['master', 'r', 'g', 'b'],
  hsv: ['master', 'h', 's', 'v'],
});

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
  let seed = ((x + 1) * 374761393 + (y + 1) * 668265263) >>> 0;
  seed = (seed ^ (seed >>> 13)) >>> 0;
  seed = Math.imul(seed, 1274126177) >>> 0;
  return (seed / 4294967295) * 2 - 1;
}

function clonePhotoEditorCurvePoints(points) {
  if (Array.isArray(points) && points.length === 3) {
    const legacyPoints = points.map((point, index) =>
      clampNumber(point, 0, 1, index / 2)
    );

    return [
      legacyPoints[0],
      legacyPoints[0] * 0.5 + legacyPoints[1] * 0.5,
      legacyPoints[1],
      legacyPoints[1] * 0.5 + legacyPoints[2] * 0.5,
      legacyPoints[2],
    ];
  }

  return PHOTO_EDITOR_CURVE_DEFAULT_POINTS.map((defaultValue, index) =>
    clampNumber(points?.[index], 0, 1, defaultValue)
  );
}

function normalizePhotoEditorCurveState(curve = {}) {
  const mode = PHOTO_EDITOR_CURVE_MODES.includes(curve?.mode)
    ? curve.mode
    : 'rgb';
  const validChannels = PHOTO_EDITOR_CURVE_CHANNELS[mode];
  const channel = validChannels.includes(curve?.channel)
    ? curve.channel
    : 'master';
  const sourcePoints = curve?.points || {};

  return {
    mode,
    channel,
    points: {
      rgb: {
        master: clonePhotoEditorCurvePoints(sourcePoints.rgb?.master),
        r: clonePhotoEditorCurvePoints(sourcePoints.rgb?.r),
        g: clonePhotoEditorCurvePoints(sourcePoints.rgb?.g),
        b: clonePhotoEditorCurvePoints(sourcePoints.rgb?.b),
      },
      hsv: {
        master: clonePhotoEditorCurvePoints(sourcePoints.hsv?.master),
        h: clonePhotoEditorCurvePoints(sourcePoints.hsv?.h),
        s: clonePhotoEditorCurvePoints(sourcePoints.hsv?.s),
        v: clonePhotoEditorCurvePoints(sourcePoints.hsv?.v),
      },
    },
  };
}

function getPhotoEditorCurveValue(points, inputValue) {
  const input = clampNumber(inputValue, 0, 1, 0);
  const curvePoints = clonePhotoEditorCurvePoints(points);
  const maxPointIndex = curvePoints.length - 1;
  const scaledInput = input * maxPointIndex;
  const leftIndex = Math.min(maxPointIndex - 1, Math.floor(scaledInput));
  const ratio = scaledInput - leftIndex;

  return curvePoints[leftIndex] * (1 - ratio) + curvePoints[leftIndex + 1] * ratio;
}

function buildPhotoEditorCurveLookup(points, { outputScale = 255 } = {}) {
  const lookup = new Float32Array(256);

  for (let index = 0; index < lookup.length; index += 1) {
    lookup[index] = getPhotoEditorCurveValue(points, index / 255) * outputScale;
  }

  return lookup;
}

function getPhotoEditorCurveLookupIndex(value, scale = 255) {
  return Math.max(0, Math.min(255, Math.round(value * 255 / scale)));
}

function isPhotoEditorCurvePointsIdentity(points) {
  const curvePoints = clonePhotoEditorCurvePoints(points);
  return curvePoints.every(
    (point, index) =>
      Math.abs(point - PHOTO_EDITOR_CURVE_DEFAULT_POINTS[index]) < 0.001
  );
}

function createPhotoEditorCurveLookups(curveState) {
  const curve = normalizePhotoEditorCurveState(curveState);
  const rgbIdentity = {};
  const hsvIdentity = {};

  for (const channel of PHOTO_EDITOR_CURVE_CHANNELS.rgb) {
    rgbIdentity[channel] = isPhotoEditorCurvePointsIdentity(
      curve.points.rgb[channel]
    );
  }

  for (const channel of PHOTO_EDITOR_CURVE_CHANNELS.hsv) {
    hsvIdentity[channel] = isPhotoEditorCurvePointsIdentity(
      curve.points.hsv[channel]
    );
  }

  return {
    hasRgbCurve: Object.values(rgbIdentity).some((isIdentity) => !isIdentity),
    hasHsvCurve: Object.values(hsvIdentity).some((isIdentity) => !isIdentity),
    rgbIdentity,
    hsvIdentity,
    rgb: {
      master: buildPhotoEditorCurveLookup(curve.points.rgb.master),
      r: buildPhotoEditorCurveLookup(curve.points.rgb.r),
      g: buildPhotoEditorCurveLookup(curve.points.rgb.g),
      b: buildPhotoEditorCurveLookup(curve.points.rgb.b),
    },
    hsv: {
      master: buildPhotoEditorCurveLookup(curve.points.hsv.master, {
        outputScale: 1,
      }),
      h: buildPhotoEditorCurveLookup(curve.points.hsv.h, { outputScale: 1 }),
      s: buildPhotoEditorCurveLookup(curve.points.hsv.s, { outputScale: 1 }),
      v: buildPhotoEditorCurveLookup(curve.points.hsv.v, { outputScale: 1 }),
    },
  };
}

function convertRgbToHsv(red, green, blue) {
  const r = clampNumber(red, 0, 255, 0) / 255;
  const g = clampNumber(green, 0, 255, 0) / 255;
  const b = clampNumber(blue, 0, 255, 0) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const delta = max - min;
  let hue = 0;

  if (delta > 0) {
    if (max === r) {
      hue = ((g - b) / delta) % 6;
    } else if (max === g) {
      hue = (b - r) / delta + 2;
    } else {
      hue = (r - g) / delta + 4;
    }

    hue /= 6;

    if (hue < 0) {
      hue += 1;
    }
  }

  return {
    h: hue,
    s: max === 0 ? 0 : delta / max,
    v: max,
  };
}

function convertHsvToRgb(hue, saturation, value) {
  const h = ((clampNumber(hue, 0, 1, 0) * 6) % 6 + 6) % 6;
  const s = clampNumber(saturation, 0, 1, 0);
  const v = clampNumber(value, 0, 1, 0);
  const chroma = v * s;
  const x = chroma * (1 - Math.abs((h % 2) - 1));
  const match = v - chroma;
  let red = 0;
  let green = 0;
  let blue = 0;

  if (h < 1) {
    red = chroma;
    green = x;
  } else if (h < 2) {
    red = x;
    green = chroma;
  } else if (h < 3) {
    green = chroma;
    blue = x;
  } else if (h < 4) {
    green = x;
    blue = chroma;
  } else if (h < 5) {
    red = x;
    blue = chroma;
  } else {
    red = chroma;
    blue = x;
  }

  return {
    red: (red + match) * 255,
    green: (green + match) * 255,
    blue: (blue + match) * 255,
  };
}

function applyPhotoEditorCurveStateToColor(red, green, blue, curveLookups) {
  let nextRed = red;
  let nextGreen = green;
  let nextBlue = blue;

  if (curveLookups.hasRgbCurve) {
    if (!curveLookups.rgbIdentity.master) {
      nextRed = curveLookups.rgb.master[getPhotoEditorCurveLookupIndex(nextRed)];
      nextGreen = curveLookups.rgb.master[getPhotoEditorCurveLookupIndex(nextGreen)];
      nextBlue = curveLookups.rgb.master[getPhotoEditorCurveLookupIndex(nextBlue)];
    }

    if (!curveLookups.rgbIdentity.r) {
      nextRed = curveLookups.rgb.r[getPhotoEditorCurveLookupIndex(nextRed)];
    }

    if (!curveLookups.rgbIdentity.g) {
      nextGreen = curveLookups.rgb.g[getPhotoEditorCurveLookupIndex(nextGreen)];
    }

    if (!curveLookups.rgbIdentity.b) {
      nextBlue = curveLookups.rgb.b[getPhotoEditorCurveLookupIndex(nextBlue)];
    }
  }

  if (curveLookups.hasHsvCurve) {
    const hsv = convertRgbToHsv(nextRed, nextGreen, nextBlue);
    let hue = hsv.h;
    let saturation = hsv.s;
    let value = hsv.v;

    if (!curveLookups.hsvIdentity.master) {
      value = curveLookups.hsv.master[getPhotoEditorCurveLookupIndex(value, 1)];
    }

    if (!curveLookups.hsvIdentity.h) {
      hue = curveLookups.hsv.h[getPhotoEditorCurveLookupIndex(hue, 1)];
    }

    if (!curveLookups.hsvIdentity.s) {
      saturation = curveLookups.hsv.s[
        getPhotoEditorCurveLookupIndex(saturation, 1)
      ];
    }

    if (!curveLookups.hsvIdentity.v) {
      value = curveLookups.hsv.v[getPhotoEditorCurveLookupIndex(value, 1)];
    }

    const mappedRgb = convertHsvToRgb(hue, saturation, value);
    nextRed = mappedRgb.red;
    nextGreen = mappedRgb.green;
    nextBlue = mappedRgb.blue;
  }

  return {
    red: nextRed,
    green: nextGreen,
    blue: nextBlue,
  };
}

function getPhotoEditorAdjustmentRenderParams(values, curveState = null) {
  const brightness = clampNumber(values?.brightness, -100, 100, 0);
  const exposureFactor = Math.pow(
    2,
    clampNumber(values?.exposure, -100, 100, 0) / 85
  );
  const contrast = clampNumber(values?.contrast, -60, 60, 0) * 1.1;
  const contrastFactor =
    (259 * (contrast + 255)) / (255 * (259 - contrast));
  const highlights = clampNumber(values?.highlights, -60, 60, 0) * 0.95;
  const whites = clampNumber(values?.whites, -100, 100, 0) * 1.45;
  const shadows = clampNumber(values?.shadows, -100, 100, 0) * 1.08;
  const blacks = clampNumber(values?.blacks, -100, 100, 0) * 1.16;
  const gamma = clampNumber(values?.gamma, -100, 100, 0);
  const gammaExponent = Math.pow(2, -gamma / 140);
  const temperature = clampNumber(values?.temperature, -100, 100, 0) * 1.05;
  const tint = clampNumber(values?.tint, -100, 100, 0) * 0.86;
  const saturationFactor = Math.max(
    0,
    1 + clampNumber(values?.saturation, -100, 100, 0) / 100
  );
  const vibrance = clampNumber(values?.vibrance, -100, 100, 0) / 100;
  const fade = clampNumber(values?.fade, 0, 100, 0) / 100;
  const grain = clampNumber(values?.grain, 0, 100, 0) * 0.42;
  const curveLookups = createPhotoEditorCurveLookups(curveState);
  const hasCurve = curveLookups.hasRgbCurve || curveLookups.hasHsvCurve;

  return {
    brightness,
    exposureFactor,
    contrastFactor,
    highlights,
    whites,
    shadows,
    blacks,
    gamma,
    gammaExponent,
    temperature,
    tint,
    saturationFactor,
    vibrance,
    fade,
    grain,
    curveLookups,
    hasCurve,
    hasPixelWork:
      brightness !== 0 ||
      exposureFactor !== 1 ||
      contrast !== 0 ||
      highlights !== 0 ||
      whites !== 0 ||
      shadows !== 0 ||
      blacks !== 0 ||
      gamma !== 0 ||
      temperature !== 0 ||
      tint !== 0 ||
      saturationFactor !== 1 ||
      vibrance !== 0 ||
      fade !== 0 ||
      grain !== 0 ||
      hasCurve,
  };
}

function applyPhotoEditorPixelAdjustmentsToImageData(imageData, width, params) {
  const data = imageData.data;
  const {
    brightness,
    exposureFactor,
    contrastFactor,
    highlights,
    whites,
    shadows,
    blacks,
    gamma,
    gammaExponent,
    temperature,
    tint,
    saturationFactor,
    vibrance,
    fade,
    grain,
    curveLookups,
    hasCurve,
  } = params;
  let x = 0;
  let y = 0;

  for (let index = 0; index < data.length; index += 4) {
    let red = data[index];
    let green = data[index + 1];
    let blue = data[index + 2];
    const luminance = (0.2126 * red + 0.7152 * green + 0.0722 * blue) / 255;
    const highlightWeight = smoothstep(0.44, 0.9, luminance);
    const shadowWeight = 1 - smoothstep(0, 0.48, luminance);
    const whiteWeight = Math.pow(smoothstep(0.66, 0.98, luminance), 0.92);
    const blackWeight = Math.pow(1 - smoothstep(0, 0.34, luminance), 1.12);
    const highlightLift = highlights * highlightWeight * (1 - whiteWeight * 0.72);

    red += brightness + highlightLift + shadows * shadowWeight;
    green += brightness + highlightLift + shadows * shadowWeight;
    blue += brightness + highlightLift + shadows * shadowWeight;

    if (whites !== 0) {
      const whiteAmount = whites / 100;
      const whitePush = Math.min(1, Math.abs(whiteAmount) * whiteWeight);
      const whitePointLift = whites * whiteWeight * 0.22;

      red += whitePointLift;
      green += whitePointLift;
      blue += whitePointLift;

      if (whiteAmount > 0) {
        red += (255 - red) * whitePush * 0.78;
        green += (255 - green) * whitePush * 0.78;
        blue += (255 - blue) * whitePush * 0.78;
      } else {
        red -= red * whitePush * 0.72;
        green -= green * whitePush * 0.72;
        blue -= blue * whitePush * 0.72;
      }
    }

    red += blacks * blackWeight;
    green += blacks * blackWeight;
    blue += blacks * blackWeight;
    red *= exposureFactor;
    green *= exposureFactor;
    blue *= exposureFactor;

    red += temperature * 0.55;
    green += temperature * 0.08;
    blue -= temperature * 0.55;

    red += tint * 0.28;
    green -= tint * 0.46;
    blue += tint * 0.28;

    red = (red - 128) * contrastFactor + 128;
    green = (green - 128) * contrastFactor + 128;
    blue = (blue - 128) * contrastFactor + 128;

    if (gamma !== 0) {
      red = Math.pow(clampNumber(red, 0, 255, 0) / 255, gammaExponent) * 255;
      green = Math.pow(clampNumber(green, 0, 255, 0) / 255, gammaExponent) * 255;
      blue = Math.pow(clampNumber(blue, 0, 255, 0) / 255, gammaExponent) * 255;
    }

    if (fade > 0) {
      const fadeCompression = fade * 0.36;
      const fadeLift = 42 * fade;
      red = red * (1 - fadeCompression) + fadeLift;
      green = green * (1 - fadeCompression) + fadeLift;
      blue = blue * (1 - fadeCompression) + fadeLift;
    }

    const gray = 0.2126 * red + 0.7152 * green + 0.0722 * blue;
    red = gray + (red - gray) * saturationFactor;
    green = gray + (green - gray) * saturationFactor;
    blue = gray + (blue - gray) * saturationFactor;

    if (vibrance !== 0) {
      const vibranceGray = 0.2126 * red + 0.7152 * green + 0.0722 * blue;
      const maxChannel = Math.max(red, green, blue);
      const minChannel = Math.min(red, green, blue);
      const pixelSaturation =
        maxChannel <= 0 ? 0 : clampNumber((maxChannel - minChannel) / maxChannel, 0, 1, 0);
      const vibranceWeight = vibrance > 0
        ? Math.pow(1 - pixelSaturation, 1.25)
        : 0.55 + pixelSaturation * 0.45;
      const vibranceFactor = Math.max(0, 1 + vibrance * vibranceWeight * 0.9);

      red = vibranceGray + (red - vibranceGray) * vibranceFactor;
      green = vibranceGray + (green - vibranceGray) * vibranceFactor;
      blue = vibranceGray + (blue - vibranceGray) * vibranceFactor;
    }

    if (grain > 0) {
      const noise = getDeterministicNoise(x, y) * grain;
      red += noise;
      green += noise;
      blue += noise;
    }

    if (hasCurve) {
      const mappedColor = applyPhotoEditorCurveStateToColor(
        red,
        green,
        blue,
        curveLookups
      );
      red = mappedColor.red;
      green = mappedColor.green;
      blue = mappedColor.blue;
    }

    data[index] = clampColorChannel(red);
    data[index + 1] = clampColorChannel(green);
    data[index + 2] = clampColorChannel(blue);

    x += 1;
    if (x >= width) {
      x = 0;
      y += 1;
    }
  }
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
    ? 1
    : Math.max(
        1,
        (width * absCos + height * absSin) / drawWidth,
        (width * absSin + height * absCos) / drawHeight
      );

  return {
    drawWidth,
    drawHeight,
    rotationRadians,
    coverScale,
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
    x: width / 2 + flippedX * cos - flippedY * sin,
    y: height / 2 + flippedX * sin + flippedY * cos,
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
