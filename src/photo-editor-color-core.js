(function (root, factory) {
  'use strict';

  const api = factory();
  if (typeof module === 'object' && module.exports) {
    module.exports = api;
  } else {
    root.PhotoEditorColorCore = api;
  }
})(globalThis, function () {
  'use strict';

  // Shared numeric color processing. Canvas, UI state, and Worker requests stay
  // with their callers. Pixel operations modify RGB in place and preserve alpha.
  const PHOTO_EDITOR_CURVE_DEFAULT_POINTS = Object.freeze([0, 0.25, 0.5, 0.75, 1]);

  const PHOTO_EDITOR_CURVE_MODES = Object.freeze(['rgb', 'hsv']);

  const PHOTO_EDITOR_CURVE_CHANNELS = Object.freeze({
    rgb: ['master', 'r', 'g', 'b'],
    hsv: ['master', 'h', 's', 'v'],
  });

  function clampNumber(value, min, max, fallback = 0) {
    const numericValue = Number(value);

    if (!Number.isFinite(numericValue)) {
      return fallback;
    }

    return Math.max(min, Math.min(max, numericValue));
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
        clampNumber(point, 0, 1, PHOTO_EDITOR_CURVE_DEFAULT_POINTS[index * 2])
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

  function getDefaultPhotoEditorCurveState() {
    return {
      mode: 'rgb',
      channel: 'master',
      points: {
        rgb: {
          master: [...PHOTO_EDITOR_CURVE_DEFAULT_POINTS],
          r: [...PHOTO_EDITOR_CURVE_DEFAULT_POINTS],
          g: [...PHOTO_EDITOR_CURVE_DEFAULT_POINTS],
          b: [...PHOTO_EDITOR_CURVE_DEFAULT_POINTS],
        },
        hsv: {
          master: [...PHOTO_EDITOR_CURVE_DEFAULT_POINTS],
          h: [...PHOTO_EDITOR_CURVE_DEFAULT_POINTS],
          s: [...PHOTO_EDITOR_CURVE_DEFAULT_POINTS],
          v: [...PHOTO_EDITOR_CURVE_DEFAULT_POINTS],
        },
      },
    };
  }

  function normalizePhotoEditorCurveState(curve = {}) {
    const defaults = getDefaultPhotoEditorCurveState();
    const mode = PHOTO_EDITOR_CURVE_MODES.includes(curve?.mode)
      ? curve.mode
      : defaults.mode;
    const defaultChannel = 'master';
    const validChannels = PHOTO_EDITOR_CURVE_CHANNELS[mode];
    const channel = validChannels.includes(curve?.channel)
      ? curve.channel
      : defaultChannel;
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

  function isPhotoEditorCurveStateIdentity(curveState) {
    const curve = normalizePhotoEditorCurveState(curveState);

    return PHOTO_EDITOR_CURVE_MODES.every((mode) =>
      PHOTO_EDITOR_CURVE_CHANNELS[mode].every((channel) =>
        isPhotoEditorCurvePointsIdentity(curve.points[mode][channel])
      )
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
      curve,
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

  function applyPhotoEditorCurveStateToColor(
    red,
    green,
    blue,
    curveLookups
  ) {
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

      const mappedHsv = { h: hue, s: saturation, v: value };
      const mappedRgb = convertHsvToRgb(mappedHsv.h, mappedHsv.s, mappedHsv.v);

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
    const brightness = clampNumber(values.brightness, -100, 100, 0);
    const exposureFactor = Math.pow(
      2,
      clampNumber(values.exposure, -100, 100, 0) / 85
    );
    const contrast = clampNumber(values.contrast, -60, 60, 0) * 1.1;
    const contrastFactor =
      (259 * (contrast + 255)) / (255 * (259 - contrast));
    const highlights = clampNumber(values.highlights, -60, 60, 0) * 0.95;
    const whites = clampNumber(values.whites, -100, 100, 0) * 1.45;
    const shadows = clampNumber(values.shadows, -100, 100, 0) * 1.08;
    const blacks = clampNumber(values.blacks, -100, 100, 0) * 1.16;
    const gamma = clampNumber(values.gamma, -100, 100, 0);
    const gammaExponent = Math.pow(2, -gamma / 140);
    const temperature = clampNumber(values.temperature, -100, 100, 0) * 1.05;
    const tint = clampNumber(values.tint, -100, 100, 0) * 0.86;
    const saturationFactor = Math.max(
      0,
      1 + clampNumber(values.saturation, -100, 100, 0) / 100
    );
    const vibrance = clampNumber(values.vibrance, -100, 100, 0) / 100;
    const fade = clampNumber(values.fade, 0, 100, 0) / 100;
    const grain = clampNumber(values.grain, 0, 100, 0) * 0.42;
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

  return Object.freeze({
    getDeterministicNoise,
    clonePhotoEditorCurvePoints,
    getDefaultPhotoEditorCurveState,
    normalizePhotoEditorCurveState,
    getPhotoEditorCurveValue,
    buildPhotoEditorCurveLookup,
    getPhotoEditorCurveLookupIndex,
    isPhotoEditorCurvePointsIdentity,
    isPhotoEditorCurveStateIdentity,
    createPhotoEditorCurveLookups,
    convertRgbToHsv,
    convertHsvToRgb,
    applyPhotoEditorCurveStateToColor,
    getPhotoEditorAdjustmentRenderParams,
    applyPhotoEditorPixelAdjustmentsToImageData,
  });
});
