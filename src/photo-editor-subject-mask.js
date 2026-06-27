(function () {
  'use strict';

  const SUBJECT_MASK_STATUS_VALUES = ['none', 'loading', 'ready', 'failed'];
  const SUBJECT_MASK_SOURCE_VALUES = [
    'none',
    'lightweight',
    'standard',
    'high-quality',
    'manual',
    'imported',
    'dummy',
  ];

  function clampNumber(value, min, max, fallback = 0) {
    const number = Number(value);

    if (!Number.isFinite(number)) {
      return fallback;
    }

    return Math.min(max, Math.max(min, number));
  }

  function clampColorChannel(value) {
    return Math.max(0, Math.min(255, Math.round(Number(value) || 0)));
  }

  function smoothstep(edge0, edge1, value) {
    if (edge0 === edge1) {
      return value < edge0 ? 0 : 1;
    }

    const t = clampNumber((value - edge0) / (edge1 - edge0), 0, 1, 0);
    return t * t * (3 - 2 * t);
  }

  function getDefaultSubjectMaskState(overrides = {}) {
    return {
      enabled: false,
      status: 'none',
      source: 'none',
      modelId: null,
      maskDataUrl: '',
      width: 0,
      height: 0,
      feather: 0,
      expand: 0,
      invert: false,
      opacity: 0.55,
      showOverlay: false,
      createdAt: null,
      updatedAt: null,
      errorMessage: '',
      ...overrides,
    };
  }

  function normalizeSubjectMaskState(subjectMask = {}) {
    const defaults = getDefaultSubjectMaskState();
    const maskDataUrl =
      typeof subjectMask?.maskDataUrl === 'string' &&
      subjectMask.maskDataUrl.startsWith('data:image/')
        ? subjectMask.maskDataUrl
        : '';
    const enabled = Boolean(subjectMask?.enabled && maskDataUrl);

    return {
      enabled,
      status: SUBJECT_MASK_STATUS_VALUES.includes(subjectMask?.status)
        ? subjectMask.status
        : enabled
          ? 'ready'
          : defaults.status,
      source: SUBJECT_MASK_SOURCE_VALUES.includes(subjectMask?.source)
        ? subjectMask.source
        : enabled
          ? 'manual'
          : defaults.source,
      modelId:
        typeof subjectMask?.modelId === 'string' && subjectMask.modelId.trim()
          ? subjectMask.modelId.trim()
          : null,
      maskDataUrl,
      width: Math.max(0, Math.round(Number(subjectMask?.width) || 0)),
      height: Math.max(0, Math.round(Number(subjectMask?.height) || 0)),
      feather: 0,
      expand: 0,
      invert: Boolean(subjectMask?.invert),
      opacity: clampNumber(subjectMask?.opacity, 0, 1, defaults.opacity),
      showOverlay: Boolean(subjectMask?.showOverlay),
      createdAt:
        typeof subjectMask?.createdAt === 'string'
          ? subjectMask.createdAt
          : null,
      updatedAt:
        typeof subjectMask?.updatedAt === 'string'
          ? subjectMask.updatedAt
          : null,
      errorMessage:
        typeof subjectMask?.errorMessage === 'string'
          ? subjectMask.errorMessage
          : '',
    };
  }

  function getHistogramPercentile(histogram, percentile) {
    const total = histogram.reduce((sum, count) => sum + count, 0);

    if (total <= 0) {
      return 0;
    }

    const target = total * clampNumber(percentile, 0, 1, 0);
    let cumulative = 0;

    for (let index = 0; index < histogram.length; index += 1) {
      cumulative += histogram[index] || 0;

      if (cumulative >= target) {
        return index;
      }
    }

    return histogram.length - 1;
  }

  function getOtsuThreshold(histogram, pixelCount) {
    if (!Array.isArray(histogram) || pixelCount <= 0) {
      return 128;
    }

    let totalIntensity = 0;

    for (let index = 0; index < histogram.length; index += 1) {
      totalIntensity += index * (histogram[index] || 0);
    }

    let backgroundWeight = 0;
    let backgroundIntensity = 0;
    let bestThreshold = 128;
    let bestVariance = -1;

    for (let threshold = 0; threshold < histogram.length; threshold += 1) {
      const count = histogram[threshold] || 0;
      backgroundWeight += count;

      if (backgroundWeight <= 0) {
        continue;
      }

      const foregroundWeight = pixelCount - backgroundWeight;

      if (foregroundWeight <= 0) {
        break;
      }

      backgroundIntensity += threshold * count;
      const backgroundMean = backgroundIntensity / backgroundWeight;
      const foregroundMean =
        (totalIntensity - backgroundIntensity) / foregroundWeight;
      const variance =
        backgroundWeight *
        foregroundWeight *
        (backgroundMean - foregroundMean) *
        (backgroundMean - foregroundMean);

      if (variance > bestVariance) {
        bestVariance = variance;
        bestThreshold = threshold;
      }
    }

    return bestThreshold;
  }

  function removeSmallComponents(alphaValues, width, height, threshold) {
    const pixelCount = width * height;

    if (!alphaValues || pixelCount <= 0) {
      return alphaValues;
    }

    const visited = new Uint8Array(pixelCount);
    const keep = new Uint8Array(pixelCount);
    const queue = new Int32Array(pixelCount);
    const component = new Int32Array(pixelCount);
    const minComponentSize = Math.max(28, Math.round(pixelCount * 0.0012));
    let largestComponentStart = -1;
    let largestComponentSize = 0;

    for (let start = 0; start < pixelCount; start += 1) {
      if (visited[start] || alphaValues[start] < threshold) {
        continue;
      }

      let queueStart = 0;
      let queueEnd = 0;
      let componentSize = 0;
      visited[start] = 1;
      queue[queueEnd] = start;
      queueEnd += 1;

      while (queueStart < queueEnd) {
        const current = queue[queueStart];
        queueStart += 1;
        component[componentSize] = current;
        componentSize += 1;

        const x = current % width;
        const y = Math.floor(current / width);
        const neighbors = [
          x > 0 ? current - 1 : -1,
          x < width - 1 ? current + 1 : -1,
          y > 0 ? current - width : -1,
          y < height - 1 ? current + width : -1,
        ];

        for (const neighbor of neighbors) {
          if (
            neighbor >= 0 &&
            !visited[neighbor] &&
            alphaValues[neighbor] >= threshold
          ) {
            visited[neighbor] = 1;
            queue[queueEnd] = neighbor;
            queueEnd += 1;
          }
        }
      }

      if (componentSize > largestComponentSize) {
        largestComponentSize = componentSize;
        largestComponentStart = start;
      }

      if (componentSize >= minComponentSize) {
        for (let index = 0; index < componentSize; index += 1) {
          keep[component[index]] = 1;
        }
      }
    }

    if (largestComponentStart >= 0 && largestComponentSize > 0) {
      visited.fill(0);
      let queueStart = 0;
      let queueEnd = 0;
      visited[largestComponentStart] = 1;
      queue[queueEnd] = largestComponentStart;
      queueEnd += 1;

      while (queueStart < queueEnd) {
        const current = queue[queueStart];
        queueStart += 1;
        keep[current] = 1;

        const x = current % width;
        const y = Math.floor(current / width);
        const neighbors = [
          x > 0 ? current - 1 : -1,
          x < width - 1 ? current + 1 : -1,
          y > 0 ? current - width : -1,
          y < height - 1 ? current + width : -1,
        ];

        for (const neighbor of neighbors) {
          if (
            neighbor >= 0 &&
            !visited[neighbor] &&
            alphaValues[neighbor] >= threshold
          ) {
            visited[neighbor] = 1;
            queue[queueEnd] = neighbor;
            queueEnd += 1;
          }
        }
      }
    }

    for (let index = 0; index < pixelCount; index += 1) {
      if (!keep[index] && alphaValues[index] >= threshold) {
        alphaValues[index] = Math.min(alphaValues[index], threshold - 1);
      }
    }

    return alphaValues;
  }

  function fillSmallHoles(alphaValues, width, height, threshold) {
    const pixelCount = width * height;

    if (!alphaValues || pixelCount <= 0) {
      return alphaValues;
    }

    const visited = new Uint8Array(pixelCount);
    const queue = new Int32Array(pixelCount);
    const component = new Int32Array(pixelCount);
    const minDimension = Math.max(1, Math.min(width, height));
    const maxHoleSize = Math.max(
      24,
      Math.round(pixelCount * 0.0009),
      Math.round(minDimension * 0.9)
    );

    for (let start = 0; start < pixelCount; start += 1) {
      if (visited[start] || alphaValues[start] >= threshold) {
        continue;
      }

      let queueStart = 0;
      let queueEnd = 0;
      let componentSize = 0;
      let touchesEdge = false;
      let neighborAlphaSum = 0;
      let neighborAlphaCount = 0;
      visited[start] = 1;
      queue[queueEnd] = start;
      queueEnd += 1;

      while (queueStart < queueEnd) {
        const current = queue[queueStart];
        queueStart += 1;
        component[componentSize] = current;
        componentSize += 1;

        const x = current % width;
        const y = Math.floor(current / width);

        if (x === 0 || x === width - 1 || y === 0 || y === height - 1) {
          touchesEdge = true;
        }

        const neighbors = [
          x > 0 ? current - 1 : -1,
          x < width - 1 ? current + 1 : -1,
          y > 0 ? current - width : -1,
          y < height - 1 ? current + width : -1,
        ];

        for (const neighbor of neighbors) {
          if (neighbor < 0) {
            continue;
          }

          if (alphaValues[neighbor] >= threshold) {
            neighborAlphaSum += alphaValues[neighbor];
            neighborAlphaCount += 1;
          } else if (!visited[neighbor]) {
            visited[neighbor] = 1;
            queue[queueEnd] = neighbor;
            queueEnd += 1;
          }
        }
      }

      if (
        !touchesEdge &&
        componentSize <= maxHoleSize &&
        neighborAlphaCount > 0
      ) {
        const fillValue = clampColorChannel(
          Math.max(threshold + 10, neighborAlphaSum / neighborAlphaCount)
        );

        for (let index = 0; index < componentSize; index += 1) {
          alphaValues[component[index]] = fillValue;
        }
      }
    }

    return alphaValues;
  }

  function cleanupSubjectMaskCanvas(maskCanvas) {
    if (!maskCanvas) {
      return maskCanvas;
    }

    const maskCtx = maskCanvas.getContext('2d', { willReadFrequently: true });

    if (!maskCtx) {
      return maskCanvas;
    }

    try {
      const width = maskCanvas.width;
      const height = maskCanvas.height;
      const pixelCount = width * height;
      const imageData = maskCtx.getImageData(0, 0, width, height);
      const data = imageData.data;
      const histogram = Array.from({ length: 256 }, () => 0);
      const alphaValues = new Uint8Array(pixelCount);

      for (let index = 0; index < pixelCount; index += 1) {
        const value = data[index * 4];
        alphaValues[index] = value;
        histogram[value] += 1;
      }

      const p04 = getHistogramPercentile(histogram, 0.04);
      const p35 = getHistogramPercentile(histogram, 0.35);
      const p58 = getHistogramPercentile(histogram, 0.58);
      const p82 = getHistogramPercentile(histogram, 0.82);
      const p96 = getHistogramPercentile(histogram, 0.96);

      if (p96 - p04 < 12) {
        return maskCanvas;
      }

      const otsuThreshold = getOtsuThreshold(histogram, pixelCount);
      const foregroundCount = alphaValues.reduce(
        (count, value) => count + (value >= otsuThreshold ? 1 : 0),
        0
      );
      const foregroundRatio = foregroundCount / Math.max(1, pixelCount);
      const adjustedThreshold =
        foregroundRatio > 0.68
          ? Math.max(otsuThreshold, p58)
          : clampNumber(otsuThreshold, p35, p82, otsuThreshold);
      const low = clampNumber(
        adjustedThreshold - 20,
        p04,
        245,
        adjustedThreshold
      );
      const high = clampNumber(
        adjustedThreshold + 34,
        low + 1,
        p96,
        low + 36
      );
      const componentThreshold = Math.max(96, Math.round(adjustedThreshold));

      for (let index = 0; index < pixelCount; index += 1) {
        const value = alphaValues[index];
        let nextValue = smoothstep(low, high, value) * 255;

        if (value < adjustedThreshold * 0.52) {
          nextValue = 0;
        } else if (value > high + 12) {
          nextValue = 255;
        }

        alphaValues[index] = clampColorChannel(value * 0.1 + nextValue * 0.9);
      }

      removeSmallComponents(alphaValues, width, height, componentThreshold);
      fillSmallHoles(alphaValues, width, height, componentThreshold);

      for (let index = 0; index < pixelCount; index += 1) {
        const value = alphaValues[index];
        const dataIndex = index * 4;
        data[dataIndex] = value;
        data[dataIndex + 1] = value;
        data[dataIndex + 2] = value;
        data[dataIndex + 3] = 255;
      }

      maskCtx.putImageData(imageData, 0, 0);
    } catch {
      return maskCanvas;
    }

    return maskCanvas;
  }

  globalThis.PhotoEditorSubjectMask = {
    getDefaultSubjectMaskState,
    normalizeSubjectMaskState,
    getHistogramPercentile,
    getOtsuThreshold,
    removeSmallComponents,
    fillSmallHoles,
    cleanupSubjectMaskCanvas,
  };
})();
