const { nativeImage } = require('electron');

const WIDTH = 1920;
const HEIGHT = 1080;

function clampChannel(value) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

// A broad tonal ramp exercises shadow/highlight and contrast adjustments. The
// balanced warm/cool bands expose temperature changes, and a small checker
// pattern gives clarity/texture adjustments repeatable detail to soften.
// No local photos, fonts, random values, or external image libraries are used.
function createRetouchPresetFixturePng() {
  const bitmap = Buffer.alloc(WIDTH * HEIGHT * 4);
  const colorBands = [
    [18, 0, -18],
    [-18, 0, 18],
    [0, 0, 0],
    [10, -8, 10],
    [-10, 8, -10],
    [0, 0, 0],
  ];

  for (let y = 0; y < HEIGHT; y += 1) {
    const tint = colorBands[Math.floor((y * colorBands.length) / HEIGHT)];
    for (let x = 0; x < WIDTH; x += 1) {
      const detail = (Math.floor(x / 12) + Math.floor(y / 12)) % 2 ? 10 : -10;
      const luminance = 22 + (210 * x) / (WIDTH - 1) + detail;
      const offset = (y * WIDTH + x) * 4;
      // Electron's Windows bitmap format is BGRA; every pixel is opaque.
      bitmap[offset] = clampChannel(luminance + tint[2]);
      bitmap[offset + 1] = clampChannel(luminance + tint[1]);
      bitmap[offset + 2] = clampChannel(luminance + tint[0]);
      bitmap[offset + 3] = 255;
    }
  }

  return nativeImage.createFromBitmap(bitmap, { width: WIDTH, height: HEIGHT }).toPNG();
}

module.exports = { createRetouchPresetFixturePng };
