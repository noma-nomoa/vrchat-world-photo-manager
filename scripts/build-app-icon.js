'use strict';

// Builds the app icon from the logo mark, so the taskbar, window and
// installer icon stay the same picture as the logo shown inside the app.
//
//   node_modules/.bin/electron scripts/build-app-icon.js
//
// Reads  img/newlogo/logo.svg
// Writes img/logo.ico (16 to 256 px) and img/logo.png (100 px; several smoke
// tests use this file as a small sample photo, so its size stays 100 x 100).

const fs = require('node:fs');
const path = require('node:path');
const { app, BrowserWindow } = require('electron');

const ROOT = path.join(__dirname, '..');
const SOURCE_SVG_PATH = path.join(ROOT, 'img', 'newlogo', 'logo.svg');
const ICO_PATH = path.join(ROOT, 'img', 'logo.ico');
const PNG_PATH = path.join(ROOT, 'img', 'logo.png');
const ICO_SIZES = [16, 20, 24, 32, 40, 48, 64, 128, 256];
const PNG_SIZE = 100;
// Sizes from this one up are stored as PNG inside the .ico, smaller ones as
// plain bitmaps: the layout every Windows icon reader understands.
const ICO_PNG_ENTRY_MIN_SIZE = 256;

// The logo goes in as a data URL: a canvas that drew a file:// picture
// refuses to hand its pixels back.
const sourceDataUrl = `data:image/svg+xml;base64,${fs
  .readFileSync(SOURCE_SVG_PATH)
  .toString('base64')}`;

async function renderLogo(window, size) {
  const result = await window.webContents.executeJavaScript(`
    (async () => {
      const image = new Image();
      image.src = ${JSON.stringify(sourceDataUrl)};
      await image.decode();
      const canvas = document.createElement('canvas');
      canvas.width = ${size};
      canvas.height = ${size};
      const context = canvas.getContext('2d');
      context.imageSmoothingQuality = 'high';
      context.drawImage(image, 0, 0, ${size}, ${size});
      const pixels = context.getImageData(0, 0, ${size}, ${size}).data;
      let binary = '';
      for (let start = 0; start < pixels.length; start += 8192) {
        binary += String.fromCharCode(...pixels.subarray(start, start + 8192));
      }
      return {
        png: canvas.toDataURL('image/png').split(',')[1],
        rgba: btoa(binary),
      };
    })()
  `);

  return {
    size,
    png: Buffer.from(result.png, 'base64'),
    rgba: Buffer.from(result.rgba, 'base64'),
  };
}

// A 32-bit bitmap entry: header, pixels bottom row first as BGRA, then the
// one-bit mask rows (all zero, because the alpha channel carries the shape).
function toIcoBitmap({ size, rgba }) {
  const header = Buffer.alloc(40);
  header.writeUInt32LE(40, 0);
  header.writeInt32LE(size, 4);
  header.writeInt32LE(size * 2, 8);
  header.writeUInt16LE(1, 12);
  header.writeUInt16LE(32, 14);
  header.writeUInt32LE(size * size * 4, 20);

  const pixels = Buffer.alloc(size * size * 4);

  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) {
      const source = ((size - 1 - y) * size + x) * 4;
      const target = (y * size + x) * 4;

      pixels[target] = rgba[source + 2];
      pixels[target + 1] = rgba[source + 1];
      pixels[target + 2] = rgba[source];
      pixels[target + 3] = rgba[source + 3];
    }
  }

  const maskRowBytes = Math.ceil(size / 32) * 4;

  return Buffer.concat([header, pixels, Buffer.alloc(maskRowBytes * size)]);
}

function buildIco(renders) {
  const images = renders.map((render) =>
    render.size >= ICO_PNG_ENTRY_MIN_SIZE ? render.png : toIcoBitmap(render)
  );
  const directory = Buffer.alloc(6 + renders.length * 16);
  let offset = directory.length;

  directory.writeUInt16LE(0, 0);
  directory.writeUInt16LE(1, 2);
  directory.writeUInt16LE(renders.length, 4);

  renders.forEach((render, index) => {
    const entry = 6 + index * 16;

    directory.writeUInt8(render.size >= 256 ? 0 : render.size, entry);
    directory.writeUInt8(render.size >= 256 ? 0 : render.size, entry + 1);
    directory.writeUInt16LE(1, entry + 4);
    directory.writeUInt16LE(32, entry + 6);
    directory.writeUInt32LE(images[index].length, entry + 8);
    directory.writeUInt32LE(offset, entry + 12);
    offset += images[index].length;
  });

  return Buffer.concat([directory, ...images]);
}

app.whenReady().then(async () => {
  const window = new BrowserWindow({ show: false });

  try {
    await window.loadURL('about:blank');

    const renders = [];

    for (const size of ICO_SIZES) {
      renders.push(await renderLogo(window, size));
    }

    fs.writeFileSync(ICO_PATH, buildIco(renders));
    fs.writeFileSync(PNG_PATH, (await renderLogo(window, PNG_SIZE)).png);
    console.log(
      `Wrote img/logo.ico (${ICO_SIZES.join(', ')} px) and img/logo.png (${PNG_SIZE} px).`
    );
  } catch (error) {
    console.error(error);
    process.exitCode = 1;
  } finally {
    app.quit();
  }
});
