// Seed only a newly created review candidate with synthetic, disposable images.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { app, BrowserWindow, nativeImage } = require('electron');
const repoRoot = path.resolve(__dirname, '..');
const argument = process.argv.find(value => value.startsWith('--candidate='));
assert(argument, '--candidate is required');
const candidateRoot = path.resolve(argument.slice('--candidate='.length));
const relative = path.relative(path.join(repoRoot, 'out'), candidateRoot);
assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative)
  && /^review-candidate-[^\\/]+$/.test(relative), 'Candidate must be a fresh review-candidate directory under out');
assert(fs.existsSync(path.join(candidateRoot, 'candidate.json')), 'Candidate manifest is required');
const profile = path.join(candidateRoot, 'ReviewProfile');
const appData = path.join(profile, 'AppData', 'Roaming');
const localAppData = path.join(profile, 'AppData', 'Local');
const temp = path.join(profile, 'Temp');
const samples = path.join(candidateRoot, 'SamplePhotos');
const dataDir = path.join(appData, 'WorldShot Log', 'data');
const dbPath = path.join(dataDir, 'app.sqlite');
assert(!fs.existsSync(dbPath), 'Refusing to overwrite a review database');
for (const dir of [profile, appData, localAppData, temp, samples, dataDir]) fs.mkdirSync(dir, { recursive: true });
app.setPath('home', profile);
app.setPath('appData', appData);
app.setPath('userData', path.dirname(dataDir));
app.setPath('temp', temp);
process.env.TEMP = process.env.TMP = temp;
process.env.APPDATA = process.env.WORLDSHOT_APPDATA_ROOT = appData;
process.env.LOCALAPPDATA = process.env.WORLDSHOT_LOCALAPPDATA_ROOT = localAppData;
process.env.USERPROFILE = process.env.WORLDSHOT_HOME_ROOT = profile;
const definitions = [
  { name: '01-sunset-landscape', title: '夕景 / 色と明暗', width: 1600, height: 1000, colors: ['#f09b74', '#5b3c73', '#171e38'], month: 9, day: 10 },
  { name: '02-portrait-subject', title: '人物シルエット / マスク', width: 1000, height: 1500, colors: ['#9bc8ce', '#63888e', '#213b4e'], month: 9, day: 10, portrait: true },
  { name: '03-square-night', title: '正方形 / 夜景', width: 1200, height: 1200, colors: ['#172b50', '#406584', '#161d30'], month: 9, day: 9 },
  { name: '04-soft-landscape', title: '淡い風景 / トーン', width: 1400, height: 900, colors: ['#e3ded3', '#a9b9ba', '#7b8d93'], month: 9, day: 8 },
  { name: '05-print-layout', title: 'プリント / 文字と画像', width: 1500, height: 1000, colors: ['#eab99d', '#cf7471', '#574453'], month: 8, day: 25 },
  { name: '06-blue-landscape', title: '青い風景 / 月切り替え', width: 1200, height: 800, colors: ['#b8e2df', '#549db6', '#24415f'], month: 8, day: 25 },
];
function drawFixture(definition) {
  const { width: w, height: h, colors, title, portrait } = definition;
  const canvas = document.createElement('canvas'); canvas.width = w; canvas.height = h;
  const ctx = canvas.getContext('2d');
  const sky = ctx.createLinearGradient(0, 0, 0, h); colors.forEach((color, i) => sky.addColorStop(i / 2, color));
  ctx.fillStyle = sky; ctx.fillRect(0, 0, w, h);
  ctx.fillStyle = '#ffe9b5'; ctx.beginPath(); ctx.arc(w * .72, h * .24, Math.min(w, h) * .095, 0, Math.PI * 2); ctx.fill();
  for (let layer = 0; layer < 3; layer += 1) {
    ctx.fillStyle = colors[2]; ctx.globalAlpha = .2 + layer * .21;
    ctx.beginPath(); ctx.moveTo(0, h);
    for (let x = 0; x <= w; x += w / 50) ctx.lineTo(x, h * (.53 + layer * .14) + Math.sin(x / w * 9 + layer) * h * .075);
    ctx.lineTo(w, h); ctx.closePath(); ctx.fill();
  }
  ctx.globalAlpha = 1;
  if (portrait) {
    ctx.fillStyle = '#322c3c'; ctx.beginPath(); ctx.ellipse(w * .5, h * .26, w * .13, h * .105, 0, 0, Math.PI * 2); ctx.fill();
    ctx.fillStyle = '#d5a383'; ctx.beginPath(); ctx.ellipse(w * .5, h * .285, w * .1, h * .09, 0, 0, Math.PI * 2); ctx.fill();
    ctx.fillStyle = '#203047'; ctx.beginPath(); ctx.moveTo(w * .4, h * .39); ctx.lineTo(w * .6, h * .39);
    ctx.lineTo(w * .75, h * .7); ctx.lineTo(w * .63, h * .73); ctx.lineTo(w * .59, h * .6);
    ctx.lineTo(w * .63, h * .91); ctx.lineTo(w * .53, h * .91); ctx.lineTo(w * .5, h * .72);
    ctx.lineTo(w * .47, h * .91); ctx.lineTo(w * .37, h * .91); ctx.lineTo(w * .41, h * .6);
    ctx.lineTo(w * .37, h * .73); ctx.lineTo(w * .25, h * .7); ctx.closePath(); ctx.fill();
    ctx.fillStyle = '#eee2ce'; ctx.fillRect(w * .43, h * .405, w * .14, h * .14);
  }
  ctx.strokeStyle = '#ffffff3d'; ctx.lineWidth = Math.max(1, w / 1000);
  for (let x = 0; x < w; x += w / 10) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, h); ctx.stroke(); }
  for (let y = 0; y < h; y += h / 10) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(w, y); ctx.stroke(); }
  const swatches = ['#ffffff', '#c6c6c6', '#777777', '#252525', '#e14955', '#e6bb4b', '#4aaf87', '#408dc1', '#986baa'];
  swatches.forEach((color, i) => { ctx.fillStyle = color; ctx.fillRect(w * .04 + i * w * .042, h * .055, w * .037, h * .035); });
  ctx.fillStyle = '#111827bb'; ctx.fillRect(w * .025, h * .91, w * .95, h * .065);
  ctx.fillStyle = '#ffffff'; ctx.font = `600 ${Math.round(w * .023)}px "Yu Gothic", sans-serif`;
  ctx.fillText(title, w * .045, h * .95);
  ctx.font = `${Math.round(w * .012)}px sans-serif`; ctx.textAlign = 'right';
  ctx.fillText(`SYNTHETIC REVIEW SAMPLE  ${w} x ${h}`, w * .95, h * .953);
  return canvas.toDataURL('image/png');
}
async function main() {
  await app.whenReady();
  const window = new BrowserWindow({ show: false, webPreferences: { backgroundThrottling: false } });
  await window.loadURL('data:text/html;charset=utf-8,<meta charset="utf-8"><title>Review fixtures</title>');
  const thumbnailDir = path.join(profile, 'WorldShot Log', 'thumbnails');
  fs.mkdirSync(thumbnailDir, { recursive: true });
  const db = require('../src/db').initDatabase(dbPath);
  const photos = [];
  try {
    for (const [index, definition] of definitions.entries()) {
      const dataUrl = await window.webContents.executeJavaScript(`(${drawFixture.toString()})(${JSON.stringify(definition)})`);
      const image = nativeImage.createFromDataURL(dataUrl);
      const bytes = image.toPNG();
      const fileHash = createHash('sha256').update(bytes).digest('hex');
      const filePath = path.join(samples, `${definition.name}.png`);
      fs.writeFileSync(filePath, bytes);
      const thumbnailPath = path.join(thumbnailDir, `${fileHash}.jpg`);
      fs.writeFileSync(thumbnailPath, image.resize({ width: 320, quality: 'best' }).toJPEG(85));
      const date = `2026-${String(definition.month).padStart(2, '0')}-${String(definition.day).padStart(2, '0')}`;
      const timestamp = new Date(2026, definition.month - 1, definition.day, 12, index, 0).getTime();
      // Reimport uses the file timestamp when generated PNGs lack capture EXIF.
      fs.utimesSync(filePath, new Date(timestamp), new Date(timestamp));
      const row = db.insertOrUpdatePhoto({ filePath, fileName: path.basename(filePath), fileHash,
        takenAt: `${date.replaceAll('-', '/')} 12:0${index}:00`, takenAtTimestamp: timestamp,
        groupDate: date, year: 2026, month: definition.month, day: definition.day,
        worldId: null, worldName: definition.title, worldNameManual: definition.title, worldUrl: null,
        thumbnailPath, imageWidth: definition.width, imageHeight: definition.height,
        resolutionTier: 'other', orientationTier: definition.width === definition.height ? 'square'
          : definition.width > definition.height ? 'landscape' : 'portrait',
        printNoteText: index === 4 ? '確認用のプリント文字です' : '', memoText: '合成した確認用サンプルです。自由に編集・保存できます。',
        createdAt: '2026-09-10T00:00:00.000Z', updatedAt: '2026-09-10T00:00:00.000Z' });
      db.updateFavoriteStatus(row.id, index === 0 || index === 2);
      db.replacePhotoTags(row.id, [{ name: '確認用', colorHex: '#5A80A3' },
        { name: definition.portrait ? 'マスク' : '色調整', colorHex: '#AC795A' }]);
      photos.push({ id: row.id, name: definition.title, filePath, fileHash, width: definition.width, height: definition.height });
    }
    assert.equal(db.getApplicationDataSummary().photoCount, definitions.length);
  } finally { db.close(); window.destroy(); }
  fs.writeFileSync(path.join(candidateRoot, 'review-profile.json'), JSON.stringify({ ok: true, profile, appData, localAppData,
    temp, samples, dbPath, synthetic: true, photos }, null, 2));
  console.log(`Seeded ${photos.length} synthetic review photos: ${profile}`);
}
main().then(() => app.exit(0)).catch(error => { console.error(error.stack || error); app.exit(1); });
