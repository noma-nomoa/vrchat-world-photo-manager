'use strict';

const photos = [
  { caseId: 'newest', id: 20, takenAt: '2026/09/10 01:00:30', groupDate: '2026-09-10', fileName: 'latest.png', worldName: 'ＣＡＦＥ 森', rawWorldName: 'Original cafe', worldId: 'wrld_abc', isFavorite: true, imageWidth: 1200, imageHeight: 800, memoText: '旅のメモ', printNoteText: 'ｶﾌｪ Print', photoLabels: [{ normalizedName: 'friend', name: '友達', colorHex: '#112233' }, { normalizedName: 'travel', name: '旅行' }] },
  { caseId: 'portrait', id: 9, takenAt: '2026-09-09 23:00', groupDate: '2026-09-09', fileName: 'portrait.png', worldName: '森の世界', rawWorldName: 'CAFE', worldId: 'wrld_xyz', isFavorite: false, imageWidth: '600', imageHeight: '1200', memoText: 'Cafe meeting', printNoteText: '旅', photoLabels: [{ normalizedName: 'friend', name: 'different title', colorHex: '#FFFFFF' }] },
  { caseId: 'square', id: 8, takenAt: '2026/09/09', fileName: 'square.png', worldName: 'カフェ', isFavorite: true, imageWidth: 800, imageHeight: 800, memoText: '森', printNoteText: 'ＣＡＦＥ', photoLabels: [{ normalizedName: 'travel', name: '旅行' }] },
  { caseId: 'explicit-orientation', id: 7, takenAt: '2026/09/08 21:00:01', fileName: 'explicit.png', worldName: '別の場所', isFavorite: 'yes', orientationTier: 'portrait', imageWidth: 1200, imageHeight: 600, memoText: 'ｶﾌｪ', photoLabels: [{ normalizedName: 'friend' }, { normalizedName: 'friend', name: 'duplicate' }, null, {}] },
  { caseId: 'group-fallback', id: 6, takenAt: 'invalid', groupDate: '2026-09-07', takenAtTimestamp: 9999999999999, fileName: 'fallback.png', worldName: 'Café', rawWorldName: 'Cafe\u0301', imageWidth: 400, imageHeight: 800, printNoteText: '森の旅', photoLabels: [{ normalizedName: 'FRIEND', name: 'Uppercase label' }] },
  { caseId: 'missing-dimensions', id: 5, fileName: 'unknown.png', worldName: '한글', memoText: 'ＣＡＦＥ', printNoteText: '旅', photoLabels: 'invalid' },
  { caseId: 'null-dimensions', id: 4, fileName: 'null.png', imageWidth: null, imageHeight: null, worldName: '森', isFavorite: true, photoLabels: [] },
  { caseId: 'empty-dimensions', id: 3, fileName: 'empty.png', imageWidth: '', imageHeight: '', rawWorldName: 'カフェ', memoText: '旅', photoLabels: [{ normalizedName: 'travel', name: '旅行' }] },
  { caseId: 'zero-dimensions', id: 2, fileName: 'zero.png', imageWidth: 0, imageHeight: 0, worldId: 'WRLD_ABC', memoText: '森', photoLabels: [{ normalizedName: 'friend', name: '友達' }] },
  { caseId: 'invalid-dimensions', id: 1, fileName: 'invalid.png', imageWidth: 'NaN', imageHeight: 1, worldName: 'ＣＡＦＥ', photoLabels: [] },
  { caseId: 'tie-alpha', id: 30, takenAt: '2026-09-08 10:00:00', fileName: 'alpha.png', worldName: 'ＣＡＦＥ', imageWidth: 1200, imageHeight: 800, photoLabels: [{ normalizedName: 'friend' }] },
  { caseId: 'tie-zulu', id: 30, takenAt: '2026-09-08 10:00:00', fileName: 'zulu.png', worldName: '森', imageWidth: 1200, imageHeight: 800, photoLabels: [{ normalizedName: 'travel' }] },
  { caseId: 'tie-duplicate-first', id: 30, takenAt: '2026-09-08 10:00:00', fileName: 'same.png', worldName: 'ＣＡＦＥ', imageWidth: 1200, imageHeight: 800, isFavorite: true, photoLabels: [{ normalizedName: 'friend' }, { normalizedName: 'travel' }] },
  { caseId: 'tie-duplicate-second', id: 30, takenAt: '2026-09-08 10:00:00', fileName: 'same.png', worldName: '旅', imageWidth: 1200, imageHeight: 800, isFavorite: true, photoLabels: [{ normalizedName: 'friend' }, { normalizedName: 'travel' }] },
  { caseId: 'same-date-larger-id', id: 31, takenAt: '2026-09-08 10:00:00', fileName: 'alpha.png', worldName: 'Cafe', memoText: '旅', imageWidth: 100, imageHeight: 200, photoLabels: [{ normalizedName: 'friend' }] },
  { caseId: 'rollover-date', id: '32', takenAt: '2026/02/30 25:61:61', groupDate: '2099-12-31', fileName: 'rollover.png', rawWorldName: '森', imageWidth: -1, imageHeight: -2, photoLabels: [{ normalizedName: 'travel' }] },
  { caseId: 'trimmed-date', id: '33', takenAt: ' 2026/09/08\t10:00 ', fileName: 'trimmed.png', worldName: 'Mixed Case Cafe', orientationTier: 'unknown', imageWidth: 100, imageHeight: 200 },
  { caseId: 'raw-name-and-id', id: 34, takenAt: '2026-09-08 10:00:00', fileName: 'かな.png', worldName: '', rawWorldName: 'World 森', worldId: 'wrld_CAFE', imageWidth: 100, imageHeight: 100, memoText: 'Cafe\u0301 ｶﾌｪ 한글', printNoteText: '旅 森', photoLabels: [{ normalizedName: 'friend' }] },
];

function createQueries() {
  const queries = [];
  for (const favoriteOnly of [false, true]) {
    for (const orientation of ['all', 'landscape', 'portrait', 'square']) {
      for (const labelFilters of [[], ['friend'], ['friend', 'travel'], ['missing']]) {
        for (const labelMode of ['and', 'or']) {
          for (const searchScope of ['world', 'memo', 'printNote']) {
            for (const searchText of ['', 'ＣＡＦＥ', '森', 'ｶﾌｪ', '旅']) {
              for (const sortOrder of ['asc', 'desc']) {
                queries.push({ favoriteOnly, orientation, labelFilters: [...labelFilters], labelMode, searchScope, searchText, sortOrder });
              }
            }
          }
        }
      }
    }
  }
  return queries;
}

const dateInputs = [null, 0, '', '   ', 'bad', '2026-09-10', '2026/09/10 01:02', '2026/09/10 01:02:03',
  ' 2026-09-10\t01:02:03 ', '2026-9-10', '26-09-10', '2026-09-10T01:02:03Z', '2026/09/10 1:02:03',
  '2026-09-10 01:02:03.456', '2026-09-10 trailing', '2026/02/30 25:61:61', '2026-00-00',
  '0000-01-01', '0099-01-01', '0100-01-01', '2024-02-29', '2025-02-29', '2026-03-08 02:30:00',
  '2026-11-01 01:30:00'];
const searchInputs = [null, 0, '', '  ', '\tＣＡＦＥ 森\n', 'Cafe\u0301', 'ＣＡＦÉ', 'ｶﾌｪ', 'カフェ', '한글', ' İ I i ı ', 'Ａ\u3000Ｂ'];

module.exports = { photos, createQueries, dateInputs, searchInputs };
