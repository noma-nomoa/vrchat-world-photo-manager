const refreshTrackedFoldersButton = document.getElementById(
  'refresh-tracked-folders-btn'
);
const settingsButton = document.getElementById('settings-btn');
const regenerateThumbnailsButton = document.getElementById(
  'regenerate-thumbnails-btn'
);
const importStatus = document.getElementById('import-status');
const processingProgress = document.getElementById('processing-progress');
const processingProgressLabel = document.getElementById(
  'processing-progress-label'
);
const processingProgressValue = document.getElementById(
  'processing-progress-value'
);
const processingProgressFill = document.getElementById(
  'processing-progress-fill'
);
const processingProgressTrack = processingProgress?.querySelector(
  '.processing-progress-track'
);

// Main layout surfaces for sidebar, month header, and gallery content.
const sidebarTree = document.getElementById('sidebar-tree');
const currentMonthLabel = document.getElementById('current-month-label');
const currentMonthCount = document.getElementById('current-month-count');
const monthGalleryEmpty = document.getElementById('month-gallery-empty');
const monthGalleryList = document.getElementById('month-gallery-list');
const mainContent = document.querySelector('.main-content');
const dropOverlay = document.getElementById('drop-overlay');
const topStickyShell = document.querySelector('.top-sticky-shell');
const sidebar = document.querySelector('.sidebar');

// Header filter controls.
const favoriteFilterButton = document.getElementById('favorite-filter-btn');
const orientationFilterButton = document.getElementById(
  'orientation-filter-btn'
);
const orientationFilterDropdown = document.getElementById(
  'orientation-filter-dropdown'
);
const orientationFilterLabel = document.getElementById(
  'orientation-filter-label'
);
const orientationFilterMenu = document.getElementById('orientation-filter-menu');
const orientationFilterItems = Array.from(
  document.querySelectorAll('[data-orientation-filter]')
);
const photoLabelFilterDropdown = document.getElementById(
  'photo-label-filter-dropdown'
);
const photoLabelFilterButton = document.getElementById('photo-label-filter-btn');
const photoLabelFilterLabel = document.getElementById(
  'photo-label-filter-label'
);
const photoLabelFilterMenu = document.getElementById('photo-label-filter-menu');
const worldNameFilterDropdown = document.getElementById(
  'world-name-filter-dropdown'
);
const worldNameFilterButton = document.getElementById('world-name-filter-btn');
const worldNameFilterLabel = document.getElementById('world-name-filter-label');
const worldNameFilterMenu = document.getElementById('world-name-filter-menu');
const worldNameFilterInput = document.getElementById('world-name-filter-input');
const worldNameFilterClearButton = document.getElementById(
  'world-name-filter-clear-btn'
);

// Photo detail modal and its primary content areas.
const imageModal = document.getElementById('image-modal');
const imageModalBackdrop = document.getElementById('image-modal-backdrop');
const imageModalClose = document.getElementById('image-modal-close');
const imageModalContent = imageModal?.querySelector('.image-modal-content');
const imageModalBody = imageModal?.querySelector('.image-modal-body');
const imageModalImageWrap = imageModal?.querySelector('.image-modal-image-wrap');
const imageModalInfo = imageModal?.querySelector('.image-modal-info');
const modalImage = document.getElementById('modal-image');
const modalFileName = document.getElementById('modal-file-name');
const modalTakenAt = document.getElementById('modal-taken-at');
const modalResolutionTier = document.getElementById('modal-resolution-tier');
const modalWorldName = document.getElementById('modal-world-name');
const modalWorldId = document.getElementById('modal-world-id');
const modalWorldHero = document.querySelector('.modal-world-hero');
const modalWorldLabel = modalWorldHero?.querySelector('.modal-world-label');
const modalWorldDescription = document.getElementById('modal-world-description');
const modalWorldTags = document.getElementById('modal-world-tags');
const modalPhotoMemoInput = document.getElementById('modal-photo-memo-input');
const modalPhotoMemoSaveButton = document.getElementById('modal-photo-memo-save-btn');
const modalPhotoMemoStatus = document.getElementById('modal-photo-memo-status');

const modalWorldLink = document.getElementById('modal-world-link');
const modalOpenWorldButton = document.getElementById('modal-open-world-btn');
const modalOpenOriginalButton = document.getElementById(
  'modal-open-original-btn'
);
const modalOpenFolderButton = document.getElementById('modal-open-folder-btn');
const modalDeletePhotoButton = document.getElementById(
  'modal-delete-photo-btn'
);

const modalFavoriteButton = document.getElementById('modal-favorite-btn');
const modalFavoriteIcon = document.getElementById('modal-favorite-icon');
const worldNameEditorActions = document.querySelector('.world-name-editor-actions');
const modalDangerActions = document.querySelector('.modal-danger-actions');

// World name / URL edit modal.
const openWorldNameEditButton = document.getElementById(
  'open-world-name-edit-btn'
);
const worldNameEditModal = document.getElementById('world-name-edit-modal');
const worldNameEditBackdrop = document.getElementById(
  'world-name-edit-backdrop'
);
const worldNameEditClose = document.getElementById('world-name-edit-close');
const modalWorldNameInput = document.getElementById('modal-world-name-input');
const modalWorldUrlInput = document.getElementById('modal-world-url-input');
const saveWorldNameButton = document.getElementById('save-world-name-btn');
const clearWorldNameButton = document.getElementById('clear-world-name-btn');
const rereadWorldNameButton = document.getElementById(
  'reread-world-name-btn'
);
const worldNameSaveStatus = document.getElementById('world-name-save-status');

// Settings modal, tracked folders, and maintenance actions.
const settingsModal = document.getElementById('settings-modal');
const settingsModalBackdrop = document.getElementById(
  'settings-modal-backdrop'
);
const settingsModalClose = document.getElementById('settings-modal-close');
const settingsModalContent = settingsModal?.querySelector(
  '.settings-modal-content'
);
const settingsModalBody = settingsModal?.querySelector('.sub-modal-body');
const settingsSectionHeader = settingsModal?.querySelector(
  '.settings-section-header'
);
const addTrackedFolderButton = document.getElementById('add-tracked-folder-btn');
const trackedFolderList = document.getElementById('tracked-folder-list');
const deleteCurrentMonthRegistrationsButton = document.getElementById(
  'delete-current-month-registrations-btn'
);
const deleteAllRegistrationsButton = document.getElementById(
  'delete-all-registrations-btn'
);
const clearThumbnailCacheButton = document.getElementById(
  'clear-thumbnail-cache-btn'
);
const resetDatabaseButton = document.getElementById('reset-database-btn');
const toolbar = document.querySelector('.toolbar');
const toolbarRight = toolbar?.querySelector('.toolbar-right');
const pageHeaderActions = document.querySelector('.page-header-actions');
const fontOptionButtons = Array.from(
  document.querySelectorAll('[data-font-option]')
);

// Shared confirm modal and toast feedback.
const confirmModal = document.getElementById('confirm-modal');
const confirmModalBackdrop = document.getElementById(
  'confirm-modal-backdrop'
);
const confirmModalClose = document.getElementById('confirm-modal-close');
const confirmModalTitle = document.getElementById('confirm-modal-title');
const confirmModalMessage = document.getElementById('confirm-modal-message');
const confirmModalCancelButton = document.getElementById(
  'confirm-modal-cancel-btn'
);
const confirmModalConfirmButton = document.getElementById(
  'confirm-modal-confirm-btn'
);

const toast = document.getElementById('toast');

const themeToggleButton = document.getElementById('theme-toggle-btn');
const themeToggleIcon = document.getElementById('theme-toggle-icon');

const THEME_STORAGE_KEY = 'vrchat-world-photo-manager-theme';
const FONT_STORAGE_KEY = 'vrchat-world-photo-manager-font';

// Batch selection controls for the current month view.
const selectionModeButton = document.getElementById('selection-mode-btn');
const bulkFavoriteButton = document.getElementById('bulk-favorite-btn');
const bulkDeleteButton = document.getElementById('bulk-delete-btn');

// Sidebar/month/gallery state for the active selection.
let sidebarData = [];
let currentSelection = null;
let currentPhotos = [];
let allCurrentMonthPhotos = [];
let currentModalPhoto = null;
let imageModalAnimationTimer = null;
let imageModalSwitchTimer = null;
let modalShellRestoreTimer = null;
let modalWorldMetadataRequestId = 0;
let modalPhotoLabelsRequestId = 0;
let modalImageRecoveryRequestId = 0;
let trackedFolders = [];
let regenerateThumbnailMonthSelect = null;
let regenerateThumbnailMonthDropdown = null;
let regenerateThumbnailMonthButton = null;
let regenerateThumbnailMonthLabel = null;
let regenerateThumbnailMonthMenu = null;
let regenerateThumbnailMonthValue = '';
let isRegenerateThumbnailMonthMenuOpen = false;
let regenerateThumbnailMonthMenuCloseTimer = null;
let trackedFolderAccordionButton = null;
let trackedFolderAccordionPanel = null;
let modalResolutionHeroBadge = null;
let modalTakenAtHero = null;
let imageModalPrevButton = null;
let imageModalNextButton = null;
let modalPhotoLabelsBlock = null;
let modalPhotoLabelsList = null;
let openPhotoLabelEditorButton = null;
let photoLabelModal = null;
let photoLabelBackdrop = null;
let photoLabelClose = null;
let photoLabelSelectedList = null;
let photoLabelCatalogDropdown = null;
let photoLabelCatalogButton = null;
let photoLabelCatalogMenu = null;
let photoLabelNewForm = null;
let photoLabelNewNameInput = null;
let photoLabelNewColorInput = null;
let photoLabelNewColorPreview = null;
let photoLabelCustomColorButton = null;
let photoLabelPresetList = null;
let photoLabelAddNewButton = null;
let photoLabelSaveButton = null;
let photoLabelSaveStatus = null;
let currentModalPhotoLabels = [];
let draftModalPhotoLabels = [];
let photoLabelCatalog = [];
let activePhotoLabelCatalogSelection = '';
let isPhotoLabelCatalogMenuOpen = false;
let photoLabelCatalogMenuCloseTimer = null;
let isTrackedFolderAccordionOpen = false;
let isFavoriteFilterOnly = false;
let activeOrientationFilter = 'all';
let isOrientationFilterMenuOpen = false;
let orientationFilterMenuCloseTimer = null;
let activePhotoLabelFilters = [];
let photoLabelFilterMode = 'or';
let isPhotoLabelFilterMenuOpen = false;
let photoLabelFilterMenuCloseTimer = null;
let activeWorldNameFilter = '';
let isWorldNameFilterMenuOpen = false;
let worldNameFilterMenuCloseTimer = null;
let worldNameFilterInputTimer = null;
let isSelectionMode = false;
const selectedPhotoIds = new Set();
let isImporting = false;
let isWorldMetadataSyncing = false;
let worldMetadataSyncResetTimer = null;

// Expanded tree state and transient UI timers/overlays.
const expandedYears = new Set();

let toastTimer = null;
let confirmModalResolver = null;
let dropOverlayWatchTimer = null;
let monthSwitchAnimationTimer = null;
let scrollToTopAnimationFrame = null;
let monthSelectionRequestId = 0;
let activeMonthSwitchOverlay = null;
let renderedPhotoCount = 0;
let renderedMonthGalleryKey = '';
let isAppendingMonthGalleryBatch = false;
let monthGalleryLoadCheckScheduled = false;
let monthGalleryLoadCheckTimer = null;
let renderedGalleryGroupMap = new Map();
const subModalAnimationTimers = new WeakMap();

const GALLERY_CARD_MIN_WIDTH = 220;
const GALLERY_GRID_GAP = 16;
const GALLERY_GROUP_HORIZONTAL_PADDING = 36;
const GALLERY_INITIAL_PREFETCH_ROWS = 2;
const GALLERY_INCREMENT_ROWS = 3;
const GALLERY_CARD_EXTRA_HEIGHT = 110;
const GALLERY_LOAD_AHEAD_PX = 240;
const GALLERY_LOAD_CHECK_THROTTLE_MS = 96;
const IMAGE_MODAL_ANIMATION_MS = 520;
const SUB_MODAL_ANIMATION_MS = 520;
const SCROLL_TO_TOP_MIN_DURATION_MS = 420;
const SCROLL_TO_TOP_MAX_DURATION_MS = 1400;
const ORIENTATION_FILTER_ORDER = ['all', 'landscape', 'portrait', 'square'];
const ORIENTATION_FILTER_META = {
  all: {
    buttonLabel: '向き: すべて',
    shortLabel: 'すべて',
  },
  landscape: {
    buttonLabel: '向き: 横長',
    shortLabel: '横長',
  },
  portrait: {
    buttonLabel: '向き: 縦長',
    shortLabel: '縦長',
  },
  square: {
    buttonLabel: '向き: 正方形',
    shortLabel: '正方形',
  },
};
const PHOTO_LABEL_PRESET_COLORS = [
  '#6D5EF6',
  '#4F8CFF',
  '#14B8A6',
  '#22C55E',
  '#F59E0B',
  '#F97316',
  '#EC4899',
  '#8B5CF6',
  '#EF4444',
  '#06B6D4',
];

function pad2(value) {
  return String(value).padStart(2, '0');
}

function setText(el, value, fallback = '未取得') {
  if (!el) {
    return;
  }

  el.textContent = value || fallback;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function splitTakenAtForCard(value) {
  if (typeof value !== 'string' || value.trim().length === 0) {
    return {
      dateText: '日時不明',
      timeText: '',
    };
  }

  const [dateText, timeText] = value.trim().split(/\s+/, 2);

  return {
    dateText: dateText || '日時不明',
    timeText: timeText || '',
  };
}

function updateThemeToggleIcon(themeName) {
  if (!themeToggleIcon) {
    return;
  }

  themeToggleIcon.textContent =
    themeName === 'light' ? 'light_mode' : 'mode_night';
}

function applyTheme(themeName) {
  const nextTheme = themeName === 'light' ? 'light' : 'dark';
  document.body.setAttribute('data-theme', nextTheme);
  localStorage.setItem(THEME_STORAGE_KEY, nextTheme);
  updateThemeToggleIcon(nextTheme);
}

function toggleTheme() {
  const currentTheme = document.body.getAttribute('data-theme') || 'dark';
  const nextTheme = currentTheme === 'dark' ? 'light' : 'dark';
  applyTheme(nextTheme);
}

function initializeTheme() {
  const savedTheme = localStorage.getItem(THEME_STORAGE_KEY);

  if (savedTheme === 'light' || savedTheme === 'dark') {
    applyTheme(savedTheme);
    return;
  }

  applyTheme('dark');
}

function syncFontOptionButtons(fontName) {
  fontOptionButtons.forEach((button) => {
    const isActive = button.dataset.fontOption === fontName;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  });
}

function applyFontPreference(fontName) {
  const allowedFonts = new Set([
    'standard',
    'zenmaru',
    'mplus',
    'kiwimaru',
    'sawarabimincho',
  ]);
  const nextFont =
    typeof fontName === 'string' && allowedFonts.has(fontName)
      ? fontName
      : 'standard';

  document.body.setAttribute('data-font', nextFont);
  localStorage.setItem(FONT_STORAGE_KEY, nextFont);
  syncFontOptionButtons(nextFont);
}

function initializeFontPreference() {
  const savedFont = localStorage.getItem(FONT_STORAGE_KEY);

  if (
    savedFont === 'standard' ||
    savedFont === 'zenmaru' ||
    savedFont === 'mplus' ||
    savedFont === 'kiwimaru' ||
    savedFont === 'sawarabimincho'
  ) {
    applyFontPreference(savedFont);
    return;
  }

  applyFontPreference('standard');
}

function showToast(message) {
  if (!toast) {
    return;
  }

  toast.textContent = message;
  toast.classList.remove('hidden');

  if (toastTimer) {
    clearTimeout(toastTimer);
  }

  toastTimer = setTimeout(() => {
    toast.classList.add('hidden');
    toast.textContent = '';
  }, 2200);
}

function updateProcessingProgress(payload = {}) {
  if (!processingProgress) {
    return;
  }

  const total = Number(payload.total) || 0;
  const current = Math.max(0, Number(payload.current) || 0);
  const hasDeterminateProgress = total > 0;
  const clampedCurrent = hasDeterminateProgress ? Math.min(current, total) : 0;
  const percent = hasDeterminateProgress
    ? Math.max(0, Math.min(100, Math.round((clampedCurrent / total) * 100)))
    : 0;

  processingProgress.hidden = false;
  processingProgress.classList.toggle('is-indeterminate', !hasDeterminateProgress);

  if (processingProgressLabel) {
    processingProgressLabel.textContent = payload.message || '処理中...';
  }

  if (processingProgressValue) {
    processingProgressValue.textContent = hasDeterminateProgress
      ? `${clampedCurrent} / ${total}`
      : '...';
  }

  if (processingProgressFill) {
    processingProgressFill.style.width = hasDeterminateProgress
      ? percent > 0
        ? `calc(${percent}% + 2px)`
        : '0%'
      : '';
    processingProgressFill.style.transform = hasDeterminateProgress
      ? 'translateX(0)'
      : '';
    processingProgressFill.style.animation = hasDeterminateProgress
      ? 'none'
      : '';
  }

  if (processingProgressTrack) {
    processingProgressTrack.setAttribute('aria-valuenow', String(percent));
  }
}

function resetProcessingProgress() {
  if (!processingProgress) {
    return;
  }

  processingProgress.hidden = true;
  processingProgress.classList.remove('is-indeterminate');

  if (processingProgressLabel) {
    processingProgressLabel.textContent = '処理準備中...';
  }

  if (processingProgressValue) {
    processingProgressValue.textContent = '...';
  }

  if (processingProgressFill) {
    processingProgressFill.style.width = '0%';
    processingProgressFill.style.transform = 'translateX(0)';
    processingProgressFill.style.animation = 'none';
  }

  if (processingProgressTrack) {
    processingProgressTrack.setAttribute('aria-valuenow', '0');
  }
}

function syncWorldMetadataSyncUi() {
  if (!rereadWorldNameButton) {
    return;
  }

  rereadWorldNameButton.disabled = isWorldMetadataSyncing || !currentModalPhoto;
  rereadWorldNameButton.setAttribute(
    'title',
    isWorldMetadataSyncing
      ? '自動同期中は再読み込みできません'
      : 'World情報を再読み込み'
  );
}

function handleWorldMetadataSyncProgress(payload = {}) {
  if (worldMetadataSyncResetTimer) {
    clearTimeout(worldMetadataSyncResetTimer);
    worldMetadataSyncResetTimer = null;
  }

  const isCompletePhase = payload.phase === 'complete';
  isWorldMetadataSyncing = !isCompletePhase;
  syncWorldMetadataSyncUi();

  if (isImporting) {
    return;
  }

  updateProcessingProgress(payload);

  if (!isCompletePhase) {
    return;
  }

  worldMetadataSyncResetTimer = setTimeout(() => {
    if (!isImporting && !isWorldMetadataSyncing) {
      resetProcessingProgress();
    }
    worldMetadataSyncResetTimer = null;
  }, 420);
}

function applyWorldMetadataUpdated(payload = {}) {
  const updatedPhotos = Array.isArray(payload.photos) ? payload.photos : [];

  if (updatedPhotos.length === 0) {
    return;
  }
  syncBatchPhotoUpdates(updatedPhotos);
}

function buildImportStatusMessage(result, modeLabel) {
  if (!result || result.canceled) {
    return `${modeLabel}はキャンセルされました`;
  }

  if (result.emptyFolder) {
    return 'フォルダ内に対応画像がありませんでした';
  }

  if (result.emptyDrop) {
    return 'ドロップされた項目に対応画像がありませんでした';
  }

  return [
    `${modeLabel}: ${result.importedCount}件反映`,
    `新要E${result.newCount}件`,
    `更新 ${result.updatedCount}件`,
    result.failedCount > 0 ? `失敗 ${result.failedCount}件` : null,
  ]
    .filter(Boolean)
    .join(' / ');
}

function buildRegenerateThumbnailsMessage(result) {
  if (!result?.ok) {
    return `サムネイル再生成に失敗しました: ${
      result?.message || '不明なエラー'
    }`;
  }

  return [
    `サムネイル再生成 ${result.regeneratedCount}件更新`,
    `スキップ ${result.skippedCount}件`,
    result.failedCount > 0 ? `失敗 ${result.failedCount}件` : null,
  ]
    .filter(Boolean)
    .join(' / ');
}

function buildScopedRegenerateThumbnailsMessage(result) {
  if (!result?.ok) {
    return `サムネイル再生成に失敗しました: ${
      result?.message || '不明なエラー'
    }`;
  }

  const targetMonthLabel =
    Number.isInteger(result?.targetMonth?.year) &&
    Number.isInteger(result?.targetMonth?.month)
      ? `${result.targetMonth.year}年${result.targetMonth.month}月`
      : '全期間';

  return [
    `${targetMonthLabel}: サムネイル再生成 ${result.regeneratedCount}件`,
    `スキップ ${result.skippedCount}件`,
    result.failedCount > 0 ? `失敗 ${result.failedCount}件` : null,
  ]
    .filter(Boolean)
    .join(' / ');
}

function getSidebarMonthOptions() {
  return sidebarData.flatMap((yearEntry) =>
    (Array.isArray(yearEntry.months) ? yearEntry.months : []).map((monthEntry) => ({
      year: yearEntry.year,
      month: monthEntry.month,
      count: monthEntry.count,
    }))
  );
}

function renderRegenerateThumbnailMonthOptions() {
  if (!regenerateThumbnailMonthSelect) {
    return;
  }

  const monthOptions = getSidebarMonthOptions();
  const preferredValue =
    currentSelection &&
    monthOptions.some(
      (option) =>
        option.year === currentSelection.year &&
        option.month === currentSelection.month
    )
      ? `${currentSelection.year}-${pad2(currentSelection.month)}`
      : regenerateThumbnailMonthSelect.value;

  regenerateThumbnailMonthSelect.innerHTML = '';

  if (monthOptions.length === 0) {
    const emptyOption = document.createElement('option');
    emptyOption.value = '';
    emptyOption.textContent = '対象月がありません';
    regenerateThumbnailMonthSelect.appendChild(emptyOption);
    regenerateThumbnailMonthSelect.disabled = true;
    syncRegenerateThumbnailMonthDropdownFromSelect();
    return;
  }

  monthOptions.forEach((option) => {
    const selectOption = document.createElement('option');
    selectOption.value = `${option.year}-${pad2(option.month)}`;
    selectOption.textContent = `${option.year}年${option.month}月 (${option.count}枚)`;
    regenerateThumbnailMonthSelect.appendChild(selectOption);
  });

  regenerateThumbnailMonthSelect.disabled = false;

  if (
    typeof preferredValue === 'string' &&
    preferredValue.length > 0 &&
    monthOptions.some(
      (option) => `${option.year}-${pad2(option.month)}` === preferredValue
    )
  ) {
    regenerateThumbnailMonthSelect.value = preferredValue;
  }

  syncRegenerateThumbnailMonthDropdownFromSelect();
}

function setRegenerateThumbnailMonthMenuOpen(isOpen) {
  const hasOptions = Boolean(
    regenerateThumbnailMonthSelect && regenerateThumbnailMonthSelect.options.length > 0
  );
  const nextOpen = Boolean(isOpen) && !isImporting && hasOptions;
  isRegenerateThumbnailMonthMenuOpen = nextOpen;
  setAnimatedDropdownOpenState({
    dropdown: regenerateThumbnailMonthDropdown,
    button: regenerateThumbnailMonthButton,
    menu: regenerateThumbnailMonthMenu,
    isOpen: nextOpen,
    closeTimerRef: {
      get current() {
        return regenerateThumbnailMonthMenuCloseTimer;
      },
      set current(value) {
        regenerateThumbnailMonthMenuCloseTimer = value;
      },
    },
  });
}

function closeRegenerateThumbnailMonthMenu() {
  setRegenerateThumbnailMonthMenuOpen(false);
}

function syncRegenerateThumbnailMonthDropdownFromSelect() {
  if (
    !regenerateThumbnailMonthSelect ||
    !regenerateThumbnailMonthDropdown ||
    !regenerateThumbnailMonthButton ||
    !regenerateThumbnailMonthLabel ||
    !regenerateThumbnailMonthMenu
  ) {
    return;
  }

  const options = Array.from(regenerateThumbnailMonthSelect.options);
  regenerateThumbnailMonthValue = regenerateThumbnailMonthSelect.value || '';
  regenerateThumbnailMonthMenu.innerHTML = '';

  if (options.length === 0) {
    regenerateThumbnailMonthLabel.textContent = '対象月がありません';
    regenerateThumbnailMonthButton.disabled = true;
    closeRegenerateThumbnailMonthMenu();
    return;
  }

  const selectedOption =
    options.find((option) => option.value === regenerateThumbnailMonthValue) ||
    options[0];

  regenerateThumbnailMonthValue = selectedOption?.value || '';
  regenerateThumbnailMonthSelect.value = regenerateThumbnailMonthValue;
  regenerateThumbnailMonthLabel.textContent =
    selectedOption?.textContent || '再生成する月を選択';
  regenerateThumbnailMonthButton.disabled = false;

  options.forEach((option) => {
    const isActive = option.value === regenerateThumbnailMonthValue;
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'header-dropdown-item header-dropdown-item-with-meta';
    item.classList.toggle('is-active', isActive);
    item.dataset.regenerateThumbnailMonth = option.value;
    item.setAttribute('role', 'menuitemradio');
    item.setAttribute('aria-checked', isActive ? 'true' : 'false');

    const matched = option.textContent?.match(/^(.+?)\s*\((\d+)枚\)$/);
    const labelText = matched?.[1] || option.textContent || '';
    const countText = matched?.[2] || '';

    const itemLabel = document.createElement('span');
    itemLabel.className = 'header-dropdown-item-label';
    itemLabel.textContent = labelText;
    item.appendChild(itemLabel);

    const itemSide = document.createElement('span');
    itemSide.className = 'header-dropdown-item-side';

    if (countText) {
      const count = document.createElement('span');
      count.className = 'header-dropdown-meta-badge';
      count.textContent = countText;
      itemSide.appendChild(count);
    }

    const check = document.createElement('span');
    check.className = 'material-symbols-outlined header-dropdown-check';
    check.textContent = 'check';
    itemSide.appendChild(check);

    item.appendChild(itemSide);
    regenerateThumbnailMonthMenu.appendChild(item);
  });
}

function buildTrackedFoldersRefreshMessage(result) {
  if (!result || result.canceled) {
    return '更新はキャンセルされました';
  }

  if (result.ok === false) {
    return `更新に失敗しました: ${result.message || '不明なエラー'}`;
  }

  if (result.noTrackedFolders) {
    return '更新対象のフォルダがまだ登録されていません';
  }

  if (result.emptyRefresh) {
    return [
      `更新確認: 新規0件`,
      `追跡 ${result.trackedFolderCount || 0}件`,
      result.skippedKnownCount > 0
        ? `既知 ${result.skippedKnownCount}件`
        : null,
      result.missingFolderPaths?.length > 0
        ? `未検出 ${result.missingFolderPaths.length}件`
        : null,
    ]
      .filter(Boolean)
      .join(' / ');
  }

  return [
    `更新: ${result.importedCount || 0}件取込`,
    `新要E${result.newCount || 0}件`,
    result.updatedCount > 0 ? `再取込 ${result.updatedCount}件` : null,
    result.skippedKnownCount > 0 ? `既知 ${result.skippedKnownCount}件` : null,
    result.missingFolderPaths?.length > 0
      ? `未検出 ${result.missingFolderPaths.length}件`
      : null,
    result.failedCount > 0 ? `失敗 ${result.failedCount}件` : null,
  ]
    .filter(Boolean)
    .join(' / ');
}

function syncFavoriteButtonState(button, isFavorite) {
  if (!button) {
    return;
  }

  button.classList.toggle('is-active', Boolean(isFavorite));
}

const HEADER_DROPDOWN_CLOSE_DELAY_MS = 360;

function setAnimatedDropdownOpenState({
  dropdown,
  button,
  menu,
  isOpen,
  closeTimerRef,
}) {
  if (!dropdown || !button || !menu || !closeTimerRef) {
    return;
  }

  if (closeTimerRef.current) {
    clearTimeout(closeTimerRef.current);
    closeTimerRef.current = null;
  }

  if (isOpen) {
    menu.hidden = false;
    requestAnimationFrame(() => {
      dropdown.classList.add('is-open');
      button.setAttribute('aria-expanded', 'true');
    });
    return;
  }

  dropdown.classList.remove('is-open');
  button.setAttribute('aria-expanded', 'false');
  closeTimerRef.current = setTimeout(() => {
    if (!dropdown.classList.contains('is-open')) {
      menu.hidden = true;
    }
    closeTimerRef.current = null;
  }, HEADER_DROPDOWN_CLOSE_DELAY_MS);
}


function applyCurrentPhotoFilter() {
  currentPhotos = allCurrentMonthPhotos.filter(photoMatchesCurrentFilters);
}

function getOrientationFilterMeta(filterValue) {
  return ORIENTATION_FILTER_META[filterValue] || ORIENTATION_FILTER_META.all;
}

function setOrientationFilterMenuOpen(isOpen) {
  const nextOpen = Boolean(isOpen) && !isImporting && Boolean(currentSelection);
  isOrientationFilterMenuOpen = nextOpen;
  setAnimatedDropdownOpenState({
    dropdown: orientationFilterDropdown,
    button: orientationFilterButton,
    menu: orientationFilterMenu,
    isOpen: nextOpen,
    closeTimerRef: {
      get current() {
        return orientationFilterMenuCloseTimer;
      },
      set current(value) {
        orientationFilterMenuCloseTimer = value;
      },
    },
  });
}

function closeOrientationFilterMenu() {
  setOrientationFilterMenuOpen(false);
}

function getCurrentMonthPhotoLabelCatalog() {
  const labelMap = new Map();

  for (const photo of allCurrentMonthPhotos) {
    const seenNormalizedNames = new Set();

    for (const label of Array.isArray(photo.photoLabels) ? photo.photoLabels : []) {
      if (!label?.normalizedName || seenNormalizedNames.has(label.normalizedName)) {
        continue;
      }

      seenNormalizedNames.add(label.normalizedName);

      if (!labelMap.has(label.normalizedName)) {
        labelMap.set(label.normalizedName, {
          normalizedName: label.normalizedName,
          name: label.name || label.normalizedName,
          colorHex: label.colorHex || '#6D5EF6',
          photoCount: 0,
        });
      }

      labelMap.get(label.normalizedName).photoCount += 1;
    }
  }

  return Array.from(labelMap.values()).sort((left, right) =>
    left.name.localeCompare(right.name, 'ja')
  );
}

function getSelectedPhotoLabelEntries() {
  if (activePhotoLabelFilters.length === 0) {
    return [];
  }

  const labelCatalog = getCurrentMonthPhotoLabelCatalog();

  return activePhotoLabelFilters.map((normalizedName) => {
    const matchedLabel = labelCatalog.find(
      (label) => label.normalizedName === normalizedName
    );

    return {
      normalizedName,
      name: matchedLabel?.name || normalizedName,
      colorHex: matchedLabel?.colorHex || '#6D5EF6',
    };
  });
}

function getPhotoLabelFilterModeLabel() {
  return photoLabelFilterMode === 'and' ? 'AND' : 'OR';
}

function getSelectedPhotoLabelFilterText({
  includePrefix = true,
  includeMode = true,
} = {}) {
  if (activePhotoLabelFilters.length === 0) {
    return includePrefix ? 'ラベル: すべて' : 'すべて';
  }

  const joined = getSelectedPhotoLabelEntries()
    .map((label) => label.name)
    .join(' / ');

  const shouldShowMode = includeMode && activePhotoLabelFilters.length > 1;
  const modeSuffix = shouldShowMode ? ` ${getPhotoLabelFilterModeLabel()}` : '';

  if (includePrefix) {
    return `ラベル${modeSuffix}: ${joined}`;
  }

  return shouldShowMode ? `${getPhotoLabelFilterModeLabel()}: ${joined}` : joined;
}

function getPhotoLabelFilterButtonText() {
  return getSelectedPhotoLabelFilterText({ includePrefix: true });
}

function renderPhotoLabelFilterMenu() {
  if (!photoLabelFilterMenu) {
    return;
  }

  const labelCatalog = getCurrentMonthPhotoLabelCatalog();
  photoLabelFilterMenu.innerHTML = '';

  const modeToggle = document.createElement('div');
  modeToggle.className = 'header-dropdown-mode-toggle';

  for (const mode of ['or', 'and']) {
    const modeButton = document.createElement('button');
    modeButton.type = 'button';
    modeButton.className = 'header-dropdown-mode-button';
    modeButton.classList.toggle('is-active', photoLabelFilterMode === mode);
    modeButton.dataset.photoLabelFilterMode = mode;
    modeButton.setAttribute('aria-pressed', photoLabelFilterMode === mode ? 'true' : 'false');
    modeButton.textContent = mode === 'or' ? 'いずれか (OR)' : 'すべて (AND)';
    modeToggle.appendChild(modeButton);
  }

  photoLabelFilterMenu.appendChild(modeToggle);

  const allButton = document.createElement('button');
  allButton.type = 'button';
  allButton.className = 'header-dropdown-item';
  allButton.classList.toggle('is-active', activePhotoLabelFilters.length === 0);
  allButton.dataset.photoLabelFilter = '__all__';
  allButton.setAttribute('role', 'menuitemcheckbox');
  allButton.setAttribute(
    'aria-checked',
    activePhotoLabelFilters.length === 0 ? 'true' : 'false'
  );

  const allLabel = document.createElement('span');
  allLabel.className = 'header-dropdown-item-label';
  allLabel.textContent = 'すべて';
  allButton.appendChild(allLabel);

  const allCheck = document.createElement('span');
  allCheck.className = 'material-symbols-outlined header-dropdown-check';
  allCheck.textContent = 'check';
  allButton.appendChild(allCheck);

  photoLabelFilterMenu.appendChild(allButton);

  if (labelCatalog.length === 0) {
    const emptyState = document.createElement('div');
    emptyState.className = 'header-dropdown-empty';
    emptyState.textContent = 'この月にはラベルがありません';
    photoLabelFilterMenu.appendChild(emptyState);
    return;
  }

  for (const label of labelCatalog) {
    const isActive = activePhotoLabelFilters.includes(label.normalizedName);
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'header-dropdown-item header-dropdown-item-with-meta';
    item.classList.toggle('is-active', isActive);
    item.dataset.photoLabelFilter = label.normalizedName;
    item.setAttribute('role', 'menuitemcheckbox');
    item.setAttribute('aria-checked', isActive ? 'true' : 'false');

    const labelWrap = document.createElement('span');
    labelWrap.className = 'header-dropdown-item-meta';

    const swatch = document.createElement('span');
    swatch.className = 'header-dropdown-color-dot';
    swatch.style.setProperty('--photo-label-color', label.colorHex || '#6D5EF6');
    labelWrap.appendChild(swatch);

    const labelText = document.createElement('span');
    labelText.className = 'header-dropdown-item-label';
    labelText.textContent = label.name;
    labelWrap.appendChild(labelText);

    const side = document.createElement('span');
    side.className = 'header-dropdown-item-side';

    const count = document.createElement('span');
    count.className = 'header-dropdown-meta-badge';
    count.textContent = String(label.photoCount);
    side.appendChild(count);

    const check = document.createElement('span');
    check.className = 'material-symbols-outlined header-dropdown-check';
    check.textContent = 'check';
    side.appendChild(check);

    item.appendChild(labelWrap);
    item.appendChild(side);
    photoLabelFilterMenu.appendChild(item);
  }
}

function setPhotoLabelFilterMenuOpen(isOpen) {
  const nextOpen = Boolean(isOpen) && !isImporting && Boolean(currentSelection);
  isPhotoLabelFilterMenuOpen = nextOpen;

  if (nextOpen) {
    renderPhotoLabelFilterMenu();
  }

  setAnimatedDropdownOpenState({
    dropdown: photoLabelFilterDropdown,
    button: photoLabelFilterButton,
    menu: photoLabelFilterMenu,
    isOpen: nextOpen,
    closeTimerRef: {
      get current() {
        return photoLabelFilterMenuCloseTimer;
      },
      set current(value) {
        photoLabelFilterMenuCloseTimer = value;
      },
    },
  });
}

function closePhotoLabelFilterMenu() {
  setPhotoLabelFilterMenuOpen(false);
}

function normalizeWorldNameFilterText(value) {
  if (typeof value !== 'string') {
    return '';
  }

  return value.trim();
}

function getNormalizedWorldNameFilterText(value) {
  return normalizeWorldNameFilterText(value)
    .normalize('NFKC')
    .toLocaleLowerCase('ja-JP');
}

function getWorldNameFilterSummaryText({ includePrefix = true } = {}) {
  if (!activeWorldNameFilter) {
    return includePrefix ? 'World: すべて' : 'すべて';
  }

  return includePrefix ? `World: ${activeWorldNameFilter}` : activeWorldNameFilter;
}

function getWorldNameFilterButtonText() {
  return getWorldNameFilterSummaryText({ includePrefix: true });
}

function clearWorldNameFilterInputTimer() {
  if (!worldNameFilterInputTimer) {
    return;
  }

  clearTimeout(worldNameFilterInputTimer);
  worldNameFilterInputTimer = null;
}

function setWorldNameFilterMenuOpen(isOpen) {
  const nextOpen = Boolean(isOpen) && !isImporting && Boolean(currentSelection);
  isWorldNameFilterMenuOpen = nextOpen;
  setAnimatedDropdownOpenState({
    dropdown: worldNameFilterDropdown,
    button: worldNameFilterButton,
    menu: worldNameFilterMenu,
    isOpen: nextOpen,
    closeTimerRef: {
      get current() {
        return worldNameFilterMenuCloseTimer;
      },
      set current(value) {
        worldNameFilterMenuCloseTimer = value;
      },
    },
  });

  if (!nextOpen) {
    clearWorldNameFilterInputTimer();
    return;
  }

  requestAnimationFrame(() => {
    worldNameFilterInput?.focus({ preventScroll: true });
    worldNameFilterInput?.select();
  });
}

function closeWorldNameFilterMenu() {
  setWorldNameFilterMenuOpen(false);
}

function getManagedDropdownClosers() {
  return [
    {
      isOpen: () => isOrientationFilterMenuOpen,
      dropdown: orientationFilterDropdown,
      close: closeOrientationFilterMenu,
    },
    {
      isOpen: () => isPhotoLabelFilterMenuOpen,
      dropdown: photoLabelFilterDropdown,
      close: closePhotoLabelFilterMenu,
    },
    {
      isOpen: () => isWorldNameFilterMenuOpen,
      dropdown: worldNameFilterDropdown,
      close: closeWorldNameFilterMenu,
    },
    {
      isOpen: () => isRegenerateThumbnailMonthMenuOpen,
      dropdown: regenerateThumbnailMonthDropdown,
      close: closeRegenerateThumbnailMonthMenu,
    },
    {
      isOpen: () => isPhotoLabelCatalogMenuOpen,
      dropdown: photoLabelCatalogDropdown,
      close: () => setPhotoLabelCatalogMenuOpen(false),
    },
  ];
}

function closeManagedDropdownsFromOutsideClick(target) {
  for (const entry of getManagedDropdownClosers()) {
    if (entry.isOpen() && entry.dropdown && !entry.dropdown.contains(target)) {
      entry.close();
    }
  }
}

function closeManagedDropdownFromEscape() {
  for (const entry of getManagedDropdownClosers()) {
    if (entry.isOpen()) {
      entry.close();
      return true;
    }
  }

  return false;
}

function isAnyPhotoFilterActive() {
  return (
    isFavoriteFilterOnly ||
    activeOrientationFilter !== 'all' ||
    activePhotoLabelFilters.length > 0 ||
    activeWorldNameFilter.length > 0
  );
}

function getPhotoOrientationTier(photo) {
  if (!photo) {
    return null;
  }

  if (photo.orientationTier) {
    return photo.orientationTier;
  }

  const width = Number(photo.imageWidth);
  const height = Number(photo.imageHeight);

  if (!Number.isFinite(width) || !Number.isFinite(height)) {
    return null;
  }

  if (width === height) {
    return 'square';
  }

  return width > height ? 'landscape' : 'portrait';
}

function photoMatchesCurrentFilters(photo) {
  if (isFavoriteFilterOnly && !photo.isFavorite) {
    return false;
  }

  if (
    activeOrientationFilter !== 'all' &&
    getPhotoOrientationTier(photo) !== activeOrientationFilter
  ) {
    return false;
  }

  if (activePhotoLabelFilters.length > 0) {
    const photoLabelSet = new Set(
      (Array.isArray(photo.photoLabels) ? photo.photoLabels : [])
        .map((label) => label?.normalizedName)
        .filter(Boolean)
    );

    const matchesLabelFilter =
      photoLabelFilterMode === 'and'
        ? activePhotoLabelFilters.every((normalizedName) =>
            photoLabelSet.has(normalizedName)
          )
        : activePhotoLabelFilters.some((normalizedName) =>
            photoLabelSet.has(normalizedName)
          );

    if (!matchesLabelFilter) {
      return false;
    }
  }

  if (activeWorldNameFilter) {
    const targetText = [
      photo.worldName,
      photo.rawWorldName,
      photo.worldId,
    ]
      .filter(Boolean)
      .join(' ')
      .normalize('NFKC')
      .toLocaleLowerCase('ja-JP');

    if (!targetText.includes(getNormalizedWorldNameFilterText(activeWorldNameFilter))) {
      return false;
    }
  }

  return true;
}

function getDefaultMonthGalleryEmptyMessage() {
  return 'まだ写真がありません。画像 / フォルダをドラッグ&ドロップするか、設定から取り込めます';
}

// Keep all count / empty-state filter summaries driven by the same label list.
function getActivePhotoFilterSummaryParts() {
  const filterLabels = [];

  if (isFavoriteFilterOnly) {
    filterLabels.push('お気に入り');
  }

  if (activeOrientationFilter !== 'all') {
    filterLabels.push(getOrientationFilterMeta(activeOrientationFilter).shortLabel);
  }

  if (activePhotoLabelFilters.length > 0) {
    filterLabels.push(
      getSelectedPhotoLabelFilterText({
        includePrefix: true,
      })
    );
  }

  if (activeWorldNameFilter) {
    filterLabels.push(getWorldNameFilterSummaryText({ includePrefix: true }));
  }

  return filterLabels;
}

function buildCurrentMonthCountText() {
  if (!currentSelection) {
    return '0枚';
  }

  if (!isAnyPhotoFilterActive()) {
    return `${allCurrentMonthPhotos.length}枚`;
  }

  const filterLabels = getActivePhotoFilterSummaryParts();

  return `${currentPhotos.length}枚（${filterLabels.join(' / ')}） / 全${allCurrentMonthPhotos.length}枚`;
}

function buildFilteredEmptyMessage() {
  const filterLabels = getActivePhotoFilterSummaryParts();

  if (filterLabels.length === 0) {
    return getDefaultMonthGalleryEmptyMessage();
  }

  return `${filterLabels.join(' / ')} に一致する写真はありません`;
}

function setAnimatedHeaderText(
  element,
  nextText,
  { animate = true, durationMs = 1700 } = {}
) {
  if (!element) {
    return;
  }

  const resolvedText = typeof nextText === 'string' ? nextText : String(nextText ?? '');
  const host = element.parentElement;

  for (const existingClone of host?.querySelectorAll('.month-header-transition-clone') || []) {
    existingClone.remove();
  }

  element.classList.remove('month-header-text-enter');
  element.style.removeProperty('animation-duration');
  element.textContent = resolvedText;
}

function setAnimatedMonthCountText(nextText, { animate = true } = {}) {
  setAnimatedHeaderText(currentMonthCount, nextText, {
    animate,
    durationMs: 2840,
  });
}

function setAnimatedMonthLabelText(nextText, { animate = true } = {}) {
  setAnimatedHeaderText(currentMonthLabel, nextText, {
    animate,
    durationMs: 2320,
  });
}

function syncFavoriteFilterUi() {
  if (favoriteFilterButton) {
    favoriteFilterButton.classList.toggle('is-active', isFavoriteFilterOnly);
    favoriteFilterButton.disabled = isImporting || !currentSelection;

    const label = isFavoriteFilterOnly
      ? 'お気に入りのみ表示中'
      : 'お気に入りのみ表示';
    favoriteFilterButton.setAttribute('aria-label', label);
    favoriteFilterButton.setAttribute('title', label);
  }

  if (orientationFilterButton) {
    const orientationMeta = getOrientationFilterMeta(activeOrientationFilter);
    const label = `向きフィルタ: ${orientationMeta.shortLabel}`;

    orientationFilterButton.classList.toggle(
      'is-active',
      activeOrientationFilter !== 'all'
    );
    orientationFilterButton.disabled = isImporting || !currentSelection;
    orientationFilterButton.setAttribute('aria-label', label);
    orientationFilterButton.setAttribute('title', label);
  }

  if (orientationFilterLabel) {
    orientationFilterLabel.textContent = getOrientationFilterMeta(
      activeOrientationFilter
    ).buttonLabel;
  }

  if (photoLabelFilterButton) {
    const label = `ラベルフィルタ: ${getSelectedPhotoLabelFilterText({
      includePrefix: false,
    })}`;

    photoLabelFilterButton.classList.toggle(
      'is-active',
      activePhotoLabelFilters.length > 0
    );
    photoLabelFilterButton.disabled = isImporting || !currentSelection;
    photoLabelFilterButton.setAttribute('aria-label', label);
    photoLabelFilterButton.setAttribute('title', label);
  }

  if (photoLabelFilterLabel) {
    photoLabelFilterLabel.textContent = getPhotoLabelFilterButtonText();
  }

  for (const item of orientationFilterItems) {
    const isActive = item.dataset.orientationFilter === activeOrientationFilter;
    item.classList.toggle('is-active', isActive);
    item.setAttribute('aria-checked', isActive ? 'true' : 'false');
  }

  renderPhotoLabelFilterMenu();

  if (isImporting || !currentSelection) {
    closeOrientationFilterMenu();
    closePhotoLabelFilterMenu();
  }

  if (!currentSelection) {
    setAnimatedMonthCountText('0枚', { animate: false });
    return;
  }

  setAnimatedMonthCountText(buildCurrentMonthCountText(), { animate: false });

  if (monthGalleryEmpty) {
    monthGalleryEmpty.textContent =
      allCurrentMonthPhotos.length > 0 && currentPhotos.length === 0
        ? buildFilteredEmptyMessage()
        : getDefaultMonthGalleryEmptyMessage();
  }
}

async function togglePhotoLabelFilter(normalizedName) {
  if (!currentSelection) {
    closePhotoLabelFilterMenu();
    return;
  }

  if (normalizedName === '__all__') {
    activePhotoLabelFilters = [];
  } else if (typeof normalizedName === 'string' && normalizedName.length > 0) {
    if (activePhotoLabelFilters.includes(normalizedName)) {
      activePhotoLabelFilters = activePhotoLabelFilters.filter(
        (value) => value !== normalizedName
      );
    } else {
      activePhotoLabelFilters = [...activePhotoLabelFilters, normalizedName];
    }
  }

  applyCurrentPhotoFilter();
  syncFavoriteFilterUi();
  await refreshCurrentMonthWithFilterAnimation();
}

async function setPhotoLabelFilterMode(nextMode) {
  const normalizedMode = nextMode === 'and' ? 'and' : 'or';

  if (!currentSelection || photoLabelFilterMode === normalizedMode) {
    return;
  }

  photoLabelFilterMode = normalizedMode;
  applyCurrentPhotoFilter();
  syncFavoriteFilterUi();

  if (activePhotoLabelFilters.length === 0) {
    return;
  }

  await refreshCurrentMonthWithFilterAnimation();
}

async function applyWorldNameFilter(nextValue) {
  const normalizedValue = normalizeWorldNameFilterText(nextValue);

  if (!currentSelection || activeWorldNameFilter === normalizedValue) {
    return;
  }

  activeWorldNameFilter = normalizedValue;
  applyCurrentPhotoFilter();
  syncFavoriteFilterUi();
  await refreshCurrentMonthWithFilterAnimation();
}

function scheduleWorldNameFilterApply(nextValue) {
  clearWorldNameFilterInputTimer();

  worldNameFilterInputTimer = setTimeout(() => {
    worldNameFilterInputTimer = null;
    void applyWorldNameFilter(nextValue);
  }, 180);
}

async function setOrientationFilter(nextFilter) {
  const normalizedNextFilter = ORIENTATION_FILTER_META[nextFilter]
    ? nextFilter
    : 'all';

  if (!currentSelection || activeOrientationFilter === normalizedNextFilter) {
    closeOrientationFilterMenu();
    return;
  }

  activeOrientationFilter = normalizedNextFilter;
  applyCurrentPhotoFilter();
  syncFavoriteFilterUi();
  closeOrientationFilterMenu();
  await refreshCurrentMonthWithFilterAnimation();
}

const selectionCountBadge = document.getElementById('selection-count-badge');

function getSelectedPhotosFromCurrentCollections() {
  if (selectedPhotoIds.size === 0) {
    return [];
  }

  return allCurrentMonthPhotos.filter((photo) => selectedPhotoIds.has(photo.id));
}

function getBulkFavoriteTargetValue() {
  const selectedPhotos = getSelectedPhotosFromCurrentCollections();

  if (selectedPhotos.length === 0) {
    return true;
  }

  return !selectedPhotos.every((photo) => photo.isFavorite);
}

function syncSelectionUi() {
  if (selectionModeButton) {
    selectionModeButton.classList.toggle('is-active', isSelectionMode);
    selectionModeButton.textContent = isSelectionMode ? '選択終了' : '選択';
    selectionModeButton.disabled = isImporting || !currentSelection;
  }

  if (bulkFavoriteButton) {
    const hasSelection = isSelectionMode && selectedPhotoIds.size > 0;
    const nextFavoriteValue = getBulkFavoriteTargetValue();

    bulkFavoriteButton.disabled = isImporting || !hasSelection;
    bulkFavoriteButton.textContent = nextFavoriteValue
      ? 'お気に入り'
      : 'お気に入り解除';
    bulkFavoriteButton.classList.toggle(
      'is-active',
      hasSelection && !nextFavoriteValue
    );
  }

  if (bulkDeleteButton) {
    bulkDeleteButton.disabled =
      isImporting || !isSelectionMode || selectedPhotoIds.size === 0;

    bulkDeleteButton.textContent = '削除';
  }

  if (selectionCountBadge) {
    const shouldShow =
      Boolean(currentSelection) && isSelectionMode && selectedPhotoIds.size > 0;

    selectionCountBadge.classList.toggle('hidden', !shouldShow);

    if (shouldShow) {
      selectionCountBadge.textContent = `${selectedPhotoIds.size}件選択中`;
    } else {
      selectionCountBadge.textContent = '0件選択中';
    }
  }
}

function clearSelectionState() {
  isSelectionMode = false;
  selectedPhotoIds.clear();
  syncRenderedSelectionState();
  syncSelectionUi();
}

function syncRenderedPhotoSelectionState(photoId) {
  if (!monthGalleryList) {
    return false;
  }

  const card = monthGalleryList.querySelector(
    `.photo-card[data-photo-id="${photoId}"]`
  );

  if (!card) {
    return false;
  }

  const isSelected = selectedPhotoIds.has(photoId);
  const selectionButton = card.querySelector('.photo-card-selection-btn');

  card.classList.toggle('selection-mode', isSelectionMode);
  card.classList.toggle('is-selected', isSelected);
  selectionButton?.classList.toggle('is-selected', isSelected);

  return true;
}

function syncRenderedSelectionState() {
  if (!monthGalleryList) {
    return;
  }

  for (const card of monthGalleryList.querySelectorAll('.photo-card')) {
    const photoId = Number(card.dataset.photoId);
    const isSelected =
      Number.isFinite(photoId) && selectedPhotoIds.has(photoId);
    const selectionButton = card.querySelector('.photo-card-selection-btn');

    card.classList.toggle('selection-mode', isSelectionMode);
    card.classList.toggle('is-selected', isSelected);
    selectionButton?.classList.toggle('is-selected', isSelected);
  }
}

function togglePhotoSelection(photoId) {
  if (selectedPhotoIds.has(photoId)) {
    selectedPhotoIds.delete(photoId);
  } else {
    selectedPhotoIds.add(photoId);
  }

  syncRenderedPhotoSelectionState(photoId);
  syncSelectionUi();
}

function getLatestKnownPhotoById(photoId) {
  if (!Number.isInteger(photoId) || photoId <= 0) {
    return null;
  }

  return (
    currentPhotos.find((photo) => photo.id === photoId) ||
    allCurrentMonthPhotos.find((photo) => photo.id === photoId) ||
    (currentModalPhoto?.id === photoId ? currentModalPhoto : null)
  );
}

function updatePhotoInCurrentCollections(updatedPhoto) {
  allCurrentMonthPhotos = allCurrentMonthPhotos.map((photo) =>
    photo.id === updatedPhoto.id ? updatedPhoto : photo
  );
  applyCurrentPhotoFilter();
}

// Single-photo updates from modal actions should keep collections, cards, and
// the open modal in sync without each caller repeating the same flow.
function syncSinglePhotoUpdate(updatedPhoto, { refreshModal = true } = {}) {
  if (!updatedPhoto) {
    return null;
  }

  updatePhotoInCurrentCollections(updatedPhoto);

  const isVisibleAfterFilters = currentPhotos.some(
    (photo) => photo.id === updatedPhoto.id
  );
  let rerenderedMonthGallery = false;

  if (!isVisibleAfterFilters && isAnyPhotoFilterActive()) {
    if (!removeRenderedPhotoCards([updatedPhoto.id])) {
      renderMonthGallery({ resetProgressive: true });
      rerenderedMonthGallery = true;
    }
  } else if (!replaceRenderedPhotoCard(updatedPhoto) && currentSelection) {
    renderMonthGallery({ resetProgressive: true });
    rerenderedMonthGallery = true;
  }

  if (!rerenderedMonthGallery) {
    syncFavoriteFilterUi();
  }

  const resolvedPhoto = getLatestKnownPhotoById(updatedPhoto.id) || updatedPhoto;

  if (refreshModal && currentModalPhoto?.id === updatedPhoto.id) {
    showImageModalPhoto(resolvedPhoto);
  }

  return resolvedPhoto;
}

function syncBatchPhotoUpdates(updatedPhotos, { refreshModal = true } = {}) {
  const normalizedUpdates = (Array.isArray(updatedPhotos) ? updatedPhotos : []).filter(
    (photo) => Number.isInteger(photo?.id)
  );

  if (normalizedUpdates.length === 0) {
    return null;
  }

  const updatedPhotoMap = new Map(
    normalizedUpdates.map((photo) => [photo.id, photo])
  );

  allCurrentMonthPhotos = allCurrentMonthPhotos.map(
    (photo) => updatedPhotoMap.get(photo.id) || photo
  );
  applyCurrentPhotoFilter();

  let rerenderedMonthGallery = false;

  for (const photo of normalizedUpdates) {
    const isVisibleAfterFilters = currentPhotos.some(
      (currentPhoto) => currentPhoto.id === photo.id
    );

    if (!isVisibleAfterFilters && isAnyPhotoFilterActive()) {
      rerenderedMonthGallery = true;
      break;
    }

    if (!replaceRenderedPhotoCard(photo) && currentSelection) {
      rerenderedMonthGallery = true;
      break;
    }
  }

  if (rerenderedMonthGallery && currentSelection) {
    renderMonthGallery({ resetProgressive: true });
  } else {
    syncFavoriteFilterUi();
  }

  const nextModalPhoto = currentModalPhoto
    ? getLatestKnownPhotoById(currentModalPhoto.id)
    : null;

  if (refreshModal && nextModalPhoto) {
    showImageModalPhoto(nextModalPhoto);
  }

  return nextModalPhoto;
}

function clearThumbnailCacheInCurrentCollections() {
  if (!allCurrentMonthPhotos.length) {
    return;
  }

  allCurrentMonthPhotos = allCurrentMonthPhotos.map((photo) => ({
    ...photo,
    thumbnailPath: null,
    thumbnailUrl: null,
  }));
  applyCurrentPhotoFilter();
}

function replaceRenderedPhotoCard(updatedPhoto) {
  if (!updatedPhoto || !monthGalleryList) {
    return false;
  }

  const currentCard = monthGalleryList.querySelector(
    `.photo-card[data-photo-id="${updatedPhoto.id}"]`
  );

  if (!currentCard) {
    return false;
  }

  const nextCard = createPhotoCard(updatedPhoto);
  currentCard.replaceWith(nextCard);
  return true;
}

function removePhotoFromCurrentCollections(photoId) {
  allCurrentMonthPhotos = allCurrentMonthPhotos.filter(
    (photo) => photo.id !== photoId
  );
  applyCurrentPhotoFilter();
}

function removePhotosFromCurrentCollections(photoIds) {
  const targetIdSet = new Set(photoIds);

  allCurrentMonthPhotos = allCurrentMonthPhotos.filter(
    (photo) => !targetIdSet.has(photo.id)
  );
  applyCurrentPhotoFilter();
}

function removeRenderedPhotoCards(
  photoIds,
  { refillRenderedSlots = true } = {}
) {
  if (!monthGalleryList) {
    return false;
  }

  const targetIds = Array.from(
    new Set(
      (Array.isArray(photoIds) ? photoIds : []).filter(
        (photoId) => Number.isInteger(photoId) && photoId > 0
      )
    )
  );

  let removedRenderedCount = 0;

  for (const photoId of targetIds) {
    const card = monthGalleryList.querySelector(
      `.photo-card[data-photo-id="${photoId}"]`
    );

    if (!card) {
      continue;
    }

    const groupGrid = card.parentElement;
    const groupSection = card.closest('.gallery-group');

    card.remove();
    removedRenderedCount += 1;

    if (groupGrid && !groupGrid.querySelector('.photo-card') && groupSection) {
      groupSection.remove();

      for (const [groupDate, groupState] of renderedGalleryGroupMap.entries()) {
        if (groupState.section === groupSection) {
          renderedGalleryGroupMap.delete(groupDate);
          break;
        }
      }
    }
  }

  if (removedRenderedCount === 0) {
    return false;
  }

  renderedPhotoCount = Math.max(0, renderedPhotoCount - removedRenderedCount);

  if (refillRenderedSlots && renderedPhotoCount < currentPhotos.length) {
    appendMonthGalleryPhotoBatch(
      Math.min(currentPhotos.length, renderedPhotoCount + removedRenderedCount)
    );
  }

  const hasRenderedCards = Boolean(monthGalleryList.querySelector('.photo-card'));

  if (monthGalleryEmpty) {
    monthGalleryEmpty.style.display = hasRenderedCards ? 'none' : 'block';
  }

  if (!hasRenderedCards && currentPhotos.length === 0) {
    monthGalleryList.innerHTML = '';
  }

  scheduleMonthGalleryLoadCheck({ immediate: true });
  return true;
}

function hasFullyRenderedMonthGallery() {
  if (!monthGalleryList || allCurrentMonthPhotos.length === 0) {
    return false;
  }

  return (
    monthGalleryList.querySelectorAll('.photo-card').length >=
    allCurrentMonthPhotos.length
  );
}

function syncRenderedFavoriteFilterState() {
  if (!monthGalleryList) {
    return false;
  }

  const renderedCards = monthGalleryList.querySelectorAll('.photo-card');

  if (renderedCards.length === 0) {
    return false;
  }

  const visiblePhotoIds = isAnyPhotoFilterActive()
    ? new Set(currentPhotos.map((photo) => photo.id))
    : null;
  let visibleCardCount = 0;

  for (const card of renderedCards) {
    const photoId = Number(card.dataset.photoId);
    const shouldHide = visiblePhotoIds
      ? !Number.isFinite(photoId) || !visiblePhotoIds.has(photoId)
      : false;

    card.hidden = shouldHide;

    if (!shouldHide) {
      visibleCardCount += 1;
    }
  }

  for (const groupSection of monthGalleryList.querySelectorAll('.gallery-group')) {
    groupSection.hidden = !groupSection.querySelector('.photo-card:not([hidden])');
  }

  if (monthGalleryEmpty) {
    monthGalleryEmpty.style.display = visibleCardCount > 0 ? 'none' : 'block';
  }

  if (monthGalleryEmpty && visibleCardCount === 0) {
    monthGalleryEmpty.textContent = buildFilteredEmptyMessage();
  }

  return true;
}

function syncSidebarSelectionState() {
  if (!sidebarTree || sidebarTree.children.length === 0) {
    return false;
  }

  let hasSidebarEntries = false;

  for (const yearBlock of sidebarTree.querySelectorAll('.year-block')) {
    const year = Number(yearBlock.dataset.year);
    const toggle = yearBlock.querySelector('.year-toggle');
    const monthList = yearBlock.querySelector('.month-list');
    const isExpanded = expandedYears.has(year);

    hasSidebarEntries = true;
    monthList?.classList.toggle('hidden', !isExpanded);

    if (toggle) {
      toggle.textContent = isExpanded ? '▾' : '▸';
    }

    for (const monthButton of yearBlock.querySelectorAll('.month-button')) {
      const month = Number(monthButton.dataset.month);
      const isActive =
        Number.isFinite(year) &&
        Number.isFinite(month) &&
        currentSelection &&
        currentSelection.year === year &&
        currentSelection.month === month;

      monthButton.classList.toggle('active', Boolean(isActive));
    }
  }

  return hasSidebarEntries;
}

function applySidebarDeletionLocally(targetSelection, removedCount) {
  if (
    !targetSelection ||
    !Number.isFinite(removedCount) ||
    removedCount <= 0 ||
    sidebarData.length === 0
  ) {
    return false;
  }

  let changed = false;

  sidebarData = sidebarData
    .map((yearEntry) => {
      if (yearEntry.year !== targetSelection.year) {
        return yearEntry;
      }

      const nextMonths = yearEntry.months
        .map((monthEntry) => {
          if (monthEntry.month !== targetSelection.month) {
            return monthEntry;
          }

          changed = true;

          return {
            ...monthEntry,
            count: Math.max(0, monthEntry.count - removedCount),
          };
        })
        .filter((monthEntry) => monthEntry.count > 0);

      return {
        ...yearEntry,
        totalCount: Math.max(0, yearEntry.totalCount - removedCount),
        months: nextMonths,
      };
    })
    .filter((yearEntry) => yearEntry.totalCount > 0 && yearEntry.months.length > 0);

  if (changed) {
    renderSidebar();
  }

  return changed;
}

function getLatestSelectionFromSidebarData() {
  const latestYearEntry = sidebarData[0];
  const latestMonthEntry = latestYearEntry?.months?.[0];

  if (!latestYearEntry || !latestMonthEntry) {
    return null;
  }

  return {
    year: latestYearEntry.year,
    month: latestMonthEntry.month,
  };
}

function populateModal(item) {
  if (!item) {
    return;
  }

  modalImage.src = item.fileUrl;
  modalWorldLink.textContent = item.worldName || 'ワールド名未取得';

  if (item.worldUrl) {
    modalWorldLink.href = item.worldUrl;
    modalWorldLink.classList.remove('disabled');
    modalOpenWorldButton.disabled = false;
  } else {
    modalWorldLink.href = '#';
    modalWorldLink.classList.add('disabled');
    modalOpenWorldButton.disabled = true;
  }

  if (modalOpenOriginalButton) {
    modalOpenOriginalButton.disabled = !item.filePath;
  }

  if (modalOpenFolderButton) {
    modalOpenFolderButton.disabled = !item.filePath;
  }

  if (modalFavoriteButton) {
    syncFavoriteButtonState(modalFavoriteButton, item.isFavorite);
  }

  if (modalFavoriteIcon) {
    modalFavoriteIcon.textContent = 'star';
  }

  setText(modalResolutionTier, item.resolutionTier);
  if (modalResolutionHeroBadge) {
    const hasResolutionTier =
      typeof item.resolutionTier === 'string' &&
      item.resolutionTier.trim().length > 0;

    modalResolutionHeroBadge.textContent = hasResolutionTier
      ? item.resolutionTier.trim()
      : '';
    modalResolutionHeroBadge.classList.toggle('is-hidden', !hasResolutionTier);
  }

  if (modalTakenAtHero) {
    const hasTakenAt =
      typeof item.takenAt === 'string' && item.takenAt.trim().length > 0;

    modalTakenAtHero.textContent = hasTakenAt ? item.takenAt.trim() : '';
    modalTakenAtHero.classList.toggle('is-hidden', !hasTakenAt);
  }

  modalResolutionTier?.parentElement?.classList.add('is-hidden');
  modalTakenAt?.parentElement?.classList.add('is-hidden');
  modalFileName.textContent = item.fileName || 'ファイル名不明';
  setText(modalTakenAt, item.takenAt, '未取得');
  setText(modalWorldName, item.worldName, '未取得');
  setText(modalWorldId, item.worldId, '未取得');

  if (modalPhotoMemoInput) {
    modalPhotoMemoInput.value =
      typeof item.memoText === 'string' ? item.memoText : '';
  }

  if (modalPhotoMemoStatus) {
    modalPhotoMemoStatus.textContent = '';
  }

  if (modalPhotoMemoSaveButton) {
    modalPhotoMemoSaveButton.disabled = !item.id;
  }

  resizeModalPhotoMemoInput();
}

function resizeModalPhotoMemoInput() {
  if (!modalPhotoMemoInput) {
    return;
  }

  modalPhotoMemoInput.style.height = 'auto';

  const computedStyle = window.getComputedStyle(modalPhotoMemoInput);
  const lineHeight = Number.parseFloat(computedStyle.lineHeight) || 22;
  const paddingTop = Number.parseFloat(computedStyle.paddingTop) || 0;
  const paddingBottom = Number.parseFloat(computedStyle.paddingBottom) || 0;
  const borderTop = Number.parseFloat(computedStyle.borderTopWidth) || 0;
  const borderBottom = Number.parseFloat(computedStyle.borderBottomWidth) || 0;
  const verticalFrame = paddingTop + paddingBottom + borderTop + borderBottom;
  const minHeight = Math.ceil(lineHeight + verticalFrame);
  const maxHeight = Math.ceil(lineHeight * 3 + verticalFrame);
  const nextHeight = Math.max(
    minHeight,
    Math.min(modalPhotoMemoInput.scrollHeight, maxHeight)
  );

  modalPhotoMemoInput.style.height = `${nextHeight}px`;
  modalPhotoMemoInput.style.overflowY =
    modalPhotoMemoInput.scrollHeight > maxHeight ? 'auto' : 'hidden';
}

function formatOfficialWorldTagLabel(tag) {
  if (typeof tag !== 'string' || tag.trim().length === 0) {
    return null;
  }

  return tag
    .replace(/^author_tag_/i, '')
    .replace(/^system_/i, 'system ')
    .replace(/_/g, ' ')
    .trim();
}

function renderModalWorldTags(tags) {
  if (!modalWorldTags) {
    return;
  }

  const normalizedTags = Array.isArray(tags)
    ? tags.map(formatOfficialWorldTagLabel).filter(Boolean)
    : [];

  if (normalizedTags.length === 0) {
    modalWorldTags.innerHTML =
      '<span class="modal-world-tag is-placeholder">未取得</span>';
    return;
  }

  modalWorldTags.innerHTML = normalizedTags
    .map((tag) => `<span class="modal-world-tag">${escapeHtml(tag)}</span>`)
    .join('');
}

function normalizePhotoLabelEntry(label) {
  if (!label || typeof label.name !== 'string') {
    return null;
  }

  const name = label.name.trim();

  if (!name) {
    return null;
  }

  const normalizedName =
    typeof label.normalizedName === 'string' && label.normalizedName.trim().length > 0
      ? label.normalizedName.trim().toLowerCase()
      : name.normalize('NFC').toLowerCase();
  const colorMatch = String(label.colorHex || '').trim().match(/^#?([0-9a-fA-F]{6})$/);

  return {
    id: Number.isInteger(label.id) ? label.id : null,
    name,
    normalizedName,
    colorHex: colorMatch ? `#${colorMatch[1].toUpperCase()}` : '#6D5EF6',
    photoCount: Number(label.photoCount || 0) || 0,
  };
}

function sortPhotoLabels(labels) {
  return (Array.isArray(labels) ? labels : [])
    .map(normalizePhotoLabelEntry)
    .filter(Boolean)
    .sort((left, right) =>
      left.name.localeCompare(right.name, 'ja', { sensitivity: 'base' })
    );
}

function createPhotoLabelChipElement(
  label,
  { removable = false, onRemove = null } = {}
) {
  const chip = document.createElement('span');
  chip.className = 'photo-label-chip';
  chip.style.setProperty('--photo-label-color', label.colorHex);

  const swatch = document.createElement('span');
  swatch.className = 'photo-label-chip-swatch';
  chip.appendChild(swatch);

  const text = document.createElement('span');
  text.className = 'photo-label-chip-text';
  text.textContent = label.name;
  chip.appendChild(text);

  if (removable && typeof onRemove === 'function') {
    chip.classList.add('is-removable');

    const removeButton = document.createElement('button');
    removeButton.type = 'button';
    removeButton.className = 'photo-label-chip-remove';
    removeButton.textContent = '×';
    removeButton.setAttribute('aria-label', `${label.name} を外す`);
    removeButton.addEventListener('click', () => {
      onRemove(label);
    });
    chip.appendChild(removeButton);
  }

  return chip;
}

function renderPhotoLabelChipList(
  container,
  labels,
  { removable = false, onRemove = null, placeholder = '未設定' } = {}
) {
  if (!container) {
    return;
  }

  container.innerHTML = '';
  const normalizedLabels = sortPhotoLabels(labels);

  if (normalizedLabels.length === 0) {
    const placeholderChip = document.createElement('span');
    placeholderChip.className = 'photo-label-chip is-placeholder';
    placeholderChip.textContent = placeholder;
    container.appendChild(placeholderChip);
    return;
  }

  normalizedLabels.forEach((label) => {
    container.appendChild(
      createPhotoLabelChipElement(label, {
        removable,
        onRemove,
      })
    );
  });
}

function createPhotoCardLabelChip(label) {
  const chip = document.createElement('span');
  chip.className = 'photo-card-label-chip';
  chip.style.setProperty('--photo-label-color', label.colorHex || '#6D5EF6');

  const swatch = document.createElement('span');
  swatch.className = 'photo-card-label-chip-swatch';
  chip.appendChild(swatch);

  const text = document.createElement('span');
  text.className = 'photo-card-label-chip-text';
  text.textContent = label.name;
  chip.appendChild(text);

  return chip;
}

function renderModalPhotoLabels() {
  renderPhotoLabelChipList(modalPhotoLabelsList, currentModalPhotoLabels, {
    placeholder: '未設定',
  });

  if (openPhotoLabelEditorButton) {
    openPhotoLabelEditorButton.disabled = !currentModalPhoto;
  }
}

function setPhotoLabelCatalogMenuOpen(isOpen) {
  isPhotoLabelCatalogMenuOpen = Boolean(isOpen);
  setAnimatedDropdownOpenState({
    dropdown: photoLabelCatalogDropdown,
    button: photoLabelCatalogButton,
    menu: photoLabelCatalogMenu,
    isOpen: isPhotoLabelCatalogMenuOpen,
    closeTimerRef: {
      get current() {
        return photoLabelCatalogMenuCloseTimer;
      },
      set current(value) {
        photoLabelCatalogMenuCloseTimer = value;
      },
    },
  });
}

// The label catalog dropdown behaves like a single-purpose picker: choose one
// existing label and append it immediately to the draft selection list.
function selectPhotoLabelCatalogOption(normalizedName) {
  activePhotoLabelCatalogSelection = normalizedName || '';
  renderPhotoLabelCatalogOptions();
}

function renderPhotoLabelCatalogOptions() {
  if (!photoLabelCatalogButton || !photoLabelCatalogMenu) {
    return;
  }

  const selectedNames = new Set(
    draftModalPhotoLabels.map((label) => label.normalizedName)
  );
  const availableOptions = photoLabelCatalog.filter(
    (label) => !selectedNames.has(label.normalizedName)
  );
  const activeOption = availableOptions.find(
    (label) => label.normalizedName === activePhotoLabelCatalogSelection
  );
  const selectedOption = activeOption || null;

  activePhotoLabelCatalogSelection = selectedOption?.normalizedName || '';

  photoLabelCatalogButton.innerHTML = '';
  photoLabelCatalogButton.disabled = availableOptions.length === 0;
  photoLabelCatalogButton.classList.toggle('is-placeholder', !selectedOption);

  const buttonLabel = document.createElement('span');
  buttonLabel.className = 'photo-label-catalog-button-label';
  buttonLabel.textContent = selectedOption
    ? selectedOption.photoCount > 0
      ? `${selectedOption.name} (${selectedOption.photoCount})`
      : selectedOption.name
    : availableOptions.length > 0
      ? '既存ラベルを選択'
      : '追加できるラベルはありません';
  photoLabelCatalogButton.appendChild(buttonLabel);

  const buttonChevron = document.createElement('span');
  buttonChevron.className =
    'material-symbols-outlined orientation-filter-chevron photo-label-catalog-button-chevron';
  buttonChevron.textContent = 'expand_more';
  photoLabelCatalogButton.appendChild(buttonChevron);

  photoLabelCatalogMenu.innerHTML = '';

  if (availableOptions.length === 0) {
    const emptyState = document.createElement('div');
    emptyState.className = 'header-dropdown-empty photo-label-catalog-empty';
    emptyState.textContent = '追加できるラベルはありません';
    photoLabelCatalogMenu.appendChild(emptyState);
  }

  availableOptions.forEach((label) => {
    const optionButton = document.createElement('button');
    optionButton.type = 'button';
    optionButton.className =
      'header-dropdown-item header-dropdown-item-with-meta photo-label-catalog-option';
    optionButton.classList.toggle(
      'is-active',
      label.normalizedName === activePhotoLabelCatalogSelection
    );
    optionButton.setAttribute('role', 'menuitemradio');
    optionButton.setAttribute(
      'aria-checked',
      label.normalizedName === activePhotoLabelCatalogSelection ? 'true' : 'false'
    );

    const optionMeta = document.createElement('span');
    optionMeta.className = 'header-dropdown-item-meta';

    const optionName = document.createElement('span');
    optionName.className =
      'header-dropdown-item-label photo-label-catalog-option-name';
    optionName.textContent = label.name;
    optionMeta.appendChild(optionName);
    optionButton.appendChild(optionMeta);

    if (label.photoCount > 0) {
      const optionSide = document.createElement('span');
      optionSide.className = 'header-dropdown-item-side';

      const optionCount = document.createElement('span');
      optionCount.className =
        'header-dropdown-meta-badge photo-label-catalog-option-count';
      optionCount.textContent = String(label.photoCount);
      optionSide.appendChild(optionCount);
      optionButton.appendChild(optionSide);
    }

    optionButton.addEventListener('click', () => {
      selectPhotoLabelCatalogOption(label.normalizedName);
      addSelectedPhotoLabel();
    });

    photoLabelCatalogMenu.appendChild(optionButton);
  });

}

function renderPhotoLabelEditorSelectedList() {
  renderPhotoLabelChipList(photoLabelSelectedList, draftModalPhotoLabels, {
    removable: true,
    onRemove: (label) => {
      draftModalPhotoLabels = draftModalPhotoLabels.filter(
        (entry) => entry.normalizedName !== label.normalizedName
      );
      renderPhotoLabelEditorSelectedList();
      renderPhotoLabelCatalogOptions();
    },
    placeholder: 'ラベルはまだ設定されていません',
  });
}

function setPhotoLabelNewFormOpenLegacy(isOpen) {
  isPhotoLabelNewFormOpen = Boolean(isOpen);

  if (photoLabelNewForm) {
    photoLabelNewForm.hidden = !isPhotoLabelNewFormOpen;
  }

  if (photoLabelNewToggleButton) {
    photoLabelNewToggleButton.textContent = isPhotoLabelNewFormOpen
      ? '新規追加を閉じる'
      : '新規追加';
  }
}

function resetPhotoLabelNewFormLegacy() {
  if (photoLabelNewNameInput) {
    photoLabelNewNameInput.value = '';
  }

  if (photoLabelNewColorInput) {
    photoLabelNewColorInput.value = '#6D5EF6';
  }
}

function normalizePhotoLabelColorHex(colorHex) {
  const matched = String(colorHex || '').trim().match(/^#?([0-9a-fA-F]{6})$/);
  return matched ? `#${matched[1].toUpperCase()}` : PHOTO_LABEL_PRESET_COLORS[0];
}

function renderPhotoLabelPresetButtons() {
  if (!photoLabelPresetList || !photoLabelNewColorInput) {
    return;
  }

  const currentColor = normalizePhotoLabelColorHex(photoLabelNewColorInput.value);
  photoLabelPresetList.innerHTML = '';

  PHOTO_LABEL_PRESET_COLORS.forEach((colorHex) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'photo-label-preset-button';
    button.style.setProperty('--photo-label-color', colorHex);
    button.setAttribute('aria-label', `色 ${colorHex}`);
    button.classList.toggle('is-active', currentColor === colorHex);
    button.addEventListener('click', () => {
      setPhotoLabelDraftColor(colorHex);
    });
    photoLabelPresetList.appendChild(button);
  });
}

function setPhotoLabelDraftColor(colorHex) {
  const normalizedColor = normalizePhotoLabelColorHex(colorHex);

  if (photoLabelNewColorInput) {
    photoLabelNewColorInput.value = normalizedColor;
  }

  if (photoLabelNewColorPreview) {
    photoLabelNewColorPreview.style.setProperty(
      '--photo-label-color',
      normalizedColor
    );
  }

  renderPhotoLabelPresetButtons();
}

function setPhotoLabelNewFormOpen() {
  if (photoLabelNewForm) {
    photoLabelNewForm.hidden = false;
  }
}

function resetPhotoLabelNewForm() {
  if (photoLabelNewNameInput) {
    photoLabelNewNameInput.value = '';
  }

  setPhotoLabelDraftColor(PHOTO_LABEL_PRESET_COLORS[0]);
}

async function loadPhotoLabelCatalog() {
  const labels = await window.electronAPI.getLabelCatalog();
  photoLabelCatalog = sortPhotoLabels(labels);
  renderPhotoLabelCatalogOptions();
  return photoLabelCatalog;
}

async function loadModalPhotoLabels(item) {
  const requestId = ++modalPhotoLabelsRequestId;

  renderPhotoLabelChipList(modalPhotoLabelsList, [], {
    placeholder: '読み込み中...',
  });

  if (!item?.id) {
    currentModalPhotoLabels = [];
    renderModalPhotoLabels();
    return [];
  }

  try {
    const labels = await window.electronAPI.getPhotoLabels(item.id);

    if (
      requestId !== modalPhotoLabelsRequestId ||
      !currentModalPhoto ||
      currentModalPhoto.id !== item.id
    ) {
      return [];
    }

    currentModalPhotoLabels = sortPhotoLabels(labels);
    renderModalPhotoLabels();
    return currentModalPhotoLabels;
  } catch {
    if (
      requestId !== modalPhotoLabelsRequestId ||
      !currentModalPhoto ||
      currentModalPhoto.id !== item.id
    ) {
      return [];
    }

    currentModalPhotoLabels = [];
    renderModalPhotoLabels();
    return [];
  }
}

function clearSubModalAnimationTimer(modal) {
  if (!modal) {
    return;
  }

  const timer = subModalAnimationTimers.get(modal);

  if (!timer) {
    return;
  }

  clearTimeout(timer);
  subModalAnimationTimers.delete(modal);
}

// Keep sub-modal close wiring consistent across settings, confirm, and edit dialogs.
function bindSubModalCloseTriggers(backdrop, closeButton, closeHandler) {
  backdrop?.addEventListener('click', closeHandler);
  closeButton?.addEventListener('click', closeHandler);
}

function openSubModalElement(modal) {
  if (!modal) {
    return;
  }

  clearSubModalAnimationTimer(modal);
  modal.classList.remove('hidden', 'is-closing');
  void modal.offsetWidth;
  modal.classList.add('is-open');
}

function closeSubModalElement(modal, { onClosed } = {}) {
  if (!modal || modal.classList.contains('hidden')) {
    return;
  }

  const prefersReducedMotion = window.matchMedia?.(
    '(prefers-reduced-motion: reduce)'
  )?.matches;

  clearSubModalAnimationTimer(modal);
  modal.classList.remove('is-open');
  modal.classList.add('is-closing');

  const finalizeSubModalClose = () => {
    modal.classList.remove('is-closing');
    modal.classList.add('hidden');
    subModalAnimationTimers.delete(modal);

    if (typeof onClosed === 'function') {
      onClosed();
    }
  };

  if (prefersReducedMotion) {
    finalizeSubModalClose();
    return;
  }

  const timer = setTimeout(finalizeSubModalClose, SUB_MODAL_ANIMATION_MS);
  subModalAnimationTimers.set(modal, timer);
}

function closePhotoLabelModal() {
  if (!photoLabelModal) {
    return;
  }

  setPhotoLabelCatalogMenuOpen(false);
  activePhotoLabelCatalogSelection = '';
  draftModalPhotoLabels = [];
  setPhotoLabelNewFormOpen(false);
  resetPhotoLabelNewForm();
  closeSubModalElement(photoLabelModal, {
    onClosed: () => {
      if (photoLabelSaveStatus) {
        photoLabelSaveStatus.textContent = '';
      }
    },
  });
}

async function openPhotoLabelModal() {
  if (!currentModalPhoto || !photoLabelModal) {
    return;
  }

  openSubModalElement(photoLabelModal);
  setPhotoLabelCatalogMenuOpen(false);
  activePhotoLabelCatalogSelection = '';

  if (photoLabelSaveStatus) {
    photoLabelSaveStatus.textContent = 'ラベルを読み込み中...';
  }

  setPhotoLabelNewFormOpen(false);
  resetPhotoLabelNewForm();

  await Promise.all([
    loadPhotoLabelCatalog(),
    loadModalPhotoLabels(currentModalPhoto),
  ]);

  draftModalPhotoLabels = sortPhotoLabels(currentModalPhotoLabels);
  renderPhotoLabelEditorSelectedList();
  renderPhotoLabelCatalogOptions();

  if (photoLabelSaveStatus) {
    photoLabelSaveStatus.textContent = '';
  }
}

function addSelectedPhotoLabel() {
  if (!activePhotoLabelCatalogSelection) {
    return;
  }

  const normalizedName = activePhotoLabelCatalogSelection;

  if (!normalizedName) {
    return;
  }

  const selectedLabel = photoLabelCatalog.find(
    (label) => label.normalizedName === normalizedName
  );

  if (!selectedLabel) {
    return;
  }

  draftModalPhotoLabels = sortPhotoLabels([
    ...draftModalPhotoLabels,
    selectedLabel,
  ]);

  renderPhotoLabelEditorSelectedList();
  renderPhotoLabelCatalogOptions();
  activePhotoLabelCatalogSelection = '';
  setPhotoLabelCatalogMenuOpen(false);
}

function addNewPhotoLabelDraft() {
  const nextName = photoLabelNewNameInput?.value?.trim();
  const nextColor = photoLabelNewColorInput?.value || '#6D5EF6';

  if (!nextName) {
    if (photoLabelSaveStatus) {
      photoLabelSaveStatus.textContent = 'ラベル名を入力してください';
    }
    return;
  }

  const normalizedLabel = normalizePhotoLabelEntry({
    name: nextName,
    normalizedName: nextName.normalize('NFC').toLowerCase(),
    colorHex: nextColor,
  });

  if (!normalizedLabel) {
    if (photoLabelSaveStatus) {
      photoLabelSaveStatus.textContent = 'ラベル名を確認してください';
    }
    return;
  }

  draftModalPhotoLabels = sortPhotoLabels([
    ...draftModalPhotoLabels.filter(
      (label) => label.normalizedName !== normalizedLabel.normalizedName
    ),
    normalizedLabel,
  ]);

  renderPhotoLabelEditorSelectedList();
  renderPhotoLabelCatalogOptions();
  setPhotoLabelNewFormOpen(false);
  resetPhotoLabelNewForm();

  if (photoLabelSaveStatus) {
    photoLabelSaveStatus.textContent = '';
  }
}

async function savePhotoLabels() {
  if (!currentModalPhoto) {
    return;
  }

  if (photoLabelSaveStatus) {
    photoLabelSaveStatus.textContent = '保存中...';
  }

  const result = await window.electronAPI.replacePhotoLabels(
    currentModalPhoto.id,
    draftModalPhotoLabels
  );

  if (!result?.ok) {
    if (photoLabelSaveStatus) {
      photoLabelSaveStatus.textContent =
        result?.message || 'ラベルの保存に失敗しました';
    }
    return;
  }

  currentModalPhotoLabels = sortPhotoLabels(result.labels);
  draftModalPhotoLabels = sortPhotoLabels(result.labels);
  photoLabelCatalog = sortPhotoLabels(result.catalog);
  currentModalPhoto = {
    ...currentModalPhoto,
    photoLabels: currentModalPhotoLabels,
  };
  updatePhotoInCurrentCollections(currentModalPhoto);
  syncFavoriteFilterUi();

  if (isAnyPhotoFilterActive() && !photoMatchesCurrentFilters(currentModalPhoto)) {
    await refreshCurrentMonthWithFilterAnimation();
  } else {
    replaceRenderedPhotoCard(currentModalPhoto);
  }
  renderModalPhotoLabels();
  renderPhotoLabelEditorSelectedList();
  renderPhotoLabelCatalogOptions();
  closePhotoLabelModal();
  showToast('ラベルを保存しました');
}

function getCurrentModalPhotoIndex() {
  if (!currentModalPhoto?.id) {
    return -1;
  }

  return currentPhotos.findIndex((photo) => photo.id === currentModalPhoto.id);
}

function updateImageModalNavigationState() {
  const currentIndex = getCurrentModalPhotoIndex();
  const hasPrev = currentIndex > 0;
  const hasNext =
    currentIndex >= 0 && currentIndex < Math.max(currentPhotos.length - 1, 0);

  if (imageModalPrevButton) {
    imageModalPrevButton.disabled = !hasPrev;
    imageModalPrevButton.classList.toggle('is-disabled', !hasPrev);
  }

  if (imageModalNextButton) {
    imageModalNextButton.disabled = !hasNext;
    imageModalNextButton.classList.toggle('is-disabled', !hasNext);
  }
}

function syncImageModalPhotoLayout(item) {
  if (!imageModal) {
    return;
  }

  let orientation = null;
  const renderedWidth = Number(modalImage?.naturalWidth);
  const renderedHeight = Number(modalImage?.naturalHeight);

  if (
    modalImage &&
    modalImage.complete &&
    Number.isFinite(renderedWidth) &&
    Number.isFinite(renderedHeight) &&
    renderedWidth > 0 &&
    renderedHeight > 0
  ) {
    if (renderedWidth === renderedHeight) {
      orientation = 'square';
    } else {
      orientation = renderedWidth > renderedHeight ? 'landscape' : 'portrait';
    }
  } else {
    orientation = getPhotoOrientationTier(item);
  }

  const orientationClasses = [
    'is-photo-portrait',
    'is-photo-landscape',
    'is-photo-square',
  ];

  imageModal.classList.remove(...orientationClasses);

  if (orientation === 'portrait') {
    imageModal.classList.add('is-photo-portrait');
    return;
  }

  if (orientation === 'landscape') {
    imageModal.classList.add('is-photo-landscape');
    return;
  }

  if (orientation === 'square') {
    imageModal.classList.add('is-photo-square');
  }
}

async function loadModalWorldMetadata(item) {
  const requestId = ++modalWorldMetadataRequestId;

  if (modalWorldDescription) {
    modalWorldDescription.textContent = '未取得';
  }

  renderModalWorldTags([]);

  if (!item?.worldId) {
    return;
  }

  try {
    const metadata = await window.electronAPI.getWorldMetadata(item.worldId);

    if (
      requestId !== modalWorldMetadataRequestId ||
      !currentModalPhoto ||
      currentModalPhoto.id !== item.id
    ) {
      return;
    }

    if (modalWorldDescription) {
      modalWorldDescription.textContent =
        metadata?.worldDescription?.trim() || '未取得';
    }

    renderModalWorldTags(metadata?.worldTags);
  } catch {
    if (
      requestId !== modalWorldMetadataRequestId ||
      !currentModalPhoto ||
      currentModalPhoto.id !== item.id
    ) {
      return;
    }

    if (modalWorldDescription) {
      modalWorldDescription.textContent = '未取得';
    }

    renderModalWorldTags([]);
  }
}

function playImageModalSwitchAnimation(direction) {
  if (!imageModalBody) {
    return;
  }

  const prefersReducedMotion = window.matchMedia?.(
    '(prefers-reduced-motion: reduce)'
  )?.matches;

  if (prefersReducedMotion) {
    return;
  }

  if (imageModalSwitchTimer) {
    clearTimeout(imageModalSwitchTimer);
    imageModalSwitchTimer = null;
  }

  imageModalBody.classList.remove('is-switching-prev', 'is-switching-next');
  void imageModalBody.offsetWidth;
  imageModalBody.classList.add(
    direction === 'prev' ? 'is-switching-prev' : 'is-switching-next'
  );

  imageModalSwitchTimer = setTimeout(() => {
    imageModalBody.classList.remove('is-switching-prev', 'is-switching-next');
    imageModalSwitchTimer = null;
  }, 360);
}

function showImageModalPhoto(item, { direction = null } = {}) {
  if (!item) {
    return;
  }

  const latestPhoto = getLatestKnownPhotoById(item.id) || item;

  if (photoLabelModal && !photoLabelModal.classList.contains('hidden')) {
    closePhotoLabelModal();
  }

  currentModalPhoto = latestPhoto;
  modalImageRecoveryRequestId += 1;
  syncImageModalPhotoLayout(latestPhoto);
  populateModal(latestPhoto);
  void loadModalWorldMetadata(latestPhoto);
  void loadModalPhotoLabels(latestPhoto);

  if (imageModalInfo) {
    imageModalInfo.scrollTop = 0;
  }

  updateImageModalNavigationState();
  syncWorldMetadataSyncUi();

  if (direction) {
    playImageModalSwitchAnimation(direction);
  }
}

function stepImageModalPhoto(step) {
  const currentIndex = getCurrentModalPhotoIndex();

  if (currentIndex < 0) {
    return;
  }

  const nextPhoto = currentPhotos[currentIndex + step];

  if (!nextPhoto) {
    return;
  }

  showImageModalPhoto(nextPhoto, {
    direction: step < 0 ? 'prev' : 'next',
  });
}

function setImageModalScrollLock(isLocked) {
  document.documentElement.classList.toggle('image-modal-open', isLocked);
  document.body.classList.toggle('image-modal-open', isLocked);
}

function handleImageModalWheel(event) {
  if (
    !imageModal ||
    imageModal.classList.contains('hidden') ||
    !imageModalInfo
  ) {
    return;
  }

  const maxScrollTop =
    imageModalInfo.scrollHeight - imageModalInfo.clientHeight;

  event.preventDefault();

  if (maxScrollTop <= 0) {
    return;
  }

  imageModalInfo.scrollTop = Math.max(
    0,
    Math.min(maxScrollTop, imageModalInfo.scrollTop + event.deltaY)
  );
}

function triggerModalShellRestoreAnimation() {
  const prefersReducedMotion = window.matchMedia?.(
    '(prefers-reduced-motion: reduce)'
  )?.matches;
  const targets = [topStickyShell, sidebar].filter(Boolean);

  if (targets.length === 0 || prefersReducedMotion) {
    return;
  }

  if (modalShellRestoreTimer) {
    clearTimeout(modalShellRestoreTimer);
    modalShellRestoreTimer = null;
  }

  targets.forEach((element) => {
    element.classList.remove('modal-shell-restore');
    void element.offsetWidth;
    element.classList.add('modal-shell-restore');
  });

  modalShellRestoreTimer = setTimeout(() => {
    targets.forEach((element) => {
      element.classList.remove('modal-shell-restore');
    });
    modalShellRestoreTimer = null;
  }, 520);
}

function openImageModal(item) {
  if (!imageModal) {
    return;
  }

  if (imageModalAnimationTimer) {
    clearTimeout(imageModalAnimationTimer);
    imageModalAnimationTimer = null;
  }

  showImageModalPhoto(item);
  setImageModalScrollLock(true);
  imageModal.classList.remove('hidden');
  imageModal.classList.remove('is-closing');
  void imageModal.offsetWidth;
  imageModal.classList.add('is-open');
}

function closeImageModal() {
  if (!imageModal || imageModal.classList.contains('hidden')) {
    return;
  }

  closePhotoLabelModal();

  const prefersReducedMotion = window.matchMedia?.(
    '(prefers-reduced-motion: reduce)'
  )?.matches;

  if (imageModalAnimationTimer) {
    clearTimeout(imageModalAnimationTimer);
    imageModalAnimationTimer = null;
  }

  imageModal.classList.remove('is-open');
  imageModal.classList.add('is-closing');
  modalWorldMetadataRequestId += 1;
  modalPhotoLabelsRequestId += 1;
  modalImageRecoveryRequestId += 1;

  const finalizeImageModalClose = () => {
    if (imageModalSwitchTimer) {
      clearTimeout(imageModalSwitchTimer);
      imageModalSwitchTimer = null;
    }

    imageModalBody?.classList.remove('is-switching-prev', 'is-switching-next');
    imageModal.classList.remove('is-closing');
    imageModal.classList.add('hidden');
    modalImage.src = '';
    currentModalPhoto = null;
    currentModalPhotoLabels = [];
    worldNameSaveStatus.textContent = '';
    if (imageModalInfo) {
      imageModalInfo.scrollTop = 0;
    }
    updateImageModalNavigationState();
    imageModalAnimationTimer = null;

    requestAnimationFrame(() => {
      setImageModalScrollLock(false);
      triggerModalShellRestoreAnimation();
    });
  };

  if (prefersReducedMotion) {
    finalizeImageModalClose();
    return;
  }

  imageModalAnimationTimer = setTimeout(
    finalizeImageModalClose,
    IMAGE_MODAL_ANIMATION_MS
  );
}

function openWorldNameEditModal() {
  if (!currentModalPhoto) {
    return;
  }

  modalWorldNameInput.value =
    currentModalPhoto.worldNameManual ||
    currentModalPhoto.rawWorldName ||
    currentModalPhoto.worldName ||
    '';

  if (modalWorldUrlInput) {
    modalWorldUrlInput.value = currentModalPhoto.worldUrl || '';
  }

  worldNameSaveStatus.textContent = '';
  syncWorldMetadataSyncUi();
  openSubModalElement(worldNameEditModal);
}

function closeWorldNameEditModal() {
  closeSubModalElement(worldNameEditModal, {
    onClosed: () => {
      worldNameSaveStatus.textContent = '';
    },
  });
}

function closeConfirmModal(result = false) {
  if (confirmModal) {
    closeSubModalElement(confirmModal);
  }

  const resolver = confirmModalResolver;
  confirmModalResolver = null;

  if (resolver) {
    resolver(result);
  }
}

function openConfirmModal({
  title = '確認',
  message = 'この操作を実行しますか？',
  confirmText = '実行する',
} = {}) {
  if (!confirmModal) {
    return Promise.resolve(false);
  }

  confirmModalTitle.textContent = title;
  confirmModalMessage.textContent = message;
  confirmModalConfirmButton.textContent = confirmText;

  openSubModalElement(confirmModal);

  return new Promise((resolve) => {
    confirmModalResolver = resolve;
  });
}

function renderTrackedFolderList() {
  if (!trackedFolderList) {
    return;
  }

  if (!Array.isArray(trackedFolders) || trackedFolders.length === 0) {
    trackedFolderList.innerHTML =
      '<p class="tracked-folder-empty">まだ登録されていません</p>';
    return;
  }

  trackedFolderList.innerHTML = trackedFolders
    .map(
      (folder) => `
        <div class="tracked-folder-item">
          <p class="tracked-folder-item-path">${escapeHtml(folder.folder_path || '')}</p>
          <button
            type="button"
            class="tracked-folder-remove-button"
            data-tracked-folder-path="${escapeHtml(folder.folder_path || '')}"
          >
            削除
          </button>
        </div>
      `
    )
    .join('');
}

// Settings modal maintenance buttons are enabled or disabled from the current
// sidebar state so destructive actions stay predictable during testing.
function syncSettingsMaintenanceUi() {
  if (deleteCurrentMonthRegistrationsButton) {
    const hasSelection =
      Number.isInteger(currentSelection?.year) &&
      Number.isInteger(currentSelection?.month) &&
      sidebarData.length > 0;

    deleteCurrentMonthRegistrationsButton.disabled = isImporting || !hasSelection;
    deleteCurrentMonthRegistrationsButton.setAttribute(
      'title',
      hasSelection
        ? `${currentSelection.year}年${currentSelection.month}月の登録を削除`
        : '月を選択すると利用できます'
    );
  }

  if (deleteAllRegistrationsButton) {
    const hasAnyRegistration = sidebarData.length > 0;
    deleteAllRegistrationsButton.disabled = isImporting || !hasAnyRegistration;
    deleteAllRegistrationsButton.setAttribute(
      'title',
      hasAnyRegistration ? 'すべての登録を削除' : '削除する登録がありません'
    );
  }

  if (clearThumbnailCacheButton) {
    const hasAnyRegistration = sidebarData.length > 0;
    clearThumbnailCacheButton.disabled = isImporting || !hasAnyRegistration;
    clearThumbnailCacheButton.setAttribute(
      'title',
      hasAnyRegistration
        ? '管理サムネイルとサムネイル参照を削除'
        : '削除するサムネイルがありません'
    );
  }

  if (resetDatabaseButton) {
    const hasAnyPersistedData =
      sidebarData.length > 0 || trackedFolders.length > 0;
    resetDatabaseButton.disabled = isImporting || !hasAnyPersistedData;
    resetDatabaseButton.setAttribute(
      'title',
      hasAnyPersistedData
        ? '登録・キャッシュ・更新対象フォルダを初期化'
        : '初期化するデータがありません'
    );
  }
}

function initializeModalCloseIcons() {
  [
    imageModalClose,
    worldNameEditClose,
    photoLabelClose,
    settingsModalClose,
    confirmModalClose,
  ].forEach((button) => {
    if (!button) {
      return;
    }

    button.setAttribute('aria-label', '閉じる');
    button.textContent = '';

    const icon = document.createElement('span');
    icon.className = 'material-symbols-outlined';
    icon.textContent = 'close';
    button.appendChild(icon);
  });
}

// Some sub-modals are created dynamically, so close-button handling is also
// delegated from the document to avoid missing late-bound buttons.
function handleDelegatedSubModalClose(event) {
  const closeButton = event.target.closest('.sub-modal-close');

  if (!closeButton) {
    return;
  }

  if (closeButton === photoLabelClose) {
    event.preventDefault();
    event.stopPropagation();
    closePhotoLabelModal();
    return;
  }

  if (closeButton === worldNameEditClose) {
    event.preventDefault();
    event.stopPropagation();
    closeWorldNameEditModal();
    return;
  }

  if (closeButton === settingsModalClose) {
    event.preventDefault();
    event.stopPropagation();
    closeSettingsModal();
    return;
  }

  if (closeButton === confirmModalClose) {
    event.preventDefault();
    event.stopPropagation();
    closeConfirmModal(false);
  }
}

function initializeImageModalUi() {
  if (modalImage && !modalImage.dataset.layoutBound) {
    modalImage.addEventListener('load', () => {
      if (currentModalPhoto) {
        syncImageModalPhotoLayout(currentModalPhoto);
      }
    });
    modalImage.addEventListener('error', async () => {
      const targetPhoto = currentModalPhoto;

      if (!targetPhoto?.id || !window.electronAPI.resolvePhotoAccess) {
        return;
      }

      const requestId = ++modalImageRecoveryRequestId;
      const result = await window.electronAPI.resolvePhotoAccess({
        photoId: targetPhoto.id,
        filePath: targetPhoto.filePath,
      });

      if (requestId !== modalImageRecoveryRequestId) {
        return;
      }

      if (result?.photo) {
        updatePhotoInCurrentCollections(result.photo);
        replaceRenderedPhotoCard(result.photo);
        showImageModalPhoto(
          currentPhotos.find((photo) => photo.id === result.photo.id) ||
            result.photo
        );
        showToast('画像パスを再検出して更新しました');
        return;
      }

      if (!result?.ok) {
        showToast(
          `画像を表示できませんでした: ${result?.message || '不明なエラー'}`
        );
      }
    });
    modalImage.dataset.layoutBound = 'true';
  }

  if (imageModalImageWrap && !imageModalPrevButton && !imageModalNextButton) {
    imageModalPrevButton = document.createElement('button');
    imageModalPrevButton.type = 'button';
    imageModalPrevButton.className = 'image-modal-nav-button is-prev';
    imageModalPrevButton.setAttribute('aria-label', '前の画像');
    imageModalPrevButton.innerHTML =
      '<span class="material-symbols-outlined">chevron_left</span>';

    imageModalNextButton = document.createElement('button');
    imageModalNextButton.type = 'button';
    imageModalNextButton.className = 'image-modal-nav-button is-next';
    imageModalNextButton.setAttribute('aria-label', '次の画像');
    imageModalNextButton.innerHTML =
      '<span class="material-symbols-outlined">chevron_right</span>';

    imageModal?.appendChild(imageModalPrevButton);
    imageModal?.appendChild(imageModalNextButton);

    imageModalPrevButton.addEventListener('click', () => {
      stepImageModalPhoto(-1);
    });

    imageModalNextButton.addEventListener('click', () => {
      stepImageModalPhoto(1);
    });
  }

  if (modalWorldHero && modalWorldLabel && !modalResolutionHeroBadge) {
    const badgeRow = document.createElement('div');
    badgeRow.className = 'modal-hero-badge-row';
    modalWorldHero.insertBefore(badgeRow, modalWorldLabel);
    badgeRow.appendChild(modalWorldLabel);

    modalResolutionHeroBadge = document.createElement('span');
    modalResolutionHeroBadge.className =
      'modal-world-label modal-resolution-hero-badge is-hidden';
    badgeRow.appendChild(modalResolutionHeroBadge);
  }

  if (modalWorldHero && modalWorldLink && !modalTakenAtHero) {
    modalTakenAtHero = document.createElement('p');
    modalTakenAtHero.className = 'modal-hero-taken-at is-hidden';
    modalWorldHero.insertBefore(modalTakenAtHero, modalFileName);
  }

  if (modalOpenWorldButton) {
    modalOpenWorldButton.innerHTML = `
      <span class="primary-link-button-label">VRChatで開く</span>
      <span class="material-symbols-outlined primary-link-button-icon">open_in_new</span>
    `;
  }

  if (modalOpenOriginalButton) {
    modalOpenOriginalButton.textContent = '画像を開く';
  }

  if (modalOpenFolderButton) {
    modalOpenFolderButton.textContent = '保存先を開く';
  }

  if (openWorldNameEditButton) {
    openWorldNameEditButton.textContent = '編集';
  }

  if (rereadWorldNameButton) {
    rereadWorldNameButton.textContent = 'World情報を再読み込み';
  }

  if (
    modalDeletePhotoButton &&
    worldNameEditorActions &&
    !worldNameEditorActions.contains(modalDeletePhotoButton)
  ) {
    modalDeletePhotoButton.textContent = 'この登録を削除';
    worldNameEditorActions.appendChild(modalDeletePhotoButton);
  }

  modalDangerActions?.remove();
  updateImageModalNavigationState();
}

function initializeWorldNameEditUi() {
  const titleElement = worldNameEditModal?.querySelector('.sub-modal-body h3');
  const inputLabel = worldNameEditModal?.querySelector(
    'label[for="modal-world-name-input"]'
  );
  const urlLabel = worldNameEditModal?.querySelector(
    'label[for="modal-world-url-input"]'
  );

  if (titleElement) {
    titleElement.textContent = 'ワールド名を編集';
  }

  if (
    modalWorldNameInput &&
    saveWorldNameButton &&
    modalWorldNameInput.parentElement &&
    !modalWorldNameInput.parentElement.classList.contains('world-name-input-row')
  ) {
    const inputRow = document.createElement('div');
    inputRow.className = 'world-name-input-row';
    modalWorldNameInput.parentElement.insertBefore(inputRow, modalWorldNameInput);
    inputRow.appendChild(modalWorldNameInput);
    inputRow.appendChild(saveWorldNameButton);
  }

  if (inputLabel) {
    inputLabel.textContent = 'ワールド名';
  }

  if (urlLabel) {
    urlLabel.textContent = 'World URL';
  }

  if (saveWorldNameButton) {
    saveWorldNameButton.textContent = '保存';
  }

  if (clearWorldNameButton) {
    clearWorldNameButton.textContent = '手動設定を解除';
  }

  if (modalDeletePhotoButton && worldNameEditorActions) {
    modalDeletePhotoButton.classList.add('world-name-delete-button');

    if (!worldNameEditorActions.contains(modalDeletePhotoButton)) {
      worldNameEditorActions.appendChild(modalDeletePhotoButton);
    }
  }
}

// Photo labels are configured from the image modal, but the editor itself is
// mounted lazily here to keep the static HTML smaller and easier to recover.
function initializePhotoLabelUi() {
  if (modalWorldTags && !modalPhotoLabelsBlock) {
    const referenceBlock = modalWorldTags.closest('.modal-world-meta-block');

    if (referenceBlock?.parentElement) {
      modalPhotoLabelsBlock = document.createElement(
        'div'
      );
      modalPhotoLabelsBlock.className =
        'modal-world-meta-block modal-photo-label-block';

      const header = document.createElement('div');
      header.className = 'photo-label-block-header';

      const title = document.createElement('p');
      title.className = 'modal-world-meta-title';
      title.textContent = 'ラベル';
      header.appendChild(title);

      openPhotoLabelEditorButton = document.createElement('button');
      openPhotoLabelEditorButton.type = 'button';
      openPhotoLabelEditorButton.className = 'small-action-button';
      openPhotoLabelEditorButton.textContent = '編集';
      header.appendChild(openPhotoLabelEditorButton);

      modalPhotoLabelsList = document.createElement('div');
      modalPhotoLabelsList.className = 'modal-photo-labels';

      modalPhotoLabelsBlock.appendChild(header);
      modalPhotoLabelsBlock.appendChild(modalPhotoLabelsList);
      referenceBlock.insertAdjacentElement('afterend', modalPhotoLabelsBlock);
      renderModalPhotoLabels();
    }
  }

  if (!photoLabelModal) {
    photoLabelModal = document.createElement('div');
    photoLabelModal.id = 'photo-label-modal';
    photoLabelModal.className = 'sub-modal hidden';

    photoLabelBackdrop = document.createElement('div');
    photoLabelBackdrop.className = 'sub-modal-backdrop';
    photoLabelModal.appendChild(photoLabelBackdrop);

    const content = document.createElement('div');
    content.className = 'sub-modal-content photo-label-modal-content';
    photoLabelModal.appendChild(content);

    photoLabelClose = document.createElement('button');
    photoLabelClose.type = 'button';
    photoLabelClose.className = 'sub-modal-close';
    photoLabelClose.setAttribute('aria-label', '閉じる');
    photoLabelClose.innerHTML =
      '<span class="material-symbols-outlined">close</span>';
    content.appendChild(photoLabelClose);

    const body = document.createElement('div');
    body.className = 'sub-modal-body photo-label-modal-body';
    content.appendChild(body);

    const title = document.createElement('h3');
    title.textContent = 'ラベルを設定';
    body.appendChild(title);

    const description = document.createElement('p');
    description.className = 'sub-modal-description';
    description.textContent =
      '既存ラベルを再利用したり、新しいラベルを色付きで追加して写真ごとに設定できます。';
    body.appendChild(description);

    const selectedTitle = document.createElement('p');
    selectedTitle.className = 'photo-label-editor-section-title';
    selectedTitle.textContent = '現在のラベル';
    body.appendChild(selectedTitle);

    photoLabelSelectedList = document.createElement('div');
    photoLabelSelectedList.className = 'photo-label-selected-list';
    body.appendChild(photoLabelSelectedList);

    const pickerTitle = document.createElement('p');
    pickerTitle.className = 'photo-label-editor-section-title';
    pickerTitle.textContent = 'Choose Existing Label';
    body.appendChild(pickerTitle);

    const pickerRow = document.createElement('div');
    pickerRow.className = 'photo-label-picker-row';
    body.appendChild(pickerRow);

    photoLabelCatalogDropdown = document.createElement('div');
    photoLabelCatalogDropdown.className =
      'header-dropdown photo-label-catalog-dropdown';
    pickerRow.appendChild(photoLabelCatalogDropdown);

    photoLabelCatalogButton = document.createElement('button');
    photoLabelCatalogButton.type = 'button';
    photoLabelCatalogButton.className =
      'header-filter-button orientation-filter-button photo-label-catalog-button';
    photoLabelCatalogButton.setAttribute('aria-haspopup', 'menu');
    photoLabelCatalogButton.setAttribute('aria-expanded', 'false');
    photoLabelCatalogDropdown.appendChild(photoLabelCatalogButton);

    photoLabelCatalogMenu = document.createElement('div');
    photoLabelCatalogMenu.className =
      'header-dropdown-menu photo-label-catalog-menu';
    photoLabelCatalogMenu.setAttribute('role', 'menu');
    photoLabelCatalogMenu.hidden = true;
    photoLabelCatalogDropdown.appendChild(photoLabelCatalogMenu);

    const newTitle = document.createElement('p');
    newTitle.className = 'photo-label-editor-section-title';
    newTitle.textContent = 'Create New Label';
    body.appendChild(newTitle);

    photoLabelNewForm = document.createElement('div');
    photoLabelNewForm.className = 'photo-label-new-form';
    photoLabelNewForm.hidden = false;
    body.appendChild(photoLabelNewForm);

    const nameRow = document.createElement('div');
    nameRow.className = 'photo-label-new-name-row';
    photoLabelNewForm.appendChild(nameRow);

    photoLabelNewNameInput = document.createElement('input');
    photoLabelNewNameInput.type = 'text';
    photoLabelNewNameInput.className = 'photo-label-new-name';
    photoLabelNewNameInput.placeholder = 'New label name';
    nameRow.appendChild(photoLabelNewNameInput);

    photoLabelNewColorPreview = document.createElement('span');
    photoLabelNewColorPreview.className = 'photo-label-color-preview';
    nameRow.appendChild(photoLabelNewColorPreview);

    photoLabelCustomColorButton = document.createElement('button');
    photoLabelCustomColorButton.type = 'button';
    photoLabelCustomColorButton.className = 'photo-label-color-picker-button';
    photoLabelCustomColorButton.setAttribute('aria-label', '色を選択');
    photoLabelCustomColorButton.innerHTML =
      '<span class="material-symbols-outlined">palette</span>';
    nameRow.appendChild(photoLabelCustomColorButton);

    photoLabelNewColorInput = document.createElement('input');
    photoLabelNewColorInput.type = 'color';
    photoLabelNewColorInput.className = 'photo-label-hidden-color-input';
    photoLabelNewColorInput.value = PHOTO_LABEL_PRESET_COLORS[0];
    photoLabelNewForm.appendChild(photoLabelNewColorInput);

    photoLabelPresetList = document.createElement('div');
    photoLabelPresetList.className = 'photo-label-preset-list';
    photoLabelNewForm.appendChild(photoLabelPresetList);

    photoLabelAddNewButton = document.createElement('button');
    photoLabelAddNewButton.type = 'button';
    photoLabelAddNewButton.className = 'small-action-button';
    photoLabelAddNewButton.textContent = 'この内容で追加';
    photoLabelNewForm.appendChild(photoLabelAddNewButton);

    photoLabelSaveStatus = document.createElement('p');
    photoLabelSaveStatus.className =
      'world-name-save-status photo-label-save-status';
    body.appendChild(photoLabelSaveStatus);

    const actions = document.createElement('div');
    actions.className = 'photo-label-editor-actions';
    body.appendChild(actions);

    photoLabelSaveButton = document.createElement('button');
    photoLabelSaveButton.type = 'button';
    photoLabelSaveButton.className = 'primary-link-button photo-label-save-button';
    photoLabelSaveButton.textContent = '保存';
    actions.appendChild(photoLabelSaveButton);

    document.body.appendChild(photoLabelModal);
  }

  setPhotoLabelNewFormOpen(true);
  resetPhotoLabelNewForm();
  renderPhotoLabelPresetButtons();

  openPhotoLabelEditorButton?.addEventListener('click', () => {
    void openPhotoLabelModal();
  });

  bindSubModalCloseTriggers(
    photoLabelBackdrop,
    photoLabelClose,
    closePhotoLabelModal
  );

  photoLabelCatalogButton?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();

    if (photoLabelCatalogButton?.disabled) {
      return;
    }

    setPhotoLabelCatalogMenuOpen(!isPhotoLabelCatalogMenuOpen);
  });

  photoLabelCatalogMenu?.addEventListener('click', (event) => {
    event.stopPropagation();
  });

  photoLabelCustomColorButton?.addEventListener('click', () => {
    photoLabelNewColorInput?.click();
  });

  photoLabelNewColorInput?.addEventListener('input', () => {
    setPhotoLabelDraftColor(photoLabelNewColorInput?.value);
  });

  photoLabelAddNewButton?.addEventListener('click', () => {
    addNewPhotoLabelDraft();
  });

  photoLabelNewNameInput?.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      addNewPhotoLabelDraft();
    }
  });

  photoLabelSaveButton?.addEventListener('click', async () => {
    await savePhotoLabels();
  });
}

function syncTrackedFolderAccordionState() {
  if (!trackedFolderAccordionButton || !trackedFolderAccordionPanel) {
    return;
  }

  trackedFolderAccordionButton.setAttribute(
    'aria-expanded',
    isTrackedFolderAccordionOpen ? 'true' : 'false'
  );
  trackedFolderAccordionPanel.classList.toggle(
    'is-open',
    isTrackedFolderAccordionOpen
  );
}

function setTrackedFolderAccordionOpen(isOpen) {
  isTrackedFolderAccordionOpen = Boolean(isOpen);
  syncTrackedFolderAccordionState();
}

function initializeTrackedFolderAccordion() {
  if (
    trackedFolderAccordionButton ||
    !settingsSectionHeader ||
    !trackedFolderList
  ) {
    return;
  }

  const copyContainer = settingsSectionHeader.querySelector(':scope > div');

  if (!copyContainer) {
    return;
  }

  trackedFolderAccordionButton = document.createElement('button');
  trackedFolderAccordionButton.id = 'tracked-folder-accordion-btn';
  trackedFolderAccordionButton.type = 'button';
  trackedFolderAccordionButton.className = 'settings-accordion-button';
  trackedFolderAccordionButton.setAttribute('aria-controls', 'tracked-folder-accordion-panel');

  copyContainer.classList.add('settings-accordion-copy');
  settingsSectionHeader.insertBefore(trackedFolderAccordionButton, copyContainer);
  trackedFolderAccordionButton.appendChild(copyContainer);

  const chevron = document.createElement('span');
  chevron.className = 'material-symbols-outlined settings-accordion-chevron';
  chevron.textContent = 'expand_more';
  trackedFolderAccordionButton.appendChild(chevron);

  trackedFolderAccordionPanel = document.createElement('div');
  trackedFolderAccordionPanel.id = 'tracked-folder-accordion-panel';
  trackedFolderAccordionPanel.className = 'settings-accordion-panel';

  const panelInner = document.createElement('div');
  panelInner.className = 'settings-accordion-panel-inner';

  trackedFolderList.parentElement.insertBefore(
    trackedFolderAccordionPanel,
    trackedFolderList
  );
  trackedFolderAccordionPanel.appendChild(panelInner);
  panelInner.appendChild(trackedFolderList);

  trackedFolderAccordionButton.addEventListener('click', () => {
    setTrackedFolderAccordionOpen(!isTrackedFolderAccordionOpen);
  });

  syncTrackedFolderAccordionState();
}

async function loadTrackedFoldersForSettings() {
  trackedFolders = await window.electronAPI.getTrackedFolders();
  renderTrackedFolderList();
}

function initializeTopToolbarLayout() {
  if (refreshTrackedFoldersButton && pageHeaderActions && settingsButton) {
    refreshTrackedFoldersButton.classList.add('theme-toggle-btn');
    refreshTrackedFoldersButton.innerHTML =
      '<span class="material-symbols-outlined">sync</span>';
    refreshTrackedFoldersButton.setAttribute('aria-label', '更新');
    refreshTrackedFoldersButton.setAttribute('title', '更新');
    pageHeaderActions.insertBefore(refreshTrackedFoldersButton, settingsButton);
  }

  if (worldNameFilterDropdown && toolbar && toolbarRight) {
    let toolbarLeftGroup = toolbar.querySelector('.toolbar-left-group');

    if (!toolbarLeftGroup) {
      toolbarLeftGroup = document.createElement('div');
      toolbarLeftGroup.className = 'toolbar-left-group';
      toolbar.insertBefore(toolbarLeftGroup, toolbarRight);
    }

    worldNameFilterDropdown.classList.add('toolbar-world-filter');
    worldNameFilterDropdown.classList.add('is-static-toolbar-filter');
    toolbarLeftGroup.appendChild(worldNameFilterDropdown);
  }

  if (regenerateThumbnailsButton && settingsSectionHeader?.parentElement) {
    let utilityActions = settingsModalBody?.querySelector('.settings-utility-actions');
    const utilityAnchor = trackedFolderAccordionPanel || trackedFolderList;

    if (!utilityActions) {
      utilityActions = document.createElement('div');
      utilityActions.className = 'settings-utility-actions';
      settingsSectionHeader.parentElement.insertBefore(
        utilityActions,
        utilityAnchor || null
      );
    } else if (
      utilityAnchor &&
      utilityActions.nextElementSibling !== utilityAnchor
    ) {
      settingsSectionHeader.parentElement.insertBefore(utilityActions, utilityAnchor);
    }

    if (!regenerateThumbnailMonthSelect) {
      const monthSelect = document.createElement('select');
      monthSelect.className = 'settings-month-select';
      monthSelect.setAttribute('aria-label', 'サムネイル再生成の対象月');
      utilityActions.appendChild(monthSelect);
      regenerateThumbnailMonthSelect = monthSelect;
    }

    if (!regenerateThumbnailMonthDropdown) {
      const monthDropdown = document.createElement('div');
      monthDropdown.className = 'header-dropdown settings-month-dropdown';

      const monthButton = document.createElement('button');
      monthButton.type = 'button';
      monthButton.className =
        'header-filter-button settings-month-dropdown-button';
      monthButton.setAttribute('aria-haspopup', 'menu');
      monthButton.setAttribute('aria-expanded', 'false');
      monthButton.setAttribute('aria-label', 'サムネイル再生成の対象月');

      const monthLabel = document.createElement('span');
      monthLabel.className = 'settings-month-dropdown-label';
      monthLabel.textContent = '再生成する月を選択';
      monthButton.appendChild(monthLabel);

      const chevron = document.createElement('span');
      chevron.className =
        'material-symbols-outlined orientation-filter-chevron';
      chevron.textContent = 'expand_more';
      monthButton.appendChild(chevron);

      const monthMenu = document.createElement('div');
      monthMenu.className = 'header-dropdown-menu settings-month-dropdown-menu';
      monthMenu.hidden = true;
      monthMenu.setAttribute('role', 'menu');

      monthDropdown.appendChild(monthButton);
      monthDropdown.appendChild(monthMenu);
      utilityActions.appendChild(monthDropdown);

      regenerateThumbnailMonthDropdown = monthDropdown;
      regenerateThumbnailMonthButton = monthButton;
      regenerateThumbnailMonthLabel = monthLabel;
      regenerateThumbnailMonthMenu = monthMenu;

      monthButton.addEventListener('click', (event) => {
        event.stopPropagation();
        setRegenerateThumbnailMonthMenuOpen(!isRegenerateThumbnailMonthMenuOpen);
      });

      monthMenu.addEventListener('click', (event) => {
        const target = event.target.closest('[data-regenerate-thumbnail-month]');

        if (!target || !regenerateThumbnailMonthSelect) {
          return;
        }

        event.stopPropagation();
        regenerateThumbnailMonthValue =
          target.dataset.regenerateThumbnailMonth || '';
        regenerateThumbnailMonthSelect.value = regenerateThumbnailMonthValue;
        syncRegenerateThumbnailMonthDropdownFromSelect();
        closeRegenerateThumbnailMonthMenu();
      });
    }

    regenerateThumbnailsButton.classList.remove('secondary-toolbar-button');
    regenerateThumbnailsButton.classList.add('small-action-button');
    utilityActions.appendChild(regenerateThumbnailsButton);
    renderRegenerateThumbnailMonthOptions();
  }
}

function handleSettingsModalWheel(event) {
  if (
    !settingsModal ||
    settingsModal.classList.contains('hidden') ||
    !settingsModalBody
  ) {
    return;
  }

  const maxScrollTop =
    settingsModalBody.scrollHeight - settingsModalBody.clientHeight;

  if (maxScrollTop <= 0) {
    event.preventDefault();
    return;
  }

  if (!settingsModalBody.contains(event.target)) {
    event.preventDefault();
    settingsModalBody.scrollTop = Math.max(
      0,
      Math.min(maxScrollTop, settingsModalBody.scrollTop + event.deltaY)
    );
    return;
  }

  const isScrollingUp = event.deltaY < 0;
  const isScrollingDown = event.deltaY > 0;
  const isAtTop = settingsModalBody.scrollTop <= 0;
  const isAtBottom = settingsModalBody.scrollTop >= maxScrollTop - 1;

  if ((isScrollingUp && isAtTop) || (isScrollingDown && isAtBottom)) {
    event.preventDefault();
  }
}

async function openSettingsModal() {
  if (!settingsModal) {
    return;
  }

  initializeTrackedFolderAccordion();
  await loadTrackedFoldersForSettings();
  syncTrackedFolderAccordionState();
  syncSelectionDependentSettingsUi();

  if (settingsModalBody) {
    settingsModalBody.scrollTop = 0;
  }

  openSubModalElement(settingsModal);
}

function closeSettingsModal() {
  closeSubModalElement(settingsModal, {
    onClosed: () => {
      closeRegenerateThumbnailMonthMenu();
      if (settingsModalBody) {
        settingsModalBody.scrollTop = 0;
      }
    },
  });
}

function syncSelectionDependentSettingsUi() {
  renderRegenerateThumbnailMonthOptions();
  syncSettingsMaintenanceUi();
}

function resetCurrentMonthState() {
  currentSelection = null;
  currentPhotos = [];
  allCurrentMonthPhotos = [];
}

function clearMainContent() {
  resetCurrentMonthState();
  setAnimatedMonthLabelText('写真一覧', { animate: false });
  setAnimatedMonthCountText('0枚', { animate: false });
  monthGalleryList.innerHTML = '';
  monthGalleryEmpty.style.display = 'block';
  monthGalleryEmpty.textContent = getDefaultMonthGalleryEmptyMessage();
  resetMonthGalleryRenderState();
  clearSelectionState();
  syncFavoriteFilterUi();
  syncSelectionDependentSettingsUi();
}

function createPhotoCard(item) {
  const card = document.createElement('div');
  card.className = 'photo-card';
  card.dataset.photoId = String(item.id);
  card.draggable = false;

  if (isSelectionMode) {
    card.classList.add('selection-mode');
  }

  if (selectedPhotoIds.has(item.id)) {
    card.classList.add('is-selected');
  }

  const selectionButton = document.createElement('button');
  selectionButton.className = 'photo-card-selection-btn';
  selectionButton.type = 'button';
  selectionButton.setAttribute('aria-label', '選択を切り替え');
  selectionButton.setAttribute('title', '選択を切り替え');

  const selectionIcon = document.createElement('span');
  selectionIcon.className = 'material-symbols-outlined';
  selectionIcon.textContent = 'check';

  selectionButton.appendChild(selectionIcon);

  if (selectedPhotoIds.has(item.id)) {
    selectionButton.classList.add('is-selected');
  }

  selectionButton.addEventListener('click', (event) => {
    event.stopPropagation();

    if (!isSelectionMode) {
      return;
    }

    togglePhotoSelection(item.id);
  });

  const favoriteButton = document.createElement('button');
  favoriteButton.className = 'photo-card-favorite-btn';
  favoriteButton.type = 'button';
  favoriteButton.setAttribute('aria-label', 'お気に入りを切り替え');
  favoriteButton.setAttribute('title', 'お気に入りを切り替え');

  const favoriteIcon = document.createElement('span');
  favoriteIcon.className = 'material-symbols-outlined';
  favoriteIcon.textContent = 'star';

  favoriteButton.appendChild(favoriteIcon);
  syncFavoriteButtonState(favoriteButton, item.isFavorite);

  favoriteButton.addEventListener('click', async (event) => {
    event.stopPropagation();
    const latestPhoto = getLatestKnownPhotoById(item.id) || item;
    await toggleFavorite(item.id, !latestPhoto.isFavorite);
  });

  let visual;

  if (item.hasThumbnail && item.thumbnailUrl) {
    const image = document.createElement('img');
    image.className = 'photo-card-image';
    image.src = item.thumbnailUrl;
    image.alt = item.fileName || 'photo';
    image.loading = 'lazy';
    image.draggable = false;
    visual = image;
  } else {
    const placeholder = document.createElement('div');
    placeholder.className = 'photo-card-image photo-card-image-placeholder';
    placeholder.draggable = false;

    const icon = document.createElement('span');
    icon.className = 'material-symbols-outlined photo-card-placeholder-icon';
    icon.textContent = 'image';

    const text = document.createElement('span');
    text.className = 'photo-card-placeholder-text';
    text.textContent = 'サムネイル未生成';

    placeholder.appendChild(icon);
    placeholder.appendChild(text);
    visual = placeholder;
  }

  const { dateText, timeText } = splitTakenAtForCard(item.takenAt);

  const info = document.createElement('div');
  info.className = 'photo-card-info';

  const metaRow = document.createElement('div');
  metaRow.className = 'photo-card-meta-row';

  const date = document.createElement('p');
  date.className = 'photo-card-date';
  date.textContent = dateText;

  metaRow.appendChild(date);

  const timeInline = document.createElement('p');
  timeInline.className = 'photo-card-time-sub';
  timeInline.textContent = timeText || '譎ょ綾荳肴・';
  const time = timeInline;
  metaRow.appendChild(time);

  if (item.resolutionTier) {
    const resolutionBadge = document.createElement('span');
    resolutionBadge.className = 'photo-card-resolution-badge';
    resolutionBadge.textContent = item.resolutionTier;
    metaRow.appendChild(resolutionBadge);
  }

  /* time already appended in meta row */
  time.className = 'photo-card-time-sub';
  time.textContent = timeText || '時刻不明';

  const world = document.createElement('p');
  world.className = 'photo-card-world';
  world.textContent = item.worldName || 'ワールド名未取得';

  info.appendChild(metaRow);
  info.appendChild(world);

  if (Array.isArray(item.photoLabels) && item.photoLabels.length > 0) {
    const labels = document.createElement('div');
    labels.className = 'photo-card-labels';

    item.photoLabels.slice(0, 2).forEach((label) => {
      labels.appendChild(createPhotoCardLabelChip(label));
    });

    if (item.photoLabels.length > 2) {
      const overflowChip = document.createElement('span');
      overflowChip.className = 'photo-card-label-overflow';
      overflowChip.textContent = `+${item.photoLabels.length - 2}`;
      labels.appendChild(overflowChip);
    }

    info.appendChild(labels);
  }

  card.appendChild(selectionButton);
  card.appendChild(favoriteButton);
  card.appendChild(visual);
  card.appendChild(info);

  card.addEventListener('dragstart', (event) => {
    event.preventDefault();
  });

  card.addEventListener('click', () => {
    if (isSelectionMode) {
      togglePhotoSelection(item.id);
      return;
    }

    openImageModal(getLatestKnownPhotoById(item.id) || item);
  });

  return card;
}

function resetMonthGalleryRenderState() {
  if (monthGalleryLoadCheckTimer) {
    clearTimeout(monthGalleryLoadCheckTimer);
    monthGalleryLoadCheckTimer = null;
  }

  renderedPhotoCount = 0;
  renderedMonthGalleryKey = '';
  isAppendingMonthGalleryBatch = false;
  monthGalleryLoadCheckScheduled = false;
  renderedGalleryGroupMap = new Map();
}

function getMonthGalleryRenderKey() {
  if (!currentSelection) {
    return 'empty';
  }

  return [
    currentSelection.year,
    currentSelection.month,
    isFavoriteFilterOnly ? 'fav' : 'all',
    currentPhotos.length,
  ].join(':');
}

function calculateGalleryAvailableWidth() {
  const galleryWidth = monthGalleryList?.clientWidth || mainContent?.clientWidth || 0;
  return Math.max(
    galleryWidth - GALLERY_GROUP_HORIZONTAL_PADDING,
    GALLERY_CARD_MIN_WIDTH
  );
}

function calculateGalleryColumnCount() {
  const availableWidth = calculateGalleryAvailableWidth();

  return Math.max(
    1,
    Math.floor(
      (availableWidth + GALLERY_GRID_GAP) /
        (GALLERY_CARD_MIN_WIDTH + GALLERY_GRID_GAP)
    )
  );
}

function calculateGalleryCardHeightEstimate(columnCount) {
  const availableWidth = calculateGalleryAvailableWidth();
  const gapWidth = GALLERY_GRID_GAP * Math.max(0, columnCount - 1);
  const cardWidth = Math.max(
    GALLERY_CARD_MIN_WIDTH,
    Math.floor((availableWidth - gapWidth) / columnCount)
  );

  return cardWidth + GALLERY_CARD_EXTRA_HEIGHT;
}

function calculateInitialVisiblePhotoCount() {
  if (!monthGalleryList) {
    return currentPhotos.length;
  }

  const columns = calculateGalleryColumnCount();
  const rowHeight = Math.max(
    calculateGalleryCardHeightEstimate(columns),
    GALLERY_CARD_MIN_WIDTH
  );
  const rect = monthGalleryList.getBoundingClientRect();
  const visibleHeight = Math.max(window.innerHeight - Math.max(rect.top, 0), 320);
  const visibleRows = Math.max(1, Math.ceil(visibleHeight / rowHeight));
  const targetRows = Math.max(2, visibleRows + GALLERY_INITIAL_PREFETCH_ROWS);

  return Math.min(currentPhotos.length, Math.max(columns * targetRows, columns));
}

function calculateIncrementalVisiblePhotoCount() {
  const columns = calculateGalleryColumnCount();
  return Math.min(
    currentPhotos.length,
    renderedPhotoCount + columns * GALLERY_INCREMENT_ROWS
  );
}

function buildGalleryGroupSection(groupDate) {
  const section = document.createElement('section');
  section.className = 'gallery-group';

  const dayBox = document.createElement('div');
  dayBox.className = 'day-box';
  dayBox.textContent = groupDate;

  const grid = document.createElement('div');
  grid.className = 'gallery-grid';

  section.appendChild(dayBox);
  section.appendChild(grid);

  return { section, grid };
}

function appendMonthGalleryPhotoBatch(targetCount) {
  if (!monthGalleryList || targetCount <= renderedPhotoCount) {
    return;
  }

  const fragment = document.createDocumentFragment();

  for (let index = renderedPhotoCount; index < targetCount; index += 1) {
    const photo = currentPhotos[index];
    const groupDate = photo?.groupDate || '日付不明';

    let groupState = renderedGalleryGroupMap.get(groupDate);

    if (!groupState) {
      groupState = buildGalleryGroupSection(groupDate);
      renderedGalleryGroupMap.set(groupDate, groupState);
      fragment.appendChild(groupState.section);
    }

    groupState.grid.appendChild(createPhotoCard(photo));
  }

  renderedPhotoCount = targetCount;
  monthGalleryList.appendChild(fragment);
  syncSelectionUi();
}

function maybeLoadMoreMonthGalleryPhotos() {
  if (
    !currentSelection ||
    !monthGalleryList ||
    isAppendingMonthGalleryBatch ||
    renderedPhotoCount >= currentPhotos.length
  ) {
    return;
  }

  const rect = monthGalleryList.getBoundingClientRect();

  if (rect.bottom > window.innerHeight + GALLERY_LOAD_AHEAD_PX) {
    return;
  }

  isAppendingMonthGalleryBatch = true;

  try {
    appendMonthGalleryPhotoBatch(calculateIncrementalVisiblePhotoCount());
  } finally {
    isAppendingMonthGalleryBatch = false;
  }

  if (
    renderedPhotoCount < currentPhotos.length &&
    monthGalleryList.getBoundingClientRect().bottom <=
      window.innerHeight + GALLERY_LOAD_AHEAD_PX
  ) {
    scheduleMonthGalleryLoadCheck({ immediate: true });
  }
}

function scheduleMonthGalleryLoadCheck({ immediate = false } = {}) {
  if (monthGalleryLoadCheckScheduled) {
    if (!immediate || !monthGalleryLoadCheckTimer) {
      return;
    }

    clearTimeout(monthGalleryLoadCheckTimer);
    monthGalleryLoadCheckTimer = null;
  }

  monthGalleryLoadCheckScheduled = true;
  const delay = immediate ? 0 : GALLERY_LOAD_CHECK_THROTTLE_MS;

  monthGalleryLoadCheckTimer = setTimeout(() => {
    monthGalleryLoadCheckTimer = null;

    requestAnimationFrame(() => {
      monthGalleryLoadCheckScheduled = false;
      maybeLoadMoreMonthGalleryPhotos();
    });
  }, delay);
}

function initializeProgressiveMonthGalleryLoading() {
  window.addEventListener('scroll', scheduleMonthGalleryLoadCheck, true);
  window.addEventListener('resize', () => {
    scheduleMonthGalleryLoadCheck({ immediate: true });
  });
}

function initializeScrollToTopAnimationInterrupts() {
  const interruptScrollToTopAnimation = () => {
    stopScrollToTopAnimation();
  };

  window.addEventListener('wheel', interruptScrollToTopAnimation, {
    passive: true,
    capture: true,
  });
  window.addEventListener('touchstart', interruptScrollToTopAnimation, {
    passive: true,
    capture: true,
  });
  window.addEventListener('pointerdown', interruptScrollToTopAnimation, {
    passive: true,
    capture: true,
  });

  document.addEventListener(
    'keydown',
    (event) => {
      if (
        event.key === 'ArrowUp' ||
        event.key === 'ArrowDown' ||
        event.key === 'PageUp' ||
        event.key === 'PageDown' ||
        event.key === 'Home' ||
        event.key === 'End' ||
        event.key === ' ' ||
        event.key === 'Spacebar'
      ) {
        interruptScrollToTopAnimation();
      }
    },
    true
  );
}

function renderMonthGallery({ resetProgressive = false } = {}) {
  monthGalleryList.innerHTML = '';

  if (!currentSelection) {
    resetMonthGalleryRenderState();
    clearMainContent();
    monthGalleryEmpty.textContent = '表示する月を選択してください';
    return;
  }

  setAnimatedMonthLabelText(
    `${currentSelection.year}年${currentSelection.month}月`
  );
  // This is the canonical render path for month view state:
  // header count, empty state, and filter button text are all synchronized here.
  syncFavoriteFilterUi();

  if (currentPhotos.length === 0) {
    resetMonthGalleryRenderState();
    monthGalleryEmpty.style.display = 'block';
    monthGalleryEmpty.textContent =
      allCurrentMonthPhotos.length > 0 && isAnyPhotoFilterActive()
        ? buildFilteredEmptyMessage()
        : 'この月の写真はまだありません';

    return;
  }

  monthGalleryEmpty.style.display = 'none';

  const nextRenderKey = getMonthGalleryRenderKey();
  const previousRenderedPhotoCount =
    !resetProgressive && renderedMonthGalleryKey === nextRenderKey
      ? renderedPhotoCount
      : 0;

  renderedMonthGalleryKey = nextRenderKey;
  renderedPhotoCount = 0;
  renderedGalleryGroupMap = new Map();

  appendMonthGalleryPhotoBatch(
    Math.min(
      currentPhotos.length,
      Math.max(previousRenderedPhotoCount, calculateInitialVisiblePhotoCount())
    )
  );
  scheduleMonthGalleryLoadCheck({ immediate: true });
}

function resetMonthSwitchClasses() {
  const targets = [currentMonthLabel?.parentElement, monthGalleryList];

  for (const element of targets) {
    element?.classList.remove('month-switch-leave', 'month-switch-enter');
  }
}

function stopScrollToTopAnimation() {
  if (scrollToTopAnimationFrame) {
    cancelAnimationFrame(scrollToTopAnimationFrame);
    scrollToTopAnimationFrame = null;
  }
}

function scrollGalleryViewToTop({ animated = false } = {}) {
  const targets = [];
  const documentScrollTop =
    window.scrollY ||
    document.documentElement.scrollTop ||
    document.body.scrollTop ||
    0;

  if (mainContent && mainContent.scrollTop > 0) {
    targets.push({
      start: mainContent.scrollTop,
      apply(value) {
        mainContent.scrollTop = value;
      },
    });
  }

  if (
    monthGalleryList &&
    monthGalleryList.scrollTop > 0 &&
    monthGalleryList !== mainContent
  ) {
    targets.push({
      start: monthGalleryList.scrollTop,
      apply(value) {
        monthGalleryList.scrollTop = value;
      },
    });
  }

  if (documentScrollTop > 0) {
    targets.push({
      start: documentScrollTop,
      apply(value) {
        window.scrollTo(0, value);
      },
    });
  }

  if (targets.length === 0) {
    return;
  }

  stopScrollToTopAnimation();

  const prefersReducedMotion = window.matchMedia?.(
    '(prefers-reduced-motion: reduce)'
  )?.matches;

  if (!animated || prefersReducedMotion) {
    for (const target of targets) {
      target.apply(0);
    }
    return;
  }

  const maxDistance = Math.max(...targets.map((target) => target.start));
  const duration = Math.min(
    SCROLL_TO_TOP_MAX_DURATION_MS,
    Math.max(SCROLL_TO_TOP_MIN_DURATION_MS, maxDistance * 0.95)
  );
  const startTime = performance.now();
  const easeOut = (progress) => {
    if (progress <= 0.6) {
      const earlyProgress = progress / 0.6;
      return 0.72 * (1 - Math.pow(1 - earlyProgress, 1.08));
    }

    const lateProgress = (progress - 0.6) / 0.4;
    const smoothLate =
      lateProgress *
      lateProgress *
      lateProgress *
      (lateProgress * (lateProgress * 6 - 15) + 10);

    return 0.72 + 0.28 * smoothLate;
  };

  const step = (timestamp) => {
    const progress = Math.min((timestamp - startTime) / duration, 1);
    const easedProgress = easeOut(progress);

    for (const target of targets) {
      const nextValue = Math.round(target.start * (1 - easedProgress));
      target.apply(nextValue);
    }

    if (progress < 1) {
      scrollToTopAnimationFrame = requestAnimationFrame(step);
      return;
    }

    scrollToTopAnimationFrame = null;
  };

  scrollToTopAnimationFrame = requestAnimationFrame(step);
}

function stripIdsFromElementTree(root) {
  if (!root) {
    return;
  }

  root.removeAttribute?.('id');

  for (const element of root.querySelectorAll?.('[id]') || []) {
    element.removeAttribute('id');
  }
}

function removeMonthSwitchOverlay() {
  if (!activeMonthSwitchOverlay) {
    return;
  }

  activeMonthSwitchOverlay.remove();
  activeMonthSwitchOverlay = null;
}

function clearMonthHeaderAnimationState() {
  currentMonthLabel?.parentElement?.classList.remove(
    'month-switch-leave',
    'month-switch-enter'
  );
}

function getMonthSwitchTargets({ includeHeader = true } = {}) {
  const targets = [monthGalleryList];

  if (includeHeader) {
    targets.unshift(currentMonthLabel?.parentElement);
  }

  return targets.filter(Boolean);
}

function createMonthSwitchOverlayClone(
  sourceElement,
  containerRect,
  { viewportFixed = false } = {}
) {
  if (!sourceElement) {
    return null;
  }

  const rect = sourceElement.getBoundingClientRect();

  if (rect.width < 1 || rect.height < 1) {
    return null;
  }

  const clone = sourceElement.cloneNode(true);
  stripIdsFromElementTree(clone);

  clone.classList.remove('month-switch-enter', 'month-switch-leave');

  for (const element of clone.querySelectorAll?.(
    '.month-switch-enter, .month-switch-leave'
  ) || []) {
    element.classList.remove('month-switch-enter', 'month-switch-leave');
  }

  Object.assign(clone.style, {
    position: 'absolute',
    left: viewportFixed
      ? `${rect.left}px`
      : `${rect.left - containerRect.left}px`,
    top: viewportFixed
      ? `${rect.top}px`
      : `${rect.top - containerRect.top}px`,
    width: `${rect.width}px`,
    height: `${rect.height}px`,
    margin: '0',
    pointerEvents: 'none',
  });

  return clone;
}

function createMonthSwitchOverlay({
  includeHeader = true,
  viewportFixed = false,
} = {}) {
  if (!mainContent) {
    return null;
  }

  removeMonthSwitchOverlay();

  const overlay = document.createElement('div');
  const containerRect = mainContent.getBoundingClientRect();
  const snapshotSources = [];

  if (includeHeader) {
    snapshotSources.push(currentMonthLabel?.parentElement);
  }

  if (monthGalleryEmpty && getComputedStyle(monthGalleryEmpty).display !== 'none') {
    snapshotSources.push(monthGalleryEmpty);
  }

  if (monthGalleryList && monthGalleryList.children.length > 0) {
    snapshotSources.push(monthGalleryList);
  }

  Object.assign(overlay.style, {
    position: viewportFixed ? 'fixed' : 'absolute',
    inset: '0',
    zIndex: viewportFixed ? '320' : '8',
    pointerEvents: 'none',
    opacity: '1',
    transform: 'scale(1)',
    transformOrigin: 'center top',
    filter: 'blur(0px)',
    willChange: 'opacity, transform, filter',
  });

  for (const sourceElement of snapshotSources) {
    const clone = createMonthSwitchOverlayClone(sourceElement, containerRect, {
      viewportFixed,
    });

    if (clone) {
      overlay.appendChild(clone);
    }
  }

  if (!overlay.childElementCount) {
    return null;
  }

  if (viewportFixed) {
    document.body.appendChild(overlay);
  } else {
    mainContent.appendChild(overlay);
  }
  activeMonthSwitchOverlay = overlay;

  return overlay;
}

function fadeOutMonthSwitchOverlay(overlay) {
  if (!overlay || overlay !== activeMonthSwitchOverlay) {
    return;
  }

  overlay.style.transition =
    'opacity 1320ms cubic-bezier(0.22, 1, 0.36, 1), transform 1320ms cubic-bezier(0.22, 1, 0.36, 1), filter 1320ms cubic-bezier(0.22, 1, 0.36, 1)';

  requestAnimationFrame(() => {
    if (overlay !== activeMonthSwitchOverlay) {
      return;
    }

    overlay.style.opacity = '0';
    overlay.style.transform = 'scale(0.982)';
    overlay.style.filter = 'blur(10px)';

    setTimeout(() => {
      if (overlay === activeMonthSwitchOverlay) {
        removeMonthSwitchOverlay();
      } else {
        overlay.remove();
      }
    }, 1380);
  });
}

function playMonthSwitchAnimation({ includeHeader = true } = {}) {
  const targets = getMonthSwitchTargets({ includeHeader });

  for (const element of targets) {
    element.classList.remove('month-switch-leave', 'month-switch-enter');
    void element.offsetWidth;
    element.classList.add('month-switch-enter');
  }

  if (monthSwitchAnimationTimer) {
    clearTimeout(monthSwitchAnimationTimer);
  }

  monthSwitchAnimationTimer = setTimeout(() => {
    for (const element of targets) {
      element?.classList.remove('month-switch-enter');
    }
  }, 1720);
}

async function playMonthSwitchExitAnimation({ includeHeader = true } = {}) {
  const targets = getMonthSwitchTargets({ includeHeader });
  const hasVisibleContent =
    Boolean(currentSelection) &&
    Boolean(monthGalleryList) &&
    monthGalleryList.children.length > 0;

  if (!hasVisibleContent) {
    return;
  }

  for (const element of targets) {
    element.classList.remove('month-switch-enter', 'month-switch-leave');
    void element.offsetWidth;
    element.classList.add('month-switch-leave');
  }

  await new Promise((resolve) => {
    setTimeout(resolve, 340);
  });
}

async function refreshCurrentMonthWithFilterAnimation() {
  if (!currentSelection) {
    return;
  }

  clearMonthHeaderAnimationState();
  const monthSwitchOverlay = createMonthSwitchOverlay({ includeHeader: false });
  renderMonthGallery({ resetProgressive: true });

  playMonthSwitchAnimation({ includeHeader: false });

  requestAnimationFrame(() => {
    fadeOutMonthSwitchOverlay(monthSwitchOverlay);
  });
}

function renderSidebar() {
  sidebarTree.innerHTML = '';

  if (sidebarData.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'sidebar-empty';
    empty.textContent = 'まだ取り込みがありません';
    sidebarTree.appendChild(empty);
    return;
  }

  for (const yearEntry of sidebarData) {
    const yearBlock = document.createElement('div');
    yearBlock.className = 'year-block';
    yearBlock.dataset.year = String(yearEntry.year);

    const yearButton = document.createElement('button');
    yearButton.className = 'year-button';
    yearButton.type = 'button';

    const yearLeft = document.createElement('div');
    yearLeft.className = 'year-left';

    const toggle = document.createElement('span');
    toggle.className = 'year-toggle';
    toggle.textContent = expandedYears.has(yearEntry.year) ? '▾' : '▸';

    const label = document.createElement('span');
    label.textContent = `${yearEntry.year}年`;

    yearLeft.appendChild(toggle);
    yearLeft.appendChild(label);

    const yearCount = document.createElement('span');
    yearCount.className = 'year-count';
    yearCount.textContent = `${yearEntry.totalCount}枚`;

    yearButton.appendChild(yearLeft);
    yearButton.appendChild(yearCount);

    yearButton.addEventListener('click', () => {
      if (expandedYears.has(yearEntry.year)) {
        expandedYears.delete(yearEntry.year);
      } else {
        expandedYears.add(yearEntry.year);
      }
      renderSidebar();
    });

    const monthList = document.createElement('div');
    monthList.className = 'month-list';

    if (!expandedYears.has(yearEntry.year)) {
      monthList.classList.add('hidden');
    }

    for (const monthEntry of yearEntry.months) {
      const monthButton = document.createElement('button');
      monthButton.className = 'month-button';
      monthButton.type = 'button';
      monthButton.dataset.month = String(monthEntry.month);

      if (
        currentSelection &&
        currentSelection.year === yearEntry.year &&
        currentSelection.month === monthEntry.month
      ) {
        monthButton.classList.add('active');
      }

      const monthName = document.createElement('span');
      monthName.className = 'month-name';
      monthName.textContent = `${pad2(monthEntry.month)}月`;

      const monthCount = document.createElement('span');
      monthCount.className = 'month-count';
      monthCount.textContent = `${monthEntry.count}枚`;

      monthButton.appendChild(monthName);
      monthButton.appendChild(monthCount);

      monthButton.addEventListener('click', async () => {
        await selectMonth(yearEntry.year, monthEntry.month);
      });

      monthList.appendChild(monthButton);
    }

    yearBlock.appendChild(yearButton);
    yearBlock.appendChild(monthList);
    sidebarTree.appendChild(yearBlock);
  }
}

async function refreshSidebar() {
  sidebarData = await window.electronAPI.getSidebarData();

  if (expandedYears.size === 0) {
    for (const yearEntry of sidebarData) {
      expandedYears.add(yearEntry.year);
    }
  }

  if (currentSelection) {
    expandedYears.add(currentSelection.year);
  }

  renderSidebar();
  syncSelectionDependentSettingsUi();
}

async function selectMonth(year, month) {
  const requestId = ++monthSelectionRequestId;
  clearMonthHeaderAnimationState();
  const monthSwitchOverlay = createMonthSwitchOverlay({
    includeHeader: false,
    viewportFixed: true,
  });
  const photosPromise = window.electronAPI.getPhotosByMonth(year, month);
  let photos;

  try {
    photos = await photosPromise;
  } catch (error) {
    if (requestId === monthSelectionRequestId) {
      removeMonthSwitchOverlay();
      resetMonthSwitchClasses();
    }
    throw error;
  }

  if (requestId !== monthSelectionRequestId) {
    return;
  }

  currentSelection = { year, month };
  expandedYears.add(year);

  allCurrentMonthPhotos = photos;
  applyCurrentPhotoFilter();

  if (!syncSidebarSelectionState()) {
    renderSidebar();
  }
  syncSelectionDependentSettingsUi();
  stopScrollToTopAnimation();
  scrollGalleryViewToTop({ animated: false });
  renderMonthGallery({ resetProgressive: true });
  playMonthSwitchAnimation({ includeHeader: false });
  requestAnimationFrame(() => {
    fadeOutMonthSwitchOverlay(monthSwitchOverlay);
  });
}

async function handleImportResult(result, modeLabel) {
  importStatus.textContent = buildImportStatusMessage(result, modeLabel);

  if (!result || result.canceled) {
    return;
  }

  if (result.failedCount > 0) {
    showToast(`${modeLabel}: ${result.failedCount}件失敗しました`);
  }

  await refreshSidebar();

  if (result.selectedMonth) {
    await selectMonth(result.selectedMonth.year, result.selectedMonth.month);
    return;
  }

  const latestMonth = await window.electronAPI.getLatestMonth();

  if (latestMonth) {
    await selectMonth(latestMonth.year, latestMonth.month);
  } else {
    clearMainContent();
  }
}

async function startBackgroundWorldMetadataSync(targets) {
  if (!window.electronAPI.startWorldMetadataSync) {
    return;
  }

  const normalizedTargets = (Array.isArray(targets) ? targets : []).filter(
    (target) =>
      typeof target?.worldId === 'string' && target.worldId.trim().length > 0
  );

  if (normalizedTargets.length === 0) {
    return;
  }

  try {
    const result = await window.electronAPI.startWorldMetadataSync(
      normalizedTargets
    );

    if (!result?.ok) {
      showToast(result?.message || 'World情報の自動同期を開始できませんでした');
    }
  } catch (error) {
    showToast(`World情報の自動同期を開始できませんでした: ${error.message}`);
  }
}

async function handleTrackedFoldersRefreshResult(result, fallbackSelection) {
  importStatus.textContent = buildTrackedFoldersRefreshMessage(result);

  if (!result || result.canceled) {
    return;
  }

  if (result.ok === false) {
    showToast('更新に失敗しました');
    return;
  }

  if (result.noTrackedFolders) {
    showToast('更新対象フォルダがまだありません');
    return;
  }

  if (result.failedCount > 0) {
    showToast(`更新: ${result.failedCount}件失敗しました`);
  } else if (result.missingFolderPaths?.length > 0) {
    showToast(`見つからないフォルダが${result.missingFolderPaths.length}件あります`);
  } else if (!result.emptyRefresh && (result.importedCount || 0) > 0) {
    showToast('追跡フォルダを更新しました');
  } else {
    showToast('追跡フォルダを確認しました');
  }

  if (result.emptyRefresh) {
    return;
  }

  await refreshSidebar();

  if (result.selectedMonth) {
    await selectMonth(result.selectedMonth.year, result.selectedMonth.month);
    return;
  }

  if (fallbackSelection) {
    await selectMonth(fallbackSelection.year, fallbackSelection.month);
    return;
  }

  const latestMonth = await window.electronAPI.getLatestMonth();

  if (latestMonth) {
    await selectMonth(latestMonth.year, latestMonth.month);
  } else {
    clearMainContent();
  }
}

function setImportUiBusy(isBusy) {
  isImporting = isBusy;

  if (!isBusy) {
    resetProcessingProgress();
  }

  if (refreshTrackedFoldersButton) {
    refreshTrackedFoldersButton.disabled = isBusy;
  }

  if (settingsButton) {
    settingsButton.disabled = isBusy;
  }

  if (regenerateThumbnailsButton) {
    regenerateThumbnailsButton.disabled = isBusy;
  }

  if (regenerateThumbnailMonthSelect) {
    regenerateThumbnailMonthSelect.disabled =
      isBusy || regenerateThumbnailMonthSelect.options.length === 0;
  }

  if (regenerateThumbnailMonthButton) {
    regenerateThumbnailMonthButton.disabled =
      isBusy || !regenerateThumbnailMonthSelect || regenerateThumbnailMonthSelect.options.length === 0;
  }

  if (favoriteFilterButton) {
    favoriteFilterButton.disabled = isBusy || !currentSelection;
  }

  if (orientationFilterButton) {
    orientationFilterButton.disabled = isBusy || !currentSelection;
  }

  if (worldNameFilterButton) {
    worldNameFilterButton.disabled = isBusy || !currentSelection;
  }

  if (selectionModeButton) {
    selectionModeButton.disabled = isBusy || !currentSelection;
  }

  if (bulkFavoriteButton) {
    bulkFavoriteButton.disabled =
      isBusy || !isSelectionMode || selectedPhotoIds.size === 0;
  }

  if (bulkDeleteButton) {
    bulkDeleteButton.disabled =
      isBusy || !isSelectionMode || selectedPhotoIds.size === 0;
  }

  if (clearThumbnailCacheButton) {
    clearThumbnailCacheButton.disabled = isBusy || sidebarData.length === 0;
  }

  if (resetDatabaseButton) {
    resetDatabaseButton.disabled =
      isBusy || (sidebarData.length === 0 && trackedFolders.length === 0);
  }

  if (isBusy && typeof resetDropOverlay === 'function') {
    resetDropOverlay();
  }

  if (isBusy) {
    closeOrientationFilterMenu();
    closeWorldNameFilterMenu();
    closeRegenerateThumbnailMonthMenu();
  }
}

// Foreground operations share the same initial status/progress presentation.
function beginForegroundProgressOperation({
  statusMessage,
  progressMessage = '',
  showProgress = true,
}) {
  setImportUiBusy(true);
  importStatus.textContent = statusMessage;

  if (showProgress) {
    updateProcessingProgress({
      message: progressMessage || statusMessage,
    });
  }
}


async function runImportFlow(modeLabel, startMessage, importRunner) {
  if (isImporting) {
    showToast('取り込み中です。処理が終わってから次の取り込みを開始してください');
    return;
  }

  let result = null;

  beginForegroundProgressOperation({
    statusMessage: startMessage,
    progressMessage: startMessage,
  });

  try {
    result = await importRunner();
    await handleImportResult(result, modeLabel);
  } catch (error) {
    importStatus.textContent = `取り込みに失敗しました: ${error.message}`;
  } finally {
    setImportUiBusy(false);
  }

  await startBackgroundWorldMetadataSync(result?.worldMetadataTargets);
}


async function saveManualWorldEditForm({
  worldNameManual,
  worldUrl,
} = {}) {
  if (!currentModalPhoto) {
    return;
  }

  worldNameSaveStatus.textContent = '保存中...';

  const result = await window.electronAPI.updateWorldSettings(
    currentModalPhoto.id,
    {
      worldNameManual,
      worldUrl,
    }
  );

  if (!result?.ok) {
    worldNameSaveStatus.textContent = `保存に失敗しました: ${
      result?.message || '不明なエラー'
    }`;
    return;
  }

  syncSinglePhotoUpdate(result.photo);
  closeWorldNameEditModal();
  showToast('World設定を保存しました');
}

async function savePhotoMemo() {
  if (!currentModalPhoto || !modalPhotoMemoInput) {
    return;
  }

  if (modalPhotoMemoStatus) {
    modalPhotoMemoStatus.textContent = '保存中...';
  }

  if (modalPhotoMemoSaveButton) {
    modalPhotoMemoSaveButton.disabled = true;
  }

  const result = await window.electronAPI.updatePhotoMemo(
    currentModalPhoto.id,
    modalPhotoMemoInput.value
  );

  if (!result?.ok) {
    if (modalPhotoMemoStatus) {
      modalPhotoMemoStatus.textContent = `保存に失敗しました: ${
        result?.message || '不明なエラー'
      }`;
    }

  if (modalPhotoMemoSaveButton) {
    modalPhotoMemoSaveButton.disabled = false;
  }
  return;
  }

  syncSinglePhotoUpdate(result.photo);

  if (modalPhotoMemoStatus) {
    modalPhotoMemoStatus.textContent = '保存しました';
  }

  showToast('メモを保存しました');
}

async function rereadWorldName() {
  if (!currentModalPhoto) {
    return;
  }

  worldNameSaveStatus.textContent = '再取得中...';

  const result = await window.electronAPI.rereadWorldName({
    photoId: currentModalPhoto.id,
    worldUrl: modalWorldUrlInput?.value || currentModalPhoto.worldUrl || '',
  });

  if (!result || !result.ok) {
    worldNameSaveStatus.textContent = `再取得に失敗しました: ${
      result?.message || '不明なエラー'
    }`;
    return;
  }

  syncSinglePhotoUpdate(result.photo);
  closeWorldNameEditModal();
  showToast('World情報を再読み込みしました');
}

async function toggleFavorite(photoId, nextValue) {
  const result = await window.electronAPI.updateFavoriteStatus(
    photoId,
    nextValue
  );

  if (!result?.ok) {
    showToast(
      `お気に入り更新に失敗しました: ${result?.message || '不明なエラー'}`
    );
    return;
  }

  updatePhotoInCurrentCollections(result.photo);

  if (hasFullyRenderedMonthGallery()) {
    replaceRenderedPhotoCard(result.photo);
    syncRenderedFavoriteFilterState();
  } else if (
    isAnyPhotoFilterActive() &&
    !currentPhotos.some((photo) => photo.id === result.photo.id)
  ) {
    if (!removeRenderedPhotoCards([result.photo.id])) {
      renderMonthGallery({ resetProgressive: true });
    }
  } else {
    replaceRenderedPhotoCard(result.photo);
  }

  syncFavoriteFilterUi();

  if (currentModalPhoto?.id === result.photo.id) {
    showImageModalPhoto(result.photo);
  }

  showToast(nextValue ? 'お気に入りに追加しました' : 'お気に入りを解除しました');
}

async function toggleSelectedFavorites() {
  if (!isSelectionMode || selectedPhotoIds.size === 0) {
    return;
  }

  const selectedPhotos = getSelectedPhotosFromCurrentCollections();

  if (selectedPhotos.length === 0) {
    return;
  }

  const nextValue = !selectedPhotos.every((photo) => photo.isFavorite);
  const result = await window.electronAPI.updateFavoriteStatuses(
    [...selectedPhotoIds],
    nextValue
  );

  if (!result?.ok) {
    showToast(
      `お気に入り一括更新に失敗しました: ${result?.message || '不明なエラー'}`
    );
    return;
  }

  const updatedPhotos = Array.isArray(result.photos) ? result.photos : [];
  const updatedPhotoMap = new Map(
    updatedPhotos.map((photo) => [photo.id, photo])
  );

  allCurrentMonthPhotos = allCurrentMonthPhotos.map(
    (photo) => updatedPhotoMap.get(photo.id) || photo
  );
  applyCurrentPhotoFilter();
  clearSelectionState();
  renderMonthGallery({ resetProgressive: true });

  showToast(
    nextValue
      ? `選択した${updatedPhotos.length} 件をお気に入りに追加しました`
      : `選択した${updatedPhotos.length} 件のお気に入りを解除しました`
  );
}

async function refreshViewAfterDelete(
  targetSelection,
  {
    preferLocalRender = false,
    preferLocalSidebarUpdate = false,
    removedPhotoIds = [],
    removedCount = 0,
  } = {}
) {
  const updatedSidebarLocally = preferLocalSidebarUpdate
    ? applySidebarDeletionLocally(targetSelection, removedCount)
    : false;

  if (!updatedSidebarLocally) {
    await refreshSidebar();
  }

  if (targetSelection) {
    const monthStillExists = sidebarData.some(
      (yearEntry) =>
        yearEntry.year === targetSelection.year &&
        yearEntry.months.some(
          (monthEntry) => monthEntry.month === targetSelection.month
        )
    );

    if (monthStillExists) {
      const isCurrentTarget =
        currentSelection &&
        currentSelection.year === targetSelection.year &&
        currentSelection.month === targetSelection.month;

      if (preferLocalRender && isCurrentTarget) {
        const removedLocally = removeRenderedPhotoCards(removedPhotoIds);

        if (!removedLocally) {
          renderMonthGallery({ resetProgressive: true });
        } else {
          syncFavoriteFilterUi();
        }
        return;
      }

      await selectMonth(targetSelection.year, targetSelection.month);
      return;
    }
  }

  const latestMonth = getLatestSelectionFromSidebarData();

  if (latestMonth) {
    await selectMonth(latestMonth.year, latestMonth.month);
    return;
  }

  resetCurrentMonthState();
  clearSelectionState();
  renderSidebar();
  clearMainContent();
}

function closePhotoSpecificModalsForDeletedPhotos(photoIds) {
  const deletedIdSet = new Set(
    (Array.isArray(photoIds) ? photoIds : []).filter(
      (photoId) => Number.isInteger(photoId) && photoId > 0
    )
  );

  if (!deletedIdSet.size || !currentModalPhoto?.id) {
    return;
  }

  if (!deletedIdSet.has(currentModalPhoto.id)) {
    return;
  }

  closeWorldNameEditModal();
  closeImageModal();
}

// Maintenance actions share the same confirm / busy / toast flow, so keep the
// shell logic centralized and let each action only describe its own work.
async function runSettingsMaintenanceAction({
  isBlocked,
  confirmOptions,
  busyStatus,
  progressMessage = '',
  run,
  onSuccess,
  buildSuccessStatus,
  buildSuccessToast,
  buildErrorStatus,
  buildErrorToast,
}) {
  if (typeof isBlocked === 'function' && isBlocked()) {
    return null;
  }

  const confirmed = await openConfirmModal(confirmOptions);

  if (!confirmed) {
    return null;
  }

  beginForegroundProgressOperation({
    statusMessage: busyStatus,
    progressMessage,
    showProgress: Boolean(progressMessage),
  });

  try {
    const result = await run();

    if (!result?.ok) {
      throw new Error(result?.message || '処理に失敗しました');
    }

    if (typeof onSuccess === 'function') {
      await onSuccess(result);
    }

    if (typeof buildSuccessStatus === 'function') {
      importStatus.textContent = buildSuccessStatus(result);
    }

    if (typeof buildSuccessToast === 'function') {
      const successToast = buildSuccessToast(result);

      if (successToast) {
        showToast(successToast);
      }
    }

    return result;
  } catch (error) {
    const message =
      error instanceof Error ? error.message : String(error || '不明なエラー');

    if (typeof buildErrorStatus === 'function') {
      importStatus.textContent = buildErrorStatus(message);
    }

    if (typeof buildErrorToast === 'function') {
      const errorToast = buildErrorToast(message);

      if (errorToast) {
        showToast(errorToast);
      }
    }

    return null;
  } finally {
    setImportUiBusy(false);
    syncSettingsMaintenanceUi();
  }
}

// These maintenance actions intentionally reuse the normal delete / refresh
// flows so verification work exercises the same data paths as day-to-day use.
async function deleteCurrentMonthRegistrationsFromSettings() {
  const targetSelection = { ...currentSelection };

  await runSettingsMaintenanceAction({
    isBlocked: () =>
      isImporting ||
      !Number.isInteger(currentSelection?.year) ||
      !Number.isInteger(currentSelection?.month),
    confirmOptions: {
      title: '表示中の月を削除',
      message: `${targetSelection.year}年${targetSelection.month}月の登録を削除します。元画像ファイル自体は削除しません。続行しますか？`,
      confirmText: '削除する',
    },
    busyStatus: `${targetSelection.year}年${targetSelection.month}月の登録を削除中...`,
    run: () => window.electronAPI.deletePhotosByMonth(targetSelection),
    onSuccess: async (result) => {
      const deletedIds = Array.isArray(result.deletedPhotoIds)
        ? result.deletedPhotoIds
        : [];
      const deletedCount = Number(result.deletedCount) || deletedIds.length;

      closePhotoSpecificModalsForDeletedPhotos(deletedIds);
      removePhotosFromCurrentCollections(deletedIds);

      await refreshViewAfterDelete(targetSelection, {
        preferLocalRender: deletedCount > 0,
        preferLocalSidebarUpdate: deletedCount > 0,
        removedPhotoIds: deletedIds,
        removedCount: deletedCount,
      });
    },
    buildSuccessStatus: (result) => {
      const deletedIds = Array.isArray(result.deletedPhotoIds)
        ? result.deletedPhotoIds
        : [];
      const deletedCount = Number(result.deletedCount) || deletedIds.length;
      const failedCount = Number(result.failedCount) || 0;

      return failedCount > 0
        ? `${targetSelection.year}年${targetSelection.month}月: ${deletedCount}件削除 / 失敗 ${failedCount}件`
        : `${targetSelection.year}年${targetSelection.month}月: ${deletedCount}件削除`;
    },
    buildSuccessToast: (result) => {
      const failedCount = Number(result.failedCount) || 0;
      return failedCount > 0
        ? `月削除: ${failedCount}件失敗しました`
        : '表示中の月の登録を削除しました';
    },
    buildErrorStatus: (message) => `月削除に失敗しました: ${message}`,
    buildErrorToast: (message) => `月削除に失敗しました: ${message}`,
  });
}

async function deleteAllRegistrationsFromSettings() {
  const targetSelection = currentSelection ? { ...currentSelection } : null;

  await runSettingsMaintenanceAction({
    isBlocked: () => isImporting || sidebarData.length === 0,
    confirmOptions: {
      title: '全登録を削除',
      message:
        'すべての登録を削除します。元画像ファイル自体は削除しません。続行しますか？',
      confirmText: '削除する',
    },
    busyStatus: 'すべての登録を削除中...',
    run: () => window.electronAPI.deleteAllPhotos(),
    onSuccess: async (result) => {
      const deletedIds = Array.isArray(result.deletedPhotoIds)
        ? result.deletedPhotoIds
        : [];
      const deletedCount = Number(result.deletedCount) || deletedIds.length;

      closePhotoSpecificModalsForDeletedPhotos(deletedIds);
      removePhotosFromCurrentCollections(deletedIds);

      await refreshViewAfterDelete(targetSelection, {
        preferLocalRender: false,
        preferLocalSidebarUpdate: false,
        removedPhotoIds: deletedIds,
        removedCount: deletedCount,
      });
    },
    buildSuccessStatus: (result) => {
      const deletedIds = Array.isArray(result.deletedPhotoIds)
        ? result.deletedPhotoIds
        : [];
      const deletedCount = Number(result.deletedCount) || deletedIds.length;
      const failedCount = Number(result.failedCount) || 0;

      return failedCount > 0
        ? `全登録削除: ${deletedCount}件削除 / 失敗 ${failedCount}件`
        : `全登録削除: ${deletedCount}件削除`;
    },
    buildSuccessToast: (result) => {
      const failedCount = Number(result.failedCount) || 0;
      return failedCount > 0
        ? `全登録削除: ${failedCount}件失敗しました`
        : 'すべての登録を削除しました';
    },
    buildErrorStatus: (message) => `全登録削除に失敗しました: ${message}`,
    buildErrorToast: (message) => `全登録削除に失敗しました: ${message}`,
  });
}

async function clearThumbnailCacheFromSettings() {
  await runSettingsMaintenanceAction({
    isBlocked: () => isImporting || sidebarData.length === 0,
    confirmOptions: {
      title: 'サムネイルキャッシュ全削除',
      message:
        '管理しているサムネイルキャッシュをすべて削除します。元画像ファイルと登録データ自体は削除しません。続行しますか？',
      confirmText: '削除する',
    },
    busyStatus: 'サムネイルキャッシュを削除中...',
    progressMessage: 'サムネイルキャッシュを削除しています...',
    run: () => window.electronAPI.clearThumbnailCache(),
    onSuccess: async () => {
      clearThumbnailCacheInCurrentCollections();

      if (currentSelection) {
        renderMonthGallery({ resetProgressive: true });
      }
    },
    buildSuccessStatus: (result) =>
      `サムネイルキャッシュ削除: ${result.clearedCount || 0}件`,
    buildSuccessToast: () => 'サムネイルキャッシュを削除しました',
    buildErrorStatus: (message) =>
      `サムネイルキャッシュ削除に失敗しました: ${message}`,
    buildErrorToast: (message) =>
      `サムネイルキャッシュ削除に失敗しました: ${message}`,
  });
}

async function resetDatabaseFromSettings() {
  await runSettingsMaintenanceAction({
    isBlocked: () =>
      isImporting ||
      (sidebarData.length === 0 && trackedFolders.length === 0) ||
      !window.electronAPI.resetDatabase,
    confirmOptions: {
      title: 'DBを初期化',
      message:
        '登録データ、ラベル、メモ、ワールドキャッシュ、更新対象フォルダ、サムネイルキャッシュをすべて初期化します。元画像ファイル自体は削除しません。続行しますか？',
      confirmText: '初期化する',
    },
    busyStatus: 'DBを初期化中...',
    progressMessage: 'アプリデータを初期化しています...',
    run: () => window.electronAPI.resetDatabase(),
    onSuccess: async () => {
      closePhotoLabelModal();
      closeWorldNameEditModal();
      closeImageModal();

      sidebarData = [];
      trackedFolders = [];
      resetCurrentMonthState();
      expandedYears.clear();
      clearSelectionState();
      renderSidebar();
      clearMainContent();
      renderTrackedFolderList();
    },
    buildSuccessStatus: (result) =>
      `DB初期化: 写真 ${result.photoCount || 0}件 / ` +
      `フォルダ ${result.trackedFolderCount || 0}件 / ` +
      `キャッシュ ${result.worldCacheCount || 0}件 / ` +
      `ラベル ${result.tagCount || 0}件`,
    buildSuccessToast: () => 'DBを初期化しました',
    buildErrorStatus: (message) => `DB初期化に失敗しました: ${message}`,
    buildErrorToast: (message) => `DB初期化に失敗しました: ${message}`,
  });
}

function hasDraggedFiles(dataTransfer) {
  return Array.from(dataTransfer?.types || []).includes('Files');
}

function updateDropOverlayViewportRect() {
  if (!dropOverlay || !mainContent) {
    return;
  }

  const rect = mainContent.getBoundingClientRect();

  const left = Math.max(rect.left, 0);
  const top = Math.max(rect.top, 0);
  const right = Math.min(rect.right, window.innerWidth);
  const bottom = Math.min(rect.bottom, window.innerHeight);

  const width = Math.max(0, right - left);
  const height = Math.max(0, bottom - top);

  dropOverlay.style.left = `${left}px`;
  dropOverlay.style.top = `${top}px`;
  dropOverlay.style.width = `${width}px`;
  dropOverlay.style.height = `${height}px`;
}

function setDropOverlayVisible(isVisible) {
  if (!dropOverlay) {
    return;
  }

  if (isVisible) {
    updateDropOverlayViewportRect();
  }

  dropOverlay.classList.toggle('hidden', !isVisible);
}

function setDropTargetActive(isActive) {
  if (!mainContent) {
    return;
  }

  mainContent.classList.toggle('drop-target-active', isActive);
}

function clearDropOverlayWatchTimer() {
  if (!dropOverlayWatchTimer) {
    return;
  }

  clearTimeout(dropOverlayWatchTimer);
  dropOverlayWatchTimer = null;
}

function keepDropOverlayAlive() {
  clearDropOverlayWatchTimer();

  dropOverlayWatchTimer = setTimeout(() => {
    resetDropOverlay();
  }, 140);
}

function resetDropOverlay() {
  clearDropOverlayWatchTimer();
  setDropOverlayVisible(false);
  setDropTargetActive(false);

  if (dropOverlay) {
    dropOverlay.style.removeProperty('left');
    dropOverlay.style.removeProperty('top');
    dropOverlay.style.removeProperty('width');
    dropOverlay.style.removeProperty('height');
  }
}

function initializeDragAndDropImport() {
  if (!mainContent || !dropOverlay) {
    return;
  }

  window.addEventListener(
    'scroll',
    () => {
      if (!dropOverlay.classList.contains('hidden')) {
        updateDropOverlayViewportRect();
      }
    },
    true
  );

  window.addEventListener('resize', () => {
    if (!dropOverlay.classList.contains('hidden')) {
      updateDropOverlayViewportRect();
    }
  });

  window.addEventListener('dragover', (event) => {
    if (!hasDraggedFiles(event.dataTransfer)) {
      return;
    }

    event.preventDefault();
  });

  window.addEventListener('drop', (event) => {
    if (!hasDraggedFiles(event.dataTransfer)) {
      return;
    }

    if (!mainContent.contains(event.target)) {
      event.preventDefault();
      resetDropOverlay();
    }
  });

  window.addEventListener('dragend', () => {
    resetDropOverlay();
  });

  window.addEventListener('blur', () => {
    resetDropOverlay();
  });

  mainContent.addEventListener('dragenter', (event) => {
    if (isImporting || !hasDraggedFiles(event.dataTransfer)) {
      return;
    }

    event.preventDefault();
    setDropTargetActive(true);
    setDropOverlayVisible(true);
    keepDropOverlayAlive();
  });

  mainContent.addEventListener('dragover', (event) => {
    if (isImporting || !hasDraggedFiles(event.dataTransfer)) {
      return;
    }

    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
    setDropTargetActive(true);
    setDropOverlayVisible(true);
    keepDropOverlayAlive();
  });

  mainContent.addEventListener('dragleave', (event) => {
    if (isImporting || !hasDraggedFiles(event.dataTransfer)) {
      return;
    }

    event.preventDefault();

    const nextElement = document.elementFromPoint(
      event.clientX,
      event.clientY
    );

    if (nextElement && mainContent.contains(nextElement)) {
      keepDropOverlayAlive();
      return;
    }

    resetDropOverlay();
  });

  mainContent.addEventListener('drop', async (event) => {
    if (isImporting || !hasDraggedFiles(event.dataTransfer)) {
      return;
    }

    event.preventDefault();
    resetDropOverlay();

    await runImportFlow(
      'ドラッグ&ドロップ取り込み',
      'ドラッグ&ドロップ取り込み中...',
      async () => {
        const files = Array.from(event.dataTransfer.files || []);

        if (files.length === 0) {
          return {
            canceled: false,
            totalSelected: 0,
            importedCount: 0,
            newCount: 0,
            updatedCount: 0,
            failedCount: 0,
            failedFiles: [],
            selectedMonth: null,
            emptyDrop: true,
          };
        }

        return window.electronAPI.importDroppedFiles(files);
      }
    );
  });
}

async function initializeApp() {
  await refreshSidebar();

  const latestMonth = await window.electronAPI.getLatestMonth();

  if (latestMonth) {
    await selectMonth(latestMonth.year, latestMonth.month);
  } else {
    clearMainContent();
    renderSidebar();
  }

  syncFavoriteFilterUi();
}

// Toolbar and maintenance actions that kick off foreground work.
function bindForegroundActionControls() {
  regenerateThumbnailsButton?.addEventListener(
    'click',
    async (event) => {
      event.preventDefault();
      event.stopImmediatePropagation();

      if (isImporting) {
        showToast('再生成中です。処理が終わってから実行してください');
        return;
      }

      const selectedMonthValue = regenerateThumbnailMonthSelect?.value || '';

      if (!/^\d{4}-\d{2}$/.test(selectedMonthValue)) {
        showToast('再生成する月を選択してください');
        return;
      }

      const [targetYear, targetMonth] = selectedMonthValue.split('-').map(Number);

      beginForegroundProgressOperation({
        statusMessage: 'サムネイルを再生成しています...',
        progressMessage: 'サムネイルを再生成しています...',
      });

      try {
        const result = await window.electronAPI.regenerateThumbnails({
          year: targetYear,
          month: targetMonth,
        });

        importStatus.textContent = buildScopedRegenerateThumbnailsMessage(result);

        if (result?.failedCount > 0) {
          showToast(`サムネイル再生成 ${result.failedCount}件失敗しました`);
        } else if (result?.ok) {
          showToast('サムネイル再生成が完了しました');
        }

        if (currentSelection) {
          await selectMonth(currentSelection.year, currentSelection.month);
        }
      } catch (error) {
        importStatus.textContent = `サムネイル再生成に失敗しました: ${error.message}`;
      } finally {
        setImportUiBusy(false);
      }
    },
    { capture: true }
  );

  refreshTrackedFoldersButton?.addEventListener('click', async () => {
    if (isImporting) {
      showToast('別の処理中です。完了してから更新してください');
      return;
    }

    const fallbackSelection = currentSelection
      ? { ...currentSelection }
      : null;
    let result = null;

    beginForegroundProgressOperation({
      statusMessage: '追跡フォルダを更新中...',
      progressMessage: '追跡フォルダを更新中...',
    });

    try {
      result = await window.electronAPI.refreshTrackedFolders();
      setImportUiBusy(false);
      await handleTrackedFoldersRefreshResult(result, fallbackSelection);
    } catch (error) {
      importStatus.textContent = `更新に失敗しました: ${error.message}`;
    } finally {
      setImportUiBusy(false);
    }

    await startBackgroundWorldMetadataSync(result?.worldMetadataTargets);
  });
}

// Header filters stay interactive while month content changes underneath them.
function bindHeaderFilterControls() {
  favoriteFilterButton?.addEventListener('click', async () => {
    if (!currentSelection) {
      return;
    }

    isFavoriteFilterOnly = !isFavoriteFilterOnly;
    applyCurrentPhotoFilter();
    await refreshCurrentMonthWithFilterAnimation();
  });

  orientationFilterButton?.addEventListener('click', (event) => {
    event.stopPropagation();

    if (!currentSelection || isImporting) {
      return;
    }

    setOrientationFilterMenuOpen(!isOrientationFilterMenuOpen);
  });

  photoLabelFilterButton?.addEventListener('click', (event) => {
    event.stopPropagation();

    if (!currentSelection || isImporting) {
      return;
    }

    setPhotoLabelFilterMenuOpen(!isPhotoLabelFilterMenuOpen);
  });

  worldNameFilterButton?.addEventListener('click', (event) => {
    event.stopPropagation();

    if (!currentSelection || isImporting) {
      return;
    }

    setWorldNameFilterMenuOpen(!isWorldNameFilterMenuOpen);
  });

  for (const item of orientationFilterItems) {
    item.addEventListener('click', async (event) => {
      event.stopPropagation();
      await setOrientationFilter(item.dataset.orientationFilter || 'all');
    });
  }

  photoLabelFilterMenu?.addEventListener('click', async (event) => {
    const modeTarget = event.target.closest('[data-photo-label-filter-mode]');

    if (modeTarget) {
      event.stopPropagation();
      await setPhotoLabelFilterMode(modeTarget.dataset.photoLabelFilterMode || 'or');
      return;
    }

    const target = event.target.closest('[data-photo-label-filter]');

    if (!target) {
      return;
    }

    event.stopPropagation();
    await togglePhotoLabelFilter(target.dataset.photoLabelFilter || '');
  });

  worldNameFilterInput?.addEventListener('input', (event) => {
    scheduleWorldNameFilterApply(event.target.value);
  });

  worldNameFilterInput?.addEventListener('keydown', async (event) => {
    if (event.key !== 'Enter') {
      return;
    }

    event.preventDefault();
    clearWorldNameFilterInputTimer();
    await applyWorldNameFilter(event.currentTarget.value);
  });

  worldNameFilterClearButton?.addEventListener('click', async (event) => {
    event.stopPropagation();

    if (worldNameFilterInput) {
      worldNameFilterInput.value = '';
      worldNameFilterInput.focus({ preventScroll: true });
    }

    clearWorldNameFilterInputTimer();
    await applyWorldNameFilter('');
  });
}

// Modal-level editors and detail actions are grouped here so photo-specific
// behavior can be traced without scanning the entire file bottom.
function bindPhotoAndEditModalControls() {
  saveWorldNameButton?.addEventListener('click', async () => {
    await saveManualWorldEditForm({
      worldNameManual: modalWorldNameInput?.value || '',
      worldUrl: modalWorldUrlInput?.value || '',
    });
  });

  modalWorldNameInput?.addEventListener('input', () => {
    worldNameSaveStatus.textContent = '';
  });

  modalWorldUrlInput?.addEventListener('input', () => {
    worldNameSaveStatus.textContent = '';
  });

  modalPhotoMemoSaveButton?.addEventListener('click', async () => {
    await savePhotoMemo();
  });

  modalPhotoMemoInput?.addEventListener('input', () => {
    resizeModalPhotoMemoInput();

    if (modalPhotoMemoStatus) {
      modalPhotoMemoStatus.textContent = '';
    }
  });

  modalPhotoMemoInput?.addEventListener('keydown', async (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
      event.preventDefault();
      await savePhotoMemo();
    }
  });

  clearWorldNameButton?.addEventListener('click', async () => {
    await saveManualWorldEditForm({
      worldNameManual: '',
      worldUrl: currentModalPhoto?.worldUrl || '',
    });
  });

  rereadWorldNameButton?.addEventListener('click', async () => {
    await rereadWorldName();
  });

  openWorldNameEditButton?.addEventListener('click', () => {
    openWorldNameEditModal();
  });

  bindSubModalCloseTriggers(
    worldNameEditBackdrop,
    worldNameEditClose,
    closeWorldNameEditModal
  );

  bindSubModalCloseTriggers(imageModalBackdrop, imageModalClose, closeImageModal);
  bindSubModalCloseTriggers(confirmModalBackdrop, confirmModalClose, () => {
    closeConfirmModal(false);
  });

  confirmModalCancelButton?.addEventListener('click', () => {
    closeConfirmModal(false);
  });

  confirmModalConfirmButton?.addEventListener('click', () => {
    closeConfirmModal(true);
  });

  modalWorldLink?.addEventListener('click', async (event) => {
    event.preventDefault();

    if (!currentModalPhoto?.worldUrl) {
      return;
    }

    const result = await window.electronAPI.openExternalUrl(
      currentModalPhoto.worldUrl
    );

    if (!result?.ok) {
      showToast(`リンクを開けませんでした: ${result?.message || '不明なエラー'}`);
    }
  });

  modalOpenWorldButton?.addEventListener('click', async () => {
    if (!currentModalPhoto?.worldUrl) {
      return;
    }

    const result = await window.electronAPI.openExternalUrl(
      currentModalPhoto.worldUrl
    );

    if (!result?.ok) {
      showToast(`リンクを開けませんでした: ${result?.message || '不明なエラー'}`);
    }
  });

  modalOpenOriginalButton?.addEventListener('click', async () => {
    if (!currentModalPhoto?.filePath) {
      return;
    }

    const result = await window.electronAPI.openLocalFile({
      photoId: currentModalPhoto.id,
      filePath: currentModalPhoto.filePath,
    });

    if (result?.photo) {
      syncSinglePhotoUpdate(result.photo);
    }

    if (!result?.ok) {
      showToast(
        `画像を開けませんでした: ${result?.message || '不明なエラー'}`
      );
    }
    if (result.recovered) {
      showToast('画像の保存場所を更新しました');
    }
  });

  modalOpenFolderButton?.addEventListener('click', async () => {
    if (!currentModalPhoto?.filePath) {
      return;
    }

    const result = await window.electronAPI.openContainingFolder({
      photoId: currentModalPhoto.id,
      filePath: currentModalPhoto.filePath,
    });

    if (result?.photo) {
      syncSinglePhotoUpdate(result.photo);
    }

    if (!result?.ok) {
      showToast(
        `保存先フォルダを開けませんでした: ${result?.message || '不明なエラー'}`
      );
    }
    if (result.recovered) {
      showToast('画像の保存場所を更新しました');
    }
  });

  modalFavoriteButton?.addEventListener('click', async () => {
    if (!currentModalPhoto?.id) {
      return;
    }

    await toggleFavorite(currentModalPhoto.id, !currentModalPhoto.isFavorite);
  });

  modalDeletePhotoButton?.addEventListener('click', async () => {
    if (isImporting) {
      showToast('処理中です。完了してから実行してください');
      return;
    }

    if (!currentModalPhoto?.id) {
      return;
    }

    const confirmed = await openConfirmModal({
      title: '登録を削除',
      message:
        'この画像の登録を削除します。元画像ファイル自体は削除しません。続行しますか？',
      confirmText: '削除する',
    });

    if (!confirmed) {
      return;
    }

    const targetSelection = currentSelection ? { ...currentSelection } : null;

    try {
      const deleteTargetId = currentModalPhoto.id;
      const result = await window.electronAPI.deletePhoto(deleteTargetId);

      if (!result?.ok) {
        showToast(`削除に失敗しました: ${result?.message || '不明なエラー'}`);
        return;
      }

      closeWorldNameEditModal();
      closeImageModal();

      removePhotoFromCurrentCollections(deleteTargetId);
      await refreshViewAfterDelete(targetSelection, {
        preferLocalRender: true,
        preferLocalSidebarUpdate: true,
        removedPhotoIds: [deleteTargetId],
        removedCount: 1,
      });

      showToast('登録を削除しました');
    } catch (error) {
      showToast(`削除に失敗しました: ${error.message}`);
    }
  });

  imageModalContent?.addEventListener('wheel', handleImageModalWheel, {
    passive: false,
  });
}

// Settings modal keeps folder management and destructive maintenance together.
function bindSettingsModalControls() {
  settingsButton?.addEventListener('click', async () => {
    if (isImporting) {
      return;
    }

    await openSettingsModal();
  });

  bindSubModalCloseTriggers(
    settingsModalBackdrop,
    settingsModalClose,
    closeSettingsModal
  );
  settingsModalContent?.addEventListener('wheel', handleSettingsModalWheel, {
    passive: false,
  });

  addTrackedFolderButton?.addEventListener('click', async () => {
    const result = await window.electronAPI.addTrackedFolder();

    if (!result?.ok) {
      showToast(
        `フォルダの追加に失敗しました: ${result?.message || '不明なエラー'}`
      );
      return;
    }

    trackedFolders = Array.isArray(result.folders) ? result.folders : trackedFolders;
    renderTrackedFolderList();

    if (!result.canceled && result.folder?.folder_path) {
      showToast('更新対象フォルダを追加しました');
    }
  });

  trackedFolderList?.addEventListener('click', async (event) => {
    const removeButton = event.target.closest('[data-tracked-folder-path]');

    if (!removeButton) {
      return;
    }

    const folderPath = removeButton.dataset.trackedFolderPath;

    if (!folderPath) {
      return;
    }

    const confirmed = await openConfirmModal({
      title: '更新対象フォルダを削除',
      message:
        'このフォルダを更新対象一覧から外します。登録済みの写真データ自体は削除されません。続行しますか？',
      confirmText: '削除する',
    });

    if (!confirmed) {
      return;
    }

    const result = await window.electronAPI.removeTrackedFolder(folderPath);

    if (!result?.ok) {
      showToast(
        `フォルダの削除に失敗しました: ${result?.message || '不明なエラー'}`
      );
      return;
    }

    trackedFolders = Array.isArray(result.folders) ? result.folders : trackedFolders;
    renderTrackedFolderList();
    showToast('更新対象フォルダを削除しました');
  });

  deleteCurrentMonthRegistrationsButton?.addEventListener('click', async () => {
    await deleteCurrentMonthRegistrationsFromSettings();
  });

  deleteAllRegistrationsButton?.addEventListener('click', async () => {
    await deleteAllRegistrationsFromSettings();
  });

  clearThumbnailCacheButton?.addEventListener('click', async () => {
    await clearThumbnailCacheFromSettings();
  });

  resetDatabaseButton?.addEventListener('click', async () => {
    await resetDatabaseFromSettings();
  });
}

// Batch actions live in the page header and operate on current month selection.
function bindSelectionControls() {
  selectionModeButton?.addEventListener('click', () => {
    if (!currentSelection) {
      return;
    }

    if (isSelectionMode) {
      clearSelectionState();
    } else {
      isSelectionMode = true;
      selectedPhotoIds.clear();
      syncSelectionUi();
    }
    syncRenderedSelectionState();
  });

  bulkFavoriteButton?.addEventListener('click', async () => {
    await toggleSelectedFavorites();
  });

  bulkDeleteButton?.addEventListener('click', async () => {
    if (!isSelectionMode || selectedPhotoIds.size === 0) {
      return;
    }

    const confirmed = await openConfirmModal({
      title: '選択した登録を削除',
      message: `選択した${selectedPhotoIds.size} 件の登録を削除します。元画像ファイル自体は削除しません。続行しますか？`,
      confirmText: '削除する',
    });

    if (!confirmed) {
      return;
    }

    const targetIds = [...selectedPhotoIds];
    const targetSelection = currentSelection ? { ...currentSelection } : null;

    setImportUiBusy(true);
    importStatus.textContent = '選択した登録を削除中...';

    try {
      const result = await window.electronAPI.deletePhotos(targetIds);

      if (!result?.ok) {
        throw new Error(result?.message || '削除に失敗しました');
      }

      const deletedIds = Array.isArray(result.deletedPhotoIds)
        ? result.deletedPhotoIds
        : [];
      const deletedCount = deletedIds.length;
      const failedCount = Number(result.failedCount) || 0;

      removePhotosFromCurrentCollections(deletedIds);

      clearSelectionState();
      await refreshViewAfterDelete(targetSelection, {
        preferLocalRender: deletedCount > 0,
        preferLocalSidebarUpdate: deletedCount > 0,
        removedPhotoIds: deletedIds,
        removedCount: deletedCount,
      });

      importStatus.textContent =
        failedCount > 0
          ? `選択削除: ${deletedCount}件削除 / 失敗 ${failedCount}件`
          : `選択削除: ${deletedCount}件削除`;

      if (failedCount > 0) {
        showToast(`選択削除: ${failedCount}件失敗しました`);
      } else {
        showToast('選択した登録を削除しました');
      }
    } catch (error) {
      importStatus.textContent = `選択削除に失敗しました: ${error.message}`;
      showToast(`選択削除に失敗しました: ${error.message}`);
    } finally {
      setImportUiBusy(false);
      syncSelectionUi();
    }
  });
}

// Theme and font preferences are lightweight local UI settings.
function bindAppearanceControls() {
  themeToggleButton?.addEventListener('click', () => {
    toggleTheme();
  });

  fontOptionButtons.forEach((button) => {
    button.addEventListener('click', () => {
      applyFontPreference(button.dataset.fontOption || 'standard');
    });
  });
}

// Global document listeners keep dropdowns and modal keyboard behavior
// consistent across the whole app surface.
function bindGlobalDocumentInteractions() {
  document.addEventListener('click', (event) => {
    closeManagedDropdownsFromOutsideClick(event.target);
  });

  document.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape') {
      return;
    }

    if (closeManagedDropdownFromEscape()) {
      return;
    }

    if (confirmModal && !confirmModal.classList.contains('hidden')) {
      closeConfirmModal(false);
      return;
    }

    if (worldNameEditModal && !worldNameEditModal.classList.contains('hidden')) {
      closeWorldNameEditModal();
      return;
    }

    if (photoLabelModal && !photoLabelModal.classList.contains('hidden')) {
      closePhotoLabelModal();
      return;
    }

    if (settingsModal && !settingsModal.classList.contains('hidden')) {
      closeSettingsModal();
      return;
    }

    if (imageModal && !imageModal.classList.contains('hidden')) {
      closeImageModal();
    }
  });

  document.addEventListener('keydown', (event) => {
    if (
      !imageModal ||
      imageModal.classList.contains('hidden') ||
      worldNameEditModal?.classList.contains('hidden') === false ||
      photoLabelModal?.classList.contains('hidden') === false
    ) {
      return;
    }

    if (event.key === 'ArrowLeft') {
      event.preventDefault();
      stepImageModalPhoto(-1);
    }

    if (event.key === 'ArrowRight') {
      event.preventDefault();
      stepImageModalPhoto(1);
    }
  });

  document.addEventListener('click', handleDelegatedSubModalClose, true);
}

// IPC listeners are registered once so late events only touch their dedicated
// UI surfaces.
function bindIpcEventListeners() {
  // Foreground progress bars belong to explicit import/refresh/maintenance
  // actions. Late IPC events should not reopen the bar after the UI is idle.
  window.electronAPI.onProcessingProgress?.((payload) => {
    if (payload?.operation === 'world-metadata-sync') {
      handleWorldMetadataSyncProgress(payload);
      return;
    }

    if (!isImporting) {
      return;
    }

    updateProcessingProgress(payload);
  });

  window.electronAPI.onWorldMetadataUpdated?.((payload) => {
    applyWorldMetadataUpdated(payload);
  });
}

// Boot sequence for renderer-only concerns. Keeping the order explicit makes
// it easier to reason about future regressions.
function initializeRendererBindings() {
  bindForegroundActionControls();
  bindHeaderFilterControls();
  bindPhotoAndEditModalControls();
  bindSettingsModalControls();
  bindSelectionControls();
  bindAppearanceControls();
  bindGlobalDocumentInteractions();
  bindIpcEventListeners();
}

function initializeRendererUi() {
  initializeTheme();
  initializeFontPreference();
  initializeImageModalUi();
  initializeWorldNameEditUi();
  initializePhotoLabelUi();
  initializeModalCloseIcons();
  initializeTrackedFolderAccordion();
  initializeTopToolbarLayout();
  initializeDragAndDropImport();
  initializeProgressiveMonthGalleryLoading();
  initializeScrollToTopAnimationInterrupts();
}

initializeRendererUi();
initializeRendererBindings();
syncFavoriteFilterUi();
initializeApp();


