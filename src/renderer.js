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
const appRoot = document.querySelector('.app');
const sidebarTree = document.getElementById('sidebar-tree');
const currentMonthLabel = document.getElementById('current-month-label');
const currentMonthCount = document.getElementById('current-month-count');
const monthGalleryEmpty = document.getElementById('month-gallery-empty');
const monthGalleryList = document.getElementById('month-gallery-list');
const mainContent = document.querySelector('.main-content');
const mainHeader = document.querySelector('.main-header');
const mainHeaderTitleGroup = mainHeader?.firstElementChild;
const mainHeaderActions = mainHeader?.querySelector('.main-header-actions');
const dropOverlay = document.getElementById('drop-overlay');
const topStickyShell = document.querySelector('.top-sticky-shell');
const sidebar = document.querySelector('.sidebar');
const sidebarHeader = document.querySelector('.sidebar-header');
const sidebarHeaderTitle = sidebarHeader?.querySelector('h2');
const sidebarHeaderDescription = sidebarHeader?.querySelector('p');

// Keep the content area in a loading state until sidebar/month restoration
// finishes so the empty-state shell does not flash on startup.
appRoot?.classList.add('is-app-initializing');
appRoot?.setAttribute('aria-busy', 'true');

// Fail safe: if renderer startup hits an unexpected error, do not leave the
// whole window permanently covered by the loading spinner.
function forceClearInitializationState() {
  appRoot?.classList.remove('is-app-initializing');
  appRoot?.setAttribute('aria-busy', 'false');
}

let syncMainHeaderLayoutFrame = 0;

function measureInlineChildrenWidth(container) {
  if (!container) {
    return 0;
  }

  const styles = window.getComputedStyle(container);
  const gap =
    Number.parseFloat(styles.columnGap || styles.gap || '0') ||
    Number.parseFloat(styles.rowGap || '0') ||
    0;
  const visibleChildren = Array.from(container.children).filter((child) => {
    const childStyles = window.getComputedStyle(child);
    return (
      !child.hasAttribute('hidden') &&
      childStyles.display !== 'none' &&
      childStyles.visibility !== 'hidden'
    );
  });

  if (visibleChildren.length === 0) {
    return 0;
  }

  const childrenWidth = visibleChildren.reduce((total, child) => {
    return total + child.getBoundingClientRect().width;
  }, 0);

  return childrenWidth + gap * Math.max(visibleChildren.length - 1, 0);
}

function syncMainHeaderResponsiveLayout() {
  if (!mainHeader || !mainHeaderTitleGroup || !mainHeaderActions) {
    return;
  }

  // Measure the header in its unstacked state first; otherwise the stacked
  // layout's 100% action row width keeps forcing itself to stay stacked.
  mainHeader.classList.remove('is-actions-stacked');

  const availableWidth = mainHeader.clientWidth;
  const mainHeaderGap =
    Number.parseFloat(
      window.getComputedStyle(mainHeader).columnGap ||
        window.getComputedStyle(mainHeader).gap ||
        '0'
    ) || 0;
  const stackThresholdPx = 96;
  const titleWidth = measureInlineChildrenWidth(mainHeaderTitleGroup);
  const actionsWidth = measureInlineChildrenWidth(mainHeaderActions);
  const shouldStackActions =
    availableWidth > 0 &&
    titleWidth + actionsWidth + mainHeaderGap >
      availableWidth + stackThresholdPx;

  mainHeader.classList.toggle('is-actions-stacked', shouldStackActions);
}

function scheduleMainHeaderResponsiveLayout() {
  if (syncMainHeaderLayoutFrame) {
    cancelAnimationFrame(syncMainHeaderLayoutFrame);
  }

  syncMainHeaderLayoutFrame = requestAnimationFrame(() => {
    syncMainHeaderLayoutFrame = 0;
    syncMainHeaderResponsiveLayout();
  });
}

const appInitializationFailsafeTimer = setTimeout(() => {
  console.warn('[renderer startup] forcing initialization overlay to close');
  forceClearInitializationState();
}, 3000);

window.addEventListener('error', () => {
  setTimeout(forceClearInitializationState, 0);
});

window.addEventListener('unhandledrejection', () => {
  setTimeout(forceClearInitializationState, 0);
});

// Header filter controls.
const favoriteFilterButton = document.getElementById('favorite-filter-btn');
const photoSortButton = document.getElementById('photo-sort-btn');
const photoSortIcon = document.getElementById('photo-sort-icon');
const photoDensityButton = document.getElementById('photo-density-btn');
const photoDensityIcon = document.getElementById('photo-density-icon');
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
const worldNameFilterSearchButton = worldNameFilterClearButton;

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
const modalPhotoMemoBlock = modalPhotoMemoInput?.closest('.modal-world-meta-block');

const modalWorldLink = document.getElementById('modal-world-link');
const modalOpenWorldButton = document.getElementById('modal-open-world-btn');
const modalOpenOriginalButton = document.getElementById(
  'modal-open-original-btn'
);
const modalOpenFolderButton = document.getElementById('modal-open-folder-btn');
const modalDeletePhotoButton = document.getElementById(
  'modal-delete-photo-btn'
);

let modalFavoriteButton = document.getElementById('modal-favorite-btn');
let modalFavoriteIcon = document.getElementById('modal-favorite-icon');
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
const settingsFontSection = settingsModal?.querySelector('.settings-font-section');
const settingsSectionHeader = settingsModal?.querySelector(
  '.settings-section-header'
);
const settingsMaintenanceSection = settingsModal?.querySelector(
  '.settings-maintenance-section'
);
const settingsMaintenanceActions = settingsMaintenanceSection?.querySelector(
  '.settings-maintenance-actions'
);
let trackedFolderSettingsMeta = null;
let settingsMaintenanceStatus = null;
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
const reimportRegisteredPhotosButton = document.getElementById(
  'reimport-registered-photos-btn'
);
const resetDatabaseButton = document.getElementById('reset-database-btn');
const settingsUninstallLaunchButton = document.getElementById(
  'settings-uninstall-launch-btn'
);
const toolbar = document.querySelector('.toolbar');
const toolbarRight = toolbar?.querySelector('.toolbar-right');
const pageHeaderActions = document.querySelector('.page-header-actions');
const fontOptionButtons = Array.from(
  document.querySelectorAll('[data-font-option]')
);

// App uninstall modal stays separate from destructive maintenance so the
// user always chooses the uninstall mode explicitly before the final confirm.
const uninstallModal = document.getElementById('uninstall-modal');
const uninstallModalBackdrop = document.getElementById(
  'uninstall-modal-backdrop'
);
const uninstallModalClose = document.getElementById('uninstall-modal-close');
const uninstallAppButton = document.getElementById('uninstall-app-btn');
const uninstallAppAndDeleteDataButton = document.getElementById(
  'uninstall-app-and-delete-data-btn'
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
const BACKGROUND_IMAGE_STORAGE_KEY =
  'vrchat-world-photo-manager-background-image';
const PHOTO_CARD_DENSITY_STORAGE_KEY =
  'vrchat-world-photo-manager-photo-card-density';

// Batch selection controls for the current month view.
const selectionModeButton = document.getElementById('selection-mode-btn');
const bulkFavoriteButton = document.getElementById('bulk-favorite-btn');
const bulkDeleteButton = document.getElementById('bulk-delete-btn');

// Sidebar/month/gallery state for the active selection.
let sidebarData = [];
let worldSidebarData = [];
let currentSelection = null;
let currentPhotos = [];
let allCurrentMonthPhotos = [];
let currentSidebarMode = 'timeline';
let currentWorldSidebarSort = 'count';
let currentPhotoSortOrder = 'desc';
let currentPhotoCardDensity = 'default';
let currentModalPhoto = null;
let imageModalAnimationTimer = null;
let imageModalSwitchTimer = null;
let photoCardDensityAnimationTimer = null;
let modalShellRestoreTimer = null;
let modalWorldMetadataRequestId = 0;
let modalPhotoLabelsRequestId = 0;
let modalImageRecoveryRequestId = 0;
let trackedFolders = [];
let settingsOverviewSection = null;
let settingsOverviewGrid = null;
let settingsBackgroundSection = null;
let settingsBackgroundMeta = null;
let selectBackgroundImageButton = null;
let clearBackgroundImageButton = null;
let settingsUtilityActionsStack = null;
let regenerateThumbnailMonthSelect = null;
let regenerateThumbnailMonthDropdown = null;
let regenerateThumbnailMonthButton = null;
let regenerateThumbnailMonthLabel = null;
let regenerateThumbnailMonthMenu = null;
let regenerateThumbnailMonthValue = '';
let isRegenerateThumbnailMonthMenuOpen = false;
let regenerateThumbnailMonthMenuCloseTimer = null;
let reimportRegisteredPhotoMonthSelect = null;
let reimportRegisteredPhotoMonthDropdown = null;
let reimportRegisteredPhotoMonthButton = null;
let reimportRegisteredPhotoMonthLabel = null;
let reimportRegisteredPhotoMonthMenu = null;
let reimportRegisteredPhotoMonthValue = '';
let isReimportRegisteredPhotoMonthMenuOpen = false;
let reimportRegisteredPhotoMonthMenuCloseTimer = null;
let trackedFolderSettingsSection = null;
let trackedFolderSettingsActions = null;
let openTrackedFolderListButton = null;
let trackedFolderModal = null;
let trackedFolderModalBackdrop = null;
let trackedFolderModalClose = null;
let trackedFolderModalBody = null;
let worldLibraryModeButton = null;
let sidebarHeaderControls = null;
let sidebarSortCountButton = null;
let sidebarSortNameButton = null;
let sidebarHeaderControlsHideTimer = null;
let modalResolutionHeroBadge = null;
let modalPrintNoteHeroBadge = null;
let modalTakenAtHero = null;
let imageModalPrevButton = null;
let imageModalNextButton = null;
let modalPhotoLabelsBlock = null;
let modalPhotoLabelsList = null;
let modalPrintNoteBlock = null;
let modalPrintNoteValue = null;
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
let isFavoriteFilterOnly = false;
let activeOrientationFilter = 'all';
let isOrientationFilterMenuOpen = false;
let orientationFilterMenuCloseTimer = null;
let activePhotoLabelFilters = [];
let photoLabelFilterMode = 'or';
let isPhotoLabelFilterMenuOpen = false;
let photoLabelFilterMenuCloseTimer = null;
let activeToolbarSearchScope = 'world';
let draftToolbarSearchScope = 'world';
let activeWorldNameFilter = '';
let isWorldNameFilterMenuOpen = false;
let worldNameFilterMenuCloseTimer = null;
let worldNameFilterInputTimer = null;
let toolbarSearchScopeDropdown = null;
let toolbarSearchScopeButton = null;
let toolbarSearchScopeMenu = null;
let toolbarSearchClearButton = null;
let isToolbarSearchScopeMenuOpen = false;
let toolbarSearchScopeMenuCloseTimer = null;
let isSelectionMode = false;
const selectedPhotoIds = new Set();
let lastSelectionAnchorPhotoId = null;
let isSelectionDragActive = false;
let selectionDragTargetState = null;
let selectionDragPointerId = null;
let suppressSelectionModeCardClickPhotoId = null;
let lastSelectionDragPhotoId = null;
let keyboardFocusedPhotoId = null;
let isImporting = false;
let isWorldMetadataSyncing = false;
let worldMetadataSyncResetTimer = null;
let appUpdatePromptQueue = Promise.resolve();
let lastTimelineSelection = null;
let lastWorldSelection = null;

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
const TOOLBAR_SEARCH_SCOPE_META = {
  world: {
    label: 'World',
    buttonLabel: 'World',
    placeholder: 'World名を入力',
    summaryPrefix: 'World',
  },
  memo: {
    label: 'メモ',
    buttonLabel: 'メモ',
    placeholder: 'メモを入力',
    summaryPrefix: 'メモ',
  },
  printNote: {
    label: 'プリントのノート',
    buttonLabel: 'プリント',
    placeholder: 'プリントのノートを入力',
    summaryPrefix: 'プリントのノート',
  },
};

function pad2(value) {
  return String(value).padStart(2, '0');
}

// Sidebar selection can now point at either a specific month or a whole year.
// Keep the shape normalized so render, restore, and maintenance code can all
// branch on the same mode field instead of inferring intent ad hoc.
function createMonthSelection(year, month) {
  return {
    mode: 'month',
    year,
    month,
  };
}

function createYearSelection(year) {
  return {
    mode: 'year',
    year,
    month: null,
  };
}

function createWorldSelection(worldKey, worldName, worldId = null) {
  return {
    mode: 'world',
    worldKey,
    worldName,
    worldId:
      typeof worldId === 'string' && worldId.trim().length > 0
        ? worldId.trim()
        : null,
  };
}

function normalizeSelection(selection) {
  if (selection?.mode === 'world') {
    const normalizedWorldKey =
      typeof selection.worldKey === 'string' && selection.worldKey.trim().length > 0
        ? selection.worldKey.trim()
        : null;
    const normalizedWorldName =
      typeof selection.worldName === 'string' && selection.worldName.trim().length > 0
        ? selection.worldName.trim()
        : null;
    const normalizedWorldId =
      typeof selection.worldId === 'string' && selection.worldId.trim().length > 0
        ? selection.worldId.trim()
        : null;

    if (!normalizedWorldKey || !normalizedWorldName) {
      return null;
    }

    return createWorldSelection(
      normalizedWorldKey,
      normalizedWorldName,
      normalizedWorldId
    );
  }

  const normalizedYear = Number(selection?.year);
  const hasExplicitMonth =
    selection &&
    selection.month !== null &&
    selection.month !== undefined &&
    String(selection.month).trim() !== '';
  const normalizedMonth = hasExplicitMonth ? Number(selection?.month) : null;

  if (!Number.isInteger(normalizedYear)) {
    return null;
  }

  if (Number.isInteger(normalizedMonth)) {
    return createMonthSelection(normalizedYear, normalizedMonth);
  }

  return createYearSelection(normalizedYear);
}

function isMonthSelection(selection = currentSelection) {
  return normalizeSelection(selection)?.mode === 'month';
}

function isYearSelection(selection = currentSelection) {
  return normalizeSelection(selection)?.mode === 'year';
}

function isWorldSelection(selection = currentSelection) {
  return normalizeSelection(selection)?.mode === 'world';
}

function isSameSelection(leftSelection, rightSelection) {
  const left = normalizeSelection(leftSelection);
  const right = normalizeSelection(rightSelection);

  if (!left || !right) {
    return false;
  }

  if (left.mode === 'world' || right.mode === 'world') {
    return left.mode === right.mode && left.worldKey === right.worldKey;
  }

  return (
    left.mode === right.mode &&
    left.year === right.year &&
    left.month === right.month
  );
}

function getSelectionLabelText(selection = currentSelection) {
  const normalizedSelection = normalizeSelection(selection);

  if (!normalizedSelection) {
    return '写真一覧';
  }

  if (normalizedSelection.mode === 'world') {
    return normalizedSelection.worldName;
  }

  if (normalizedSelection.mode === 'year') {
    return String(normalizedSelection.year);
  }

  return `${normalizedSelection.year}/${pad2(normalizedSelection.month)}`;
}

function getDefaultSelectionEmptyMessage(selection = currentSelection) {
  const normalizedSelection = normalizeSelection(selection);

  if (!normalizedSelection) {
    return '表示する年または月を選択してください';
  }

  if (normalizedSelection.mode === 'world') {
    return 'このワールドの写真はまだありません';
  }

  return normalizedSelection.mode === 'year'
    ? 'この年の写真はまだありません'
    : 'この月の写真はまだありません';
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

function parsePhotoSortTimestamp(value) {
  if (typeof value !== 'string' || value.trim().length === 0) {
    return null;
  }

  const match = value
    .trim()
    .match(
      /^(\d{4})[/-](\d{2})[/-](\d{2})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/
    );

  if (!match) {
    return null;
  }

  const [
    ,
    yearText,
    monthText,
    dayText,
    hourText = '00',
    minuteText = '00',
    secondText = '00',
  ] = match;

  const timestamp = new Date(
    Number(yearText),
    Number(monthText) - 1,
    Number(dayText),
    Number(hourText),
    Number(minuteText),
    Number(secondText)
  ).getTime();

  return Number.isFinite(timestamp) ? timestamp : null;
}

function getPhotoSortTimestamp(photo) {
  return (
    parsePhotoSortTimestamp(photo?.takenAt) ??
    parsePhotoSortTimestamp(photo?.groupDate) ??
    0
  );
}

function comparePhotosForCurrentSortOrder(leftPhoto, rightPhoto) {
  const timestampDiff =
    getPhotoSortTimestamp(leftPhoto) - getPhotoSortTimestamp(rightPhoto);

  if (timestampDiff !== 0) {
    return currentPhotoSortOrder === 'asc' ? timestampDiff : -timestampDiff;
  }

  const idDiff = Number(leftPhoto?.id || 0) - Number(rightPhoto?.id || 0);

  if (idDiff !== 0) {
    return currentPhotoSortOrder === 'asc' ? idDiff : -idDiff;
  }

  return String(leftPhoto?.fileName || '').localeCompare(
    String(rightPhoto?.fileName || ''),
    'ja'
  );
}

function sortPhotosForCurrentSortOrder(photos) {
  return [...(Array.isArray(photos) ? photos : [])].sort(
    comparePhotosForCurrentSortOrder
  );
}

function setCurrentMonthPhotos(nextPhotos) {
  allCurrentMonthPhotos = sortPhotosForCurrentSortOrder(nextPhotos);
  applyCurrentPhotoFilter();
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

// Card density is a visual-only preference. We preserve the existing card DOM
// and let CSS decide how much metadata stays visible in compact mode.
function applyPhotoCardDensityPreference(density) {
  currentPhotoCardDensity = density === 'compact' ? 'compact' : 'default';
  document.body.classList.toggle(
    'compact-card-view',
    currentPhotoCardDensity === 'compact'
  );
  localStorage.setItem(PHOTO_CARD_DENSITY_STORAGE_KEY, currentPhotoCardDensity);
}

function syncPhotoCardDensityUi() {
  if (!photoDensityButton || !photoDensityIcon) {
    return;
  }

  const isCompact = currentPhotoCardDensity === 'compact';
  photoDensityButton.classList.toggle('is-active', isCompact);
  photoDensityButton.setAttribute(
    'aria-label',
    isCompact ? '表示サイズ: コンパクト' : '表示サイズ: 標準'
  );
  photoDensityButton.title = isCompact
    ? '表示サイズ: コンパクト'
    : '表示サイズ: 標準';
  photoDensityIcon.textContent = isCompact
    ? 'view_compact_alt'
    : 'view_comfy_alt';
}

function initializePhotoCardDensityPreference() {
  const savedDensity = localStorage.getItem(PHOTO_CARD_DENSITY_STORAGE_KEY);
  applyPhotoCardDensityPreference(savedDensity === 'compact' ? 'compact' : 'default');
}

function playPhotoCardDensityTransition() {
  if (photoCardDensityAnimationTimer) {
    clearTimeout(photoCardDensityAnimationTimer);
  }

  document.body.classList.add('is-density-switching');
  photoCardDensityAnimationTimer = setTimeout(() => {
    document.body.classList.remove('is-density-switching');
    photoCardDensityAnimationTimer = null;
  }, 320);
}

function normalizeBackgroundImagePath(filePath) {
  return typeof filePath === 'string' ? filePath.trim() : '';
}

function getStoredBackgroundImagePath() {
  return normalizeBackgroundImagePath(
    localStorage.getItem(BACKGROUND_IMAGE_STORAGE_KEY)
  );
}

function cacheBackgroundImagePath(filePath) {
  const normalizedPath = normalizeBackgroundImagePath(filePath);

  if (!normalizedPath) {
    localStorage.removeItem(BACKGROUND_IMAGE_STORAGE_KEY);
    return '';
  }

  localStorage.setItem(BACKGROUND_IMAGE_STORAGE_KEY, normalizedPath);
  return normalizedPath;
}

function buildBackgroundImageFileUrl(filePath) {
  const normalizedPath = normalizeBackgroundImagePath(filePath);

  if (!normalizedPath) {
    return '';
  }

  const slashPath = normalizedPath.replace(/\\/g, '/');
  const prefixedPath = slashPath.startsWith('/') ? slashPath : `/${slashPath}`;
  return encodeURI(`file://${prefixedPath}`);
}

function renderBackgroundImagePreference(filePath) {
  const normalizedPath = normalizeBackgroundImagePath(filePath);

  if (!normalizedPath) {
    document.body.classList.remove('has-custom-background');
    document.body.style.setProperty('--app-background-image', 'none');
    return;
  }

  document.body.classList.add('has-custom-background');
  document.body.style.setProperty(
    '--app-background-image',
    `url("${buildBackgroundImageFileUrl(normalizedPath)}")`
  );
}

async function applyBackgroundImagePreference(filePath, options = {}) {
  const { persist = true } = options;
  const normalizedPath = normalizeBackgroundImagePath(filePath);

  renderBackgroundImagePreference(normalizedPath);
  cacheBackgroundImagePath(normalizedPath);

  if (!persist || !window.electronAPI.setBackgroundImagePreference) {
    return normalizedPath;
  }

  try {
    const result = await window.electronAPI.setBackgroundImagePreference(
      normalizedPath
    );
    const savedPath = normalizeBackgroundImagePath(result?.filePath);
    renderBackgroundImagePreference(savedPath);
    cacheBackgroundImagePath(savedPath);
    return savedPath;
  } catch {
    return normalizedPath;
  }
}

async function initializeBackgroundImagePreference() {
  const cachedPath = getStoredBackgroundImagePath();
  renderBackgroundImagePreference(cachedPath);

  if (!window.electronAPI.getBackgroundImagePreference) {
    cacheBackgroundImagePath(cachedPath);
    return;
  }

  try {
    const result = await window.electronAPI.getBackgroundImagePreference();
    const persistedPath = normalizeBackgroundImagePath(result?.filePath);
    renderBackgroundImagePreference(persistedPath);
    cacheBackgroundImagePath(persistedPath);
  } catch {
    cacheBackgroundImagePath(cachedPath);
  }

  syncSettingsBackgroundUi();
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
    `新着${result.newCount}件`,
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

function hasSelectableMonthOptions(selectElement) {
  if (!selectElement || selectElement.disabled) {
    return false;
  }

  return Array.from(selectElement.options || []).some(
    (option) => typeof option.value === 'string' && option.value.trim().length > 0
  );
}

function renderRegenerateThumbnailMonthOptions() {
  if (!regenerateThumbnailMonthSelect) {
    return;
  }

  const monthOptions = getSidebarMonthOptions();
  const preferredValue =
    isMonthSelection(currentSelection) &&
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
  const hasOptions = hasSelectableMonthOptions(regenerateThumbnailMonthSelect);
  const nextOpen = Boolean(isOpen) && !isImporting && hasOptions;
  if (nextOpen) {
    closeManagedDropdownsExcept(regenerateThumbnailMonthDropdown);
  }
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

  if (!regenerateThumbnailMonthValue) {
    regenerateThumbnailMonthButton.disabled = true;
    closeRegenerateThumbnailMonthMenu();
    return;
  }

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

function renderReimportRegisteredPhotoMonthOptions() {
  if (!reimportRegisteredPhotoMonthSelect) {
    return;
  }

  const monthOptions = getSidebarMonthOptions();
  const preferredValue =
    isMonthSelection(currentSelection) &&
    monthOptions.some(
      (option) =>
        option.year === currentSelection.year &&
        option.month === currentSelection.month
    )
      ? `${currentSelection.year}-${pad2(currentSelection.month)}`
      : reimportRegisteredPhotoMonthSelect.value;

  reimportRegisteredPhotoMonthSelect.innerHTML = '';

  if (monthOptions.length === 0) {
    const emptyOption = document.createElement('option');
    emptyOption.value = '';
    emptyOption.textContent = '対象月がありません';
    reimportRegisteredPhotoMonthSelect.appendChild(emptyOption);
    reimportRegisteredPhotoMonthSelect.disabled = true;
    syncReimportRegisteredPhotoMonthDropdownFromSelect();
    return;
  }

  monthOptions.forEach((option) => {
    const selectOption = document.createElement('option');
    selectOption.value = `${option.year}-${pad2(option.month)}`;
    selectOption.textContent = `${option.year}年${option.month}月 (${option.count}枚)`;
    reimportRegisteredPhotoMonthSelect.appendChild(selectOption);
  });

  reimportRegisteredPhotoMonthSelect.disabled = false;

  if (
    typeof preferredValue === 'string' &&
    preferredValue.length > 0 &&
    monthOptions.some(
      (option) => `${option.year}-${pad2(option.month)}` === preferredValue
    )
  ) {
    reimportRegisteredPhotoMonthSelect.value = preferredValue;
  }

  syncReimportRegisteredPhotoMonthDropdownFromSelect();
}

function setReimportRegisteredPhotoMonthMenuOpen(isOpen) {
  const hasOptions = hasSelectableMonthOptions(reimportRegisteredPhotoMonthSelect);
  const nextOpen = Boolean(isOpen) && !isImporting && hasOptions;
  if (nextOpen) {
    closeManagedDropdownsExcept(reimportRegisteredPhotoMonthDropdown);
  }
  isReimportRegisteredPhotoMonthMenuOpen = nextOpen;
  setAnimatedDropdownOpenState({
    dropdown: reimportRegisteredPhotoMonthDropdown,
    button: reimportRegisteredPhotoMonthButton,
    menu: reimportRegisteredPhotoMonthMenu,
    isOpen: nextOpen,
    closeTimerRef: {
      get current() {
        return reimportRegisteredPhotoMonthMenuCloseTimer;
      },
      set current(value) {
        reimportRegisteredPhotoMonthMenuCloseTimer = value;
      },
    },
  });
}

function closeReimportRegisteredPhotoMonthMenu() {
  setReimportRegisteredPhotoMonthMenuOpen(false);
}

function syncReimportRegisteredPhotoMonthDropdownFromSelect() {
  if (
    !reimportRegisteredPhotoMonthSelect ||
    !reimportRegisteredPhotoMonthDropdown ||
    !reimportRegisteredPhotoMonthButton ||
    !reimportRegisteredPhotoMonthLabel ||
    !reimportRegisteredPhotoMonthMenu
  ) {
    return;
  }

  const options = Array.from(reimportRegisteredPhotoMonthSelect.options);
  reimportRegisteredPhotoMonthValue = reimportRegisteredPhotoMonthSelect.value || '';
  reimportRegisteredPhotoMonthMenu.innerHTML = '';

  if (options.length === 0) {
    reimportRegisteredPhotoMonthLabel.textContent = '対象月がありません';
    reimportRegisteredPhotoMonthButton.disabled = true;
    closeReimportRegisteredPhotoMonthMenu();
    return;
  }

  const selectedOption =
    options.find((option) => option.value === reimportRegisteredPhotoMonthValue) ||
    options[0];

  reimportRegisteredPhotoMonthValue = selectedOption?.value || '';
  reimportRegisteredPhotoMonthSelect.value = reimportRegisteredPhotoMonthValue;
  reimportRegisteredPhotoMonthLabel.textContent =
    selectedOption?.textContent || '再取り込みする月を選択';

  if (!reimportRegisteredPhotoMonthValue) {
    reimportRegisteredPhotoMonthButton.disabled = true;
    closeReimportRegisteredPhotoMonthMenu();
    return;
  }

  reimportRegisteredPhotoMonthButton.disabled = false;

  options.forEach((option) => {
    const isActive = option.value === reimportRegisteredPhotoMonthValue;
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'header-dropdown-item header-dropdown-item-with-meta';
    item.classList.toggle('is-active', isActive);
    item.dataset.reimportRegisteredPhotoMonth = option.value;
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
    reimportRegisteredPhotoMonthMenu.appendChild(item);
  });
}

function syncSettingsUtilityActionsUi() {
  const hasThumbnailMonthOptions = hasSelectableMonthOptions(
    regenerateThumbnailMonthSelect
  );
  const hasReimportMonthOptions = hasSelectableMonthOptions(
    reimportRegisteredPhotoMonthSelect
  );

  if (regenerateThumbnailMonthSelect) {
    regenerateThumbnailMonthSelect.disabled =
      isImporting || !hasThumbnailMonthOptions;
  }

  if (regenerateThumbnailMonthButton) {
    regenerateThumbnailMonthButton.disabled =
      isImporting || !hasThumbnailMonthOptions;
    regenerateThumbnailMonthButton.setAttribute(
      'title',
      hasThumbnailMonthOptions
        ? 'サムネイルを再生成する月を選択'
        : '対象月がありません'
    );
  }

  if (regenerateThumbnailsButton) {
    regenerateThumbnailsButton.disabled =
      isImporting || !hasThumbnailMonthOptions;
    regenerateThumbnailsButton.setAttribute(
      'title',
      hasThumbnailMonthOptions
        ? '選択中の月のサムネイルを再生成'
        : '再生成できる月がありません'
    );
  }

  if (reimportRegisteredPhotoMonthSelect) {
    reimportRegisteredPhotoMonthSelect.disabled =
      isImporting || !hasReimportMonthOptions;
  }

  if (reimportRegisteredPhotoMonthButton) {
    reimportRegisteredPhotoMonthButton.disabled =
      isImporting || !hasReimportMonthOptions;
    reimportRegisteredPhotoMonthButton.setAttribute(
      'title',
      hasReimportMonthOptions
        ? '情報を再取り込みする月を選択'
        : '対象月がありません'
    );
  }

  if (reimportRegisteredPhotosButton) {
    reimportRegisteredPhotosButton.disabled =
      isImporting || !hasReimportMonthOptions;
    reimportRegisteredPhotosButton.setAttribute(
      'title',
      hasReimportMonthOptions
        ? '選択中の月の登録画像を現在の解析ロジックで再取り込み'
        : '再取り込みできる月がありません'
    );
  }
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
    `新着${result.newCount || 0}件`,
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
  currentPhotos = sortPhotosForCurrentSortOrder(
    allCurrentMonthPhotos.filter(photoMatchesCurrentFilters)
  );
}

function getOrientationFilterMeta(filterValue) {
  return ORIENTATION_FILTER_META[filterValue] || ORIENTATION_FILTER_META.all;
}

function setOrientationFilterMenuOpen(isOpen) {
  const nextOpen = Boolean(isOpen) && !isImporting && Boolean(currentSelection);
  if (nextOpen) {
    closeManagedDropdownsExcept(orientationFilterDropdown);
  }
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

function openTrackedFolderModal() {
  if (!trackedFolderModal) {
    return;
  }

  openSubModalElement(trackedFolderModal);

  if (trackedFolderList) {
    trackedFolderList.scrollTop = 0;
  }
}

function closeTrackedFolderModal() {
  closeSubModalElement(trackedFolderModal, {
    onClosed: () => {
      if (trackedFolderList) {
        trackedFolderList.scrollTop = 0;
      }
    },
  });
}

function ensureSettingsUtilityActionsStack() {
  if (!settingsMaintenanceSection) {
    return null;
  }

  let utilityActionsStack = settingsMaintenanceSection.querySelector(
    '.settings-utility-actions-stack'
  );
  const utilityAnchor = settingsMaintenanceActions;

  if (!utilityActionsStack) {
    utilityActionsStack = document.createElement('div');
    utilityActionsStack.className = 'settings-utility-actions-stack';
    settingsMaintenanceSection.insertBefore(
      utilityActionsStack,
      utilityAnchor || null
    );
  } else if (
    utilityAnchor &&
    utilityActionsStack.nextElementSibling !== utilityAnchor
  ) {
    settingsMaintenanceSection.insertBefore(utilityActionsStack, utilityAnchor);
  }

  settingsUtilityActionsStack = utilityActionsStack;
  return utilityActionsStack;
}

function ensureSettingsUtilityActionsContainer(groupName = 'thumbnails') {
  const utilityActionsStack = ensureSettingsUtilityActionsStack();

  if (!utilityActionsStack) {
    return null;
  }

  const normalizedGroupName =
    typeof groupName === 'string' && groupName.trim().length > 0
      ? groupName.trim()
      : 'thumbnails';

  let utilityActions = Array.from(utilityActionsStack.children).find(
    (element) => element.dataset?.settingsUtilityGroup === normalizedGroupName
  );

  if (!utilityActions) {
    utilityActions = document.createElement('div');
    utilityActions.className = 'settings-utility-actions';
    utilityActions.dataset.settingsUtilityGroup = normalizedGroupName;
    utilityActionsStack.appendChild(utilityActions);
  }

  return utilityActions;
}

function ensureSettingsOverviewSection() {
  if (!settingsModalBody) {
    return null;
  }

  if (!settingsOverviewSection) {
    const section = document.createElement('div');
    section.className = 'settings-section settings-overview-section';

    const title = document.createElement('p');
    title.className = 'settings-section-title';
    title.textContent = '概要';
    section.appendChild(title);

    const grid = document.createElement('div');
    grid.className = 'settings-overview-grid';
    section.appendChild(grid);

    const firstSettingsSection = settingsModalBody.querySelector('.settings-section');

    settingsModalBody.insertBefore(section, firstSettingsSection || null);
    settingsOverviewSection = section;
    settingsOverviewGrid = grid;
  }

  return settingsOverviewSection;
}

function renderSettingsOverview(summary) {
  if (!settingsOverviewGrid) {
    return;
  }

  const normalizedSummary = {
    photoCount: Number(summary?.photoCount) || 0,
    trackedFolderCount: Number(summary?.trackedFolderCount) || 0,
    worldCacheCount: Number(summary?.worldCacheCount) || 0,
    tagCount: Number(summary?.tagCount) || 0,
  };

  const cards = [
    { label: '写真', value: normalizedSummary.photoCount },
    { label: 'フォルダ', value: normalizedSummary.trackedFolderCount },
    { label: 'ワールド数', value: normalizedSummary.worldCacheCount },
    { label: 'ラベル', value: normalizedSummary.tagCount },
  ];

  settingsOverviewGrid.innerHTML = cards
    .map(
      (item) => `
        <div class="settings-overview-card">
          <p class="settings-overview-label">${escapeHtml(item.label)}</p>
          <p class="settings-overview-value">${item.value.toLocaleString('ja-JP')}</p>
        </div>
      `
    )
    .join('');
}

async function loadSettingsOverview() {
  ensureSettingsOverviewSection();

  if (!settingsOverviewGrid || !window.electronAPI.getApplicationDataSummary) {
    return;
  }

  renderSettingsOverview({
    photoCount: 0,
    trackedFolderCount: trackedFolders.length,
    worldCacheCount: 0,
    tagCount: 0,
  });

  try {
    const summary = await window.electronAPI.getApplicationDataSummary();
    renderSettingsOverview(summary);
  } catch {
    renderSettingsOverview({
      photoCount: sidebarData.reduce(
        (sum, year) => sum + (Number(year.totalCount) || 0),
        0
      ),
      trackedFolderCount: trackedFolders.length,
      worldCacheCount: 0,
      tagCount: 0,
    });
  }
}

function ensureSettingsBackgroundSection() {
  if (!settingsModalBody || !settingsFontSection) {
    return null;
  }

  if (!settingsBackgroundSection) {
    const section = document.createElement('div');
    section.className = 'settings-section settings-background-section';

    const title = document.createElement('p');
    title.className = 'settings-section-title';
    title.textContent = '背景';
    section.appendChild(title);

    const meta = document.createElement('p');
    meta.className = 'settings-background-meta';
    section.appendChild(meta);

    const actions = document.createElement('div');
    actions.className = 'settings-background-actions';

    const selectButton = document.createElement('button');
    selectButton.type = 'button';
    selectButton.className = 'small-action-button';
    selectButton.textContent = '画像を選択';
    actions.appendChild(selectButton);

    const clearButton = document.createElement('button');
    clearButton.type = 'button';
    clearButton.className = 'small-action-button secondary';
    clearButton.textContent = 'クリア';
    actions.appendChild(clearButton);

    section.appendChild(actions);
    settingsFontSection.insertAdjacentElement('afterend', section);

    settingsBackgroundSection = section;
    settingsBackgroundMeta = meta;
    selectBackgroundImageButton = selectButton;
    clearBackgroundImageButton = clearButton;
  }

  return settingsBackgroundSection;
}

function syncSettingsBackgroundUi() {
  ensureSettingsBackgroundSection();

  if (!settingsBackgroundMeta) {
    return;
  }

  const currentPath = getStoredBackgroundImagePath();
  const fileName = currentPath
    ? currentPath.split(/[\\/]/).filter(Boolean).pop() || currentPath
    : '';

  settingsBackgroundMeta.textContent = fileName || '未設定';

  if (selectBackgroundImageButton) {
    selectBackgroundImageButton.disabled = isImporting;
  }

  if (clearBackgroundImageButton) {
    clearBackgroundImageButton.disabled = isImporting || !currentPath;
  }
}

async function selectBackgroundImageFromSettings() {
  if (!window.electronAPI.selectBackgroundImage) {
    return;
  }

  const result = await window.electronAPI.selectBackgroundImage();

  if (!result?.ok || result.canceled) {
    return;
  }

  await applyBackgroundImagePreference(result.filePath);
  syncSettingsBackgroundUi();
  showToast('背景画像を更新しました');
}

async function clearBackgroundImageFromSettings() {
  await applyBackgroundImagePreference('');
  syncSettingsBackgroundUi();
  showToast('背景画像をクリアしました');
}

// Settings modal keeps tracked folder management lightweight by showing only
// entry points inline and moving the actual list into its own sub-modal.
function initializeSettingsTrackedFolderUi() {
  if (!settingsModalBody || !addTrackedFolderButton || !trackedFolderList) {
    return;
  }

  if (clearThumbnailCacheButton) {
    clearThumbnailCacheButton.classList.remove('secondary');
    clearThumbnailCacheButton.classList.add('danger-button');
  }

  if (!trackedFolderSettingsSection) {
    trackedFolderSettingsSection = addTrackedFolderButton.closest('.settings-section');
  }

  if (!trackedFolderSettingsSection) {
    return;
  }

  trackedFolderSettingsSection.classList.add('tracked-folder-settings-section');

  if (!trackedFolderSettingsMeta) {
    trackedFolderSettingsMeta = document.createElement('p');
    trackedFolderSettingsMeta.className = 'settings-section-meta tracked-folder-settings-meta';
    const trackedFolderInsertBefore =
      trackedFolderSettingsSection.querySelector('.tracked-folder-settings-actions') ||
      trackedFolderSettingsSection.querySelector('.settings-section-header')?.nextElementSibling ||
      trackedFolderList;
    trackedFolderSettingsSection.insertBefore(
      trackedFolderSettingsMeta,
      trackedFolderInsertBefore
    );
  }

  if (!trackedFolderSettingsActions) {
    trackedFolderSettingsActions = document.createElement('div');
    trackedFolderSettingsActions.className = 'tracked-folder-settings-actions';
    trackedFolderSettingsSection.appendChild(trackedFolderSettingsActions);
  }

  if (!openTrackedFolderListButton) {
    openTrackedFolderListButton = document.createElement('button');
    openTrackedFolderListButton.type = 'button';
    openTrackedFolderListButton.className = 'small-action-button secondary';
    openTrackedFolderListButton.textContent = '一覧を表示';
    trackedFolderSettingsActions.appendChild(openTrackedFolderListButton);
  }

  addTrackedFolderButton.textContent = 'フォルダ追加';
  trackedFolderSettingsActions.appendChild(addTrackedFolderButton);
  syncTrackedFolderSettingsMeta();
  syncTrackedFolderSettingsActionsUi();

  if (!trackedFolderModal) {
    trackedFolderModal = document.createElement('div');
    trackedFolderModal.id = 'tracked-folder-modal';
    trackedFolderModal.className = 'sub-modal hidden';

    trackedFolderModalBackdrop = document.createElement('div');
    trackedFolderModalBackdrop.className = 'sub-modal-backdrop';
    trackedFolderModal.appendChild(trackedFolderModalBackdrop);

    const content = document.createElement('div');
    content.className = 'sub-modal-content tracked-folder-modal-content';
    trackedFolderModal.appendChild(content);

    trackedFolderModalClose = document.createElement('button');
    trackedFolderModalClose.type = 'button';
    trackedFolderModalClose.className = 'sub-modal-close';
    trackedFolderModalClose.setAttribute('aria-label', '更新対象フォルダ一覧を閉じる');
    const closeIcon = document.createElement('span');
    closeIcon.className = 'material-symbols-outlined';
    closeIcon.textContent = 'close';
    trackedFolderModalClose.appendChild(closeIcon);
    content.appendChild(trackedFolderModalClose);

    trackedFolderModalBody = document.createElement('div');
    trackedFolderModalBody.className = 'sub-modal-body tracked-folder-modal-body';
    content.appendChild(trackedFolderModalBody);

    const title = document.createElement('h3');
    title.textContent = '更新対象フォルダ一覧';
    trackedFolderModalBody.appendChild(title);

    document.body.appendChild(trackedFolderModal);
  }

  if (trackedFolderModalBody && trackedFolderList.parentElement !== trackedFolderModalBody) {
    trackedFolderModalBody.appendChild(trackedFolderList);
  }
}

function ensureRegenerateThumbnailMonthDropdown(utilityActions) {
  if (!utilityActions) {
    return;
  }

  if (!regenerateThumbnailMonthSelect) {
    const monthSelect = document.createElement('select');
    monthSelect.className = 'settings-month-select';
    monthSelect.setAttribute('aria-label', 'サムネイル再生成の対象月');
    utilityActions.appendChild(monthSelect);
    regenerateThumbnailMonthSelect = monthSelect;
  }

  if (regenerateThumbnailMonthDropdown) {
    return;
  }

  const monthDropdown = document.createElement('div');
  monthDropdown.className = 'header-dropdown settings-month-dropdown';

  const monthButton = document.createElement('button');
  monthButton.type = 'button';
  monthButton.className = 'header-filter-button settings-month-dropdown-button';
  monthButton.setAttribute('aria-haspopup', 'menu');
  monthButton.setAttribute('aria-expanded', 'false');
  monthButton.setAttribute('aria-label', 'サムネイル再生成の対象月');

  const monthLabel = document.createElement('span');
  monthLabel.className = 'settings-month-dropdown-label';
  monthLabel.textContent = '再生成する月を選択';
  monthButton.appendChild(monthLabel);

  const chevron = document.createElement('span');
  chevron.className = 'material-symbols-outlined orientation-filter-chevron';
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

function ensureReimportRegisteredPhotoMonthDropdown(utilityActions) {
  if (!utilityActions) {
    return;
  }

  if (!reimportRegisteredPhotoMonthSelect) {
    const monthSelect = document.createElement('select');
    monthSelect.className = 'settings-month-select';
    monthSelect.setAttribute('aria-label', '情報再取り込みの対象月');
    utilityActions.appendChild(monthSelect);
    reimportRegisteredPhotoMonthSelect = monthSelect;
  }

  if (reimportRegisteredPhotoMonthDropdown) {
    return;
  }

  const monthDropdown = document.createElement('div');
  monthDropdown.className = 'header-dropdown settings-month-dropdown';

  const monthButton = document.createElement('button');
  monthButton.type = 'button';
  monthButton.className = 'header-filter-button settings-month-dropdown-button';
  monthButton.setAttribute('aria-haspopup', 'menu');
  monthButton.setAttribute('aria-expanded', 'false');
  monthButton.setAttribute('aria-label', '情報再取り込みの対象月');

  const monthLabel = document.createElement('span');
  monthLabel.className = 'settings-month-dropdown-label';
  monthLabel.textContent = '再取り込みする月を選択';
  monthButton.appendChild(monthLabel);

  const chevron = document.createElement('span');
  chevron.className = 'material-symbols-outlined orientation-filter-chevron';
  chevron.textContent = 'expand_more';
  monthButton.appendChild(chevron);

  const monthMenu = document.createElement('div');
  monthMenu.className = 'header-dropdown-menu settings-month-dropdown-menu';
  monthMenu.hidden = true;
  monthMenu.setAttribute('role', 'menu');

  monthDropdown.appendChild(monthButton);
  monthDropdown.appendChild(monthMenu);
  utilityActions.appendChild(monthDropdown);

  reimportRegisteredPhotoMonthDropdown = monthDropdown;
  reimportRegisteredPhotoMonthButton = monthButton;
  reimportRegisteredPhotoMonthLabel = monthLabel;
  reimportRegisteredPhotoMonthMenu = monthMenu;

  monthButton.addEventListener('click', (event) => {
    event.stopPropagation();
    setReimportRegisteredPhotoMonthMenuOpen(
      !isReimportRegisteredPhotoMonthMenuOpen
    );
  });

  monthMenu.addEventListener('click', (event) => {
    const target = event.target.closest('[data-reimport-registered-photo-month]');

    if (!target || !reimportRegisteredPhotoMonthSelect) {
      return;
    }

    event.stopPropagation();
    reimportRegisteredPhotoMonthValue =
      target.dataset.reimportRegisteredPhotoMonth || '';
    reimportRegisteredPhotoMonthSelect.value = reimportRegisteredPhotoMonthValue;
    syncReimportRegisteredPhotoMonthDropdownFromSelect();
    closeReimportRegisteredPhotoMonthMenu();
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
    modeButton.textContent = mode === 'or' ? 'OR' : 'AND';
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
  if (nextOpen) {
    closeManagedDropdownsExcept(photoLabelFilterDropdown);
  }
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

function getToolbarSearchScopeMeta(scope = draftToolbarSearchScope) {
  return TOOLBAR_SEARCH_SCOPE_META[scope] || TOOLBAR_SEARCH_SCOPE_META.world;
}

function getNormalizedWorldNameFilterText(value) {
  return normalizeWorldNameFilterText(value)
    .normalize('NFKC')
    .toLocaleLowerCase('ja-JP');
}

function getWorldNameFilterSummaryText({ includePrefix = true } = {}) {
  const scopeMeta = getToolbarSearchScopeMeta(activeToolbarSearchScope);

  if (!activeWorldNameFilter) {
    return includePrefix ? `${scopeMeta.summaryPrefix}: すべて` : 'すべて';
  }

  return includePrefix
    ? `${scopeMeta.summaryPrefix}: ${activeWorldNameFilter}`
    : activeWorldNameFilter;
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

function isStaticToolbarWorldFilter() {
  return worldNameFilterDropdown?.classList.contains('is-static-toolbar-filter');
}

function ensureStaticToolbarWorldFilterVisible() {
  if (!isStaticToolbarWorldFilter() || !worldNameFilterMenu) {
    return;
  }

  worldNameFilterMenu.hidden = false;
  worldNameFilterDropdown.classList.remove('is-open');
  worldNameFilterButton?.setAttribute('aria-expanded', 'false');
}

function syncToolbarSearchInputUi() {
  const scopeMeta = getToolbarSearchScopeMeta();
  const selectionDependentDisabled = isImporting || !currentSelection;
  const hasDraftSearch =
    normalizeWorldNameFilterText(worldNameFilterInput?.value || '').length > 0;
  const hasActiveSearch = activeWorldNameFilter.length > 0;

  if (worldNameFilterInput) {
    worldNameFilterInput.placeholder = scopeMeta.placeholder;
    worldNameFilterInput.disabled = selectionDependentDisabled;
  }

  if (worldNameFilterSearchButton) {
    worldNameFilterSearchButton.textContent = '検索';
    worldNameFilterSearchButton.disabled = selectionDependentDisabled;
    worldNameFilterSearchButton.setAttribute('title', '検索を実行');
  }

  if (toolbarSearchClearButton) {
    toolbarSearchClearButton.textContent = 'クリア';
    toolbarSearchClearButton.disabled =
      selectionDependentDisabled || (!hasDraftSearch && !hasActiveSearch);
    toolbarSearchClearButton.setAttribute('title', '検索をクリア');
  }

  if (toolbarSearchScopeButton) {
    toolbarSearchScopeButton.disabled = selectionDependentDisabled;
    toolbarSearchScopeButton.setAttribute(
      'title',
      selectionDependentDisabled
        ? '写真を選択すると利用できます'
        : '検索対象を切り替え'
    );
  }
}

function renderToolbarSearchScopeMenu() {
  if (!toolbarSearchScopeButton || !toolbarSearchScopeMenu) {
    return;
  }

  const scopeMeta = getToolbarSearchScopeMeta();
  toolbarSearchScopeButton.innerHTML = `
    <span class="toolbar-search-scope-button-label">${escapeHtml(
      scopeMeta.buttonLabel
    )}</span>
    <span class="material-symbols-outlined orientation-filter-chevron">expand_more</span>
  `;

  toolbarSearchScopeMenu.innerHTML = '';

  Object.entries(TOOLBAR_SEARCH_SCOPE_META).forEach(([scope, meta]) => {
    const isActive = scope === draftToolbarSearchScope;
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'header-dropdown-item';
    item.dataset.toolbarSearchScope = scope;
    item.setAttribute('role', 'menuitemradio');
    item.setAttribute('aria-checked', isActive ? 'true' : 'false');
    item.classList.toggle('is-active', isActive);

    const label = document.createElement('span');
    label.className = 'header-dropdown-item-label';
    label.textContent = meta.label;
    item.appendChild(label);

    const check = document.createElement('span');
    check.className = 'material-symbols-outlined header-dropdown-check';
    check.textContent = 'check';
    item.appendChild(check);

    toolbarSearchScopeMenu.appendChild(item);
  });
}

function setToolbarSearchScopeMenuOpen(isOpen) {
  const nextOpen = Boolean(isOpen) && !isImporting && Boolean(currentSelection);
  if (nextOpen) {
    closeManagedDropdownsExcept(toolbarSearchScopeDropdown);
  }
  isToolbarSearchScopeMenuOpen = nextOpen;
  setAnimatedDropdownOpenState({
    dropdown: toolbarSearchScopeDropdown,
    button: toolbarSearchScopeButton,
    menu: toolbarSearchScopeMenu,
    isOpen: nextOpen,
    closeTimerRef: {
      get current() {
        return toolbarSearchScopeMenuCloseTimer;
      },
      set current(value) {
        toolbarSearchScopeMenuCloseTimer = value;
      },
    },
  });
}

function closeToolbarSearchScopeMenu() {
  setToolbarSearchScopeMenuOpen(false);
}

async function setToolbarSearchScope(nextScope) {
  const normalizedScope = TOOLBAR_SEARCH_SCOPE_META[nextScope]
    ? nextScope
    : 'world';

  if (draftToolbarSearchScope === normalizedScope) {
    closeToolbarSearchScopeMenu();
    return;
  }

  draftToolbarSearchScope = normalizedScope;
  renderToolbarSearchScopeMenu();
  syncToolbarSearchInputUi();
  closeToolbarSearchScopeMenu();
  worldNameFilterInput?.focus({ preventScroll: true });
}

async function submitWorldNameFilter({ focusCards = false } = {}) {
  const nextValue = worldNameFilterInput?.value || '';
  const normalizedValue = normalizeWorldNameFilterText(nextValue);
  const previousSearchScope = activeToolbarSearchScope;
  const previousFilterValue = activeWorldNameFilter;
  const didScopeChange = previousSearchScope !== draftToolbarSearchScope;

  clearWorldNameFilterInputTimer();
  activeToolbarSearchScope = draftToolbarSearchScope;

  if (
    currentSelection &&
    didScopeChange &&
    previousFilterValue === normalizedValue &&
    normalizedValue.length > 0
  ) {
    if (worldNameFilterInput) {
      worldNameFilterInput.value = normalizedValue;
    }
    await syncCurrentPhotoFilterPresentation({ animate: false });
  } else {
    await applyWorldNameFilter(nextValue, { animate: false });

    if (
      currentSelection &&
      didScopeChange &&
      previousFilterValue === normalizedValue &&
      normalizedValue.length === 0
    ) {
      syncFavoriteFilterUi();
    }
  }

  if (!focusCards) {
    return;
  }

  closeToolbarSearchScopeMenu();
  worldNameFilterInput?.blur();
  worldNameFilterSearchButton?.blur();

  requestAnimationFrame(() => {
    syncKeyboardFocusedPhotoCard({ force: true });
  });
}

async function clearWorldNameFilter({ keepFocus = true } = {}) {
  if (worldNameFilterInput) {
    worldNameFilterInput.value = '';
  }

  clearWorldNameFilterInputTimer();
  closeToolbarSearchScopeMenu();
  await applyWorldNameFilter('', { animate: false });
  syncToolbarSearchInputUi();

  if (keepFocus) {
    requestAnimationFrame(() => {
      worldNameFilterInput?.focus({ preventScroll: true });
    });
  }
}

function isToolbarSearchInteractionActive() {
  const activeElement = document.activeElement;

  return Boolean(
    activeElement &&
      (activeElement === worldNameFilterInput ||
        worldNameFilterMenu?.contains(activeElement) ||
        toolbarSearchScopeDropdown?.contains(activeElement))
  );
}

function setWorldNameFilterMenuOpen(isOpen) {
  if (isStaticToolbarWorldFilter()) {
    isWorldNameFilterMenuOpen = false;
    clearWorldNameFilterInputTimer();
    ensureStaticToolbarWorldFilterVisible();

    if (isOpen) {
      requestAnimationFrame(() => {
        worldNameFilterInput?.focus({ preventScroll: true });
        worldNameFilterInput?.select();
      });
    }

    return;
  }

  const nextOpen = Boolean(isOpen) && !isImporting && Boolean(currentSelection);
  if (nextOpen) {
    closeManagedDropdownsExcept(worldNameFilterDropdown);
  }
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
      isOpen: () => isToolbarSearchScopeMenuOpen,
      dropdown: toolbarSearchScopeDropdown,
      close: closeToolbarSearchScopeMenu,
    },
    {
      isOpen: () => isRegenerateThumbnailMonthMenuOpen,
      dropdown: regenerateThumbnailMonthDropdown,
      close: closeRegenerateThumbnailMonthMenu,
    },
    {
      isOpen: () => isReimportRegisteredPhotoMonthMenuOpen,
      dropdown: reimportRegisteredPhotoMonthDropdown,
      close: closeReimportRegisteredPhotoMonthMenu,
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

function closeManagedDropdownsExcept(activeDropdown) {
  for (const entry of getManagedDropdownClosers()) {
    if (!entry.isOpen()) {
      continue;
    }

    if (entry.dropdown === activeDropdown) {
      continue;
    }

    entry.close();
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

  if (Number.isFinite(width) && Number.isFinite(height)) {
    if (width === height) {
      return 'square';
    }

    return width > height ? 'landscape' : 'portrait';
  }

  if (photo.orientationTier) {
    return photo.orientationTier;
  }

  return null;
}

function getToolbarSearchTargetText(photo) {
  if (!photo) {
    return '';
  }

  if (activeToolbarSearchScope === 'memo') {
    return [photo.memoText]
      .filter(Boolean)
      .join(' ')
      .normalize('NFKC')
      .toLocaleLowerCase('ja-JP');
  }

  if (activeToolbarSearchScope === 'printNote') {
    return [photo.printNoteText]
      .filter(Boolean)
      .join(' ')
      .normalize('NFKC')
      .toLocaleLowerCase('ja-JP');
  }

  return [photo.worldName, photo.rawWorldName, photo.worldId]
    .filter(Boolean)
    .join(' ')
    .normalize('NFKC')
    .toLocaleLowerCase('ja-JP');
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
    const targetText = getToolbarSearchTargetText(photo);

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
    return getDefaultSelectionEmptyMessage();
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
  scheduleMainHeaderResponsiveLayout();
}

function setAnimatedMonthLabelText(nextText, { animate = true } = {}) {
  setAnimatedHeaderText(currentMonthLabel, nextText, {
    animate,
    durationMs: 2320,
  });
  scheduleMainHeaderResponsiveLayout();
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

  if (photoSortButton) {
    const isOldestFirst = currentPhotoSortOrder === 'asc';
    const label = isOldestFirst ? '並び順: 古い順' : '並び順: 新しい順';

    photoSortButton.classList.toggle('is-active', isOldestFirst);
    photoSortButton.disabled = isImporting || !currentSelection;
    photoSortButton.setAttribute('aria-label', label);
    photoSortButton.setAttribute('title', label);
  }

  if (photoSortIcon) {
    photoSortIcon.textContent =
      currentPhotoSortOrder === 'asc' ? 'arrow_upward_alt' : 'arrow_downward_alt';
  }

  syncPhotoCardDensityUi();

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
    orientationFilterLabel.textContent = '向き';
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
    photoLabelFilterLabel.textContent = 'ラベル';
  }

  if (worldNameFilterButton) {
    const label = `検索: ${getWorldNameFilterSummaryText({
      includePrefix: true,
    })}`;

    worldNameFilterButton.classList.toggle(
      'is-active',
      activeWorldNameFilter.length > 0
    );
    worldNameFilterButton.disabled = isImporting || !currentSelection;
    worldNameFilterButton.setAttribute('aria-label', label);
    worldNameFilterButton.setAttribute('title', label);
  }

  if (worldNameFilterLabel) {
    worldNameFilterLabel.textContent = getToolbarSearchScopeMeta().buttonLabel;
  }

  for (const item of orientationFilterItems) {
    const isActive = item.dataset.orientationFilter === activeOrientationFilter;
    item.classList.toggle('is-active', isActive);
    item.setAttribute('aria-checked', isActive ? 'true' : 'false');
  }

  renderPhotoLabelFilterMenu();
  renderToolbarSearchScopeMenu();
  syncToolbarSearchInputUi();

  if (isImporting || !currentSelection) {
    closeOrientationFilterMenu();
    closePhotoLabelFilterMenu();
    closeToolbarSearchScopeMenu();
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

// Any filter change should recalculate the current month, refresh button/count
// text, and optionally replay the gallery transition in one shared step.
async function syncCurrentPhotoFilterPresentation({ animate = true } = {}) {
  applyCurrentPhotoFilter();
  syncFavoriteFilterUi();

  if (animate) {
    await refreshCurrentMonthWithFilterAnimation();
    return;
  }

  renderMonthGallery({ resetProgressive: true });
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

  await syncCurrentPhotoFilterPresentation();
}

async function setPhotoLabelFilterMode(nextMode) {
  const normalizedMode = nextMode === 'and' ? 'and' : 'or';

  if (!currentSelection || photoLabelFilterMode === normalizedMode) {
    return;
  }

  photoLabelFilterMode = normalizedMode;

  if (activePhotoLabelFilters.length === 0) {
    await syncCurrentPhotoFilterPresentation({ animate: false });
    return;
  }

  await syncCurrentPhotoFilterPresentation();
}

async function applyWorldNameFilter(nextValue, { animate = true } = {}) {
  const normalizedValue = normalizeWorldNameFilterText(nextValue);

  if (!currentSelection || activeWorldNameFilter === normalizedValue) {
    if (worldNameFilterInput && worldNameFilterInput.value !== normalizedValue) {
      worldNameFilterInput.value = normalizedValue;
    }
    return;
  }

  activeWorldNameFilter = normalizedValue;
  if (worldNameFilterInput) {
    worldNameFilterInput.value = normalizedValue;
  }
  await syncCurrentPhotoFilterPresentation({ animate });
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
  closeOrientationFilterMenu();
  await syncCurrentPhotoFilterPresentation();
}

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

function syncSelectionModeButtonState() {
  if (selectionModeButton) {
    selectionModeButton.classList.toggle('is-active', isSelectionMode);
    selectionModeButton.textContent = isSelectionMode
      ? `${selectedPhotoIds.size}件選択中`
      : '選択';
    selectionModeButton.disabled = isImporting || !currentSelection;
  }
}

function syncBulkFavoriteButtonState() {
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
}

function syncBulkDeleteButtonState() {
  if (bulkDeleteButton) {
    bulkDeleteButton.disabled =
      isImporting || !isSelectionMode || selectedPhotoIds.size === 0;

    bulkDeleteButton.textContent = '削除';
  }
}

function syncSelectionUi() {
  syncSelectionModeButtonState();
  syncBulkFavoriteButtonState();
  syncBulkDeleteButtonState();
}

function clearSelectionState() {
  isSelectionMode = false;
  selectedPhotoIds.clear();
  lastSelectionAnchorPhotoId = null;
  isSelectionDragActive = false;
  selectionDragTargetState = null;
  selectionDragPointerId = null;
  suppressSelectionModeCardClickPhotoId = null;
  lastSelectionDragPhotoId = null;
  syncRenderedSelectionState();
  syncSelectionUi();
  syncKeyboardFocusedPhotoCard();
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

function isEditableKeyboardTarget(target) {
  return Boolean(
    target?.closest?.(
      'input, textarea, select, [contenteditable="true"], [contenteditable=""], [role="textbox"]'
    )
  );
}

function getRenderedVisiblePhotoCards() {
  if (!monthGalleryList) {
    return [];
  }

  return Array.from(monthGalleryList.querySelectorAll('.photo-card')).filter(
    (card) => !card.hidden
  );
}

function clearKeyboardFocusedPhotoCard() {
  if (!monthGalleryList) {
    keyboardFocusedPhotoId = null;
    return;
  }

  monthGalleryList
    .querySelectorAll('.photo-card.is-keyboard-focused')
    .forEach((card) => {
      card.classList.remove('is-keyboard-focused');
      card.removeAttribute('tabindex');
    });
}

function setKeyboardFocusedPhoto(photoId, { scroll = true } = {}) {
  if (!Number.isInteger(photoId) || photoId <= 0 || !monthGalleryList) {
    return false;
  }

  const card = monthGalleryList.querySelector(
    `.photo-card[data-photo-id="${photoId}"]`
  );

  if (!card || card.hidden) {
    return false;
  }

  clearKeyboardFocusedPhotoCard();
  keyboardFocusedPhotoId = photoId;
  card.classList.add('is-keyboard-focused');
  card.tabIndex = -1;
  card.focus({ preventScroll: true });

  if (scroll) {
    card.scrollIntoView({
      block: 'nearest',
      inline: 'nearest',
      behavior: 'smooth',
    });
  }

  return true;
}

function syncKeyboardFocusedPhotoCard({ force = false } = {}) {
  if (!force && isToolbarSearchInteractionActive()) {
    return;
  }

  const visibleCards = getRenderedVisiblePhotoCards();

  if (visibleCards.length === 0) {
    clearKeyboardFocusedPhotoCard();
    keyboardFocusedPhotoId = null;
    return;
  }

  if (
    Number.isInteger(keyboardFocusedPhotoId) &&
    setKeyboardFocusedPhoto(keyboardFocusedPhotoId, { scroll: false })
  ) {
    return;
  }

  keyboardFocusedPhotoId = Number(visibleCards[0].dataset.photoId) || null;

  if (keyboardFocusedPhotoId) {
    setKeyboardFocusedPhoto(keyboardFocusedPhotoId, { scroll: false });
  }
}

function moveKeyboardFocusedPhoto(delta) {
  const visibleCards = getRenderedVisiblePhotoCards();

  if (visibleCards.length === 0) {
    return;
  }

  const visiblePhotoIds = visibleCards.map((card) => Number(card.dataset.photoId));
  const currentIndex = visiblePhotoIds.indexOf(keyboardFocusedPhotoId);
  const baseIndex = currentIndex === -1 ? 0 : currentIndex;
  const nextIndex = Math.max(
    0,
    Math.min(visiblePhotoIds.length - 1, baseIndex + delta)
  );

  setKeyboardFocusedPhoto(visiblePhotoIds[nextIndex]);
}

function getVisiblePhotoCardRows() {
  const rowTolerance = 24;
  const cards = getRenderedVisiblePhotoCards()
    .map((card) => ({
      card,
      photoId: Number(card.dataset.photoId),
      rect: card.getBoundingClientRect(),
    }))
    .filter((entry) => Number.isInteger(entry.photoId) && entry.photoId > 0)
    .sort((left, right) => {
      const topDelta = left.rect.top - right.rect.top;
      if (Math.abs(topDelta) > rowTolerance) {
        return topDelta;
      }

      return left.rect.left - right.rect.left;
    });

  const rows = [];

  for (const entry of cards) {
    const previousRow = rows.at(-1);

    if (!previousRow || Math.abs(previousRow.top - entry.rect.top) > rowTolerance) {
      rows.push({
        top: entry.rect.top,
        items: [entry],
      });
      continue;
    }

    previousRow.items.push(entry);
  }

  for (const row of rows) {
    row.items.sort((left, right) => left.rect.left - right.rect.left);
  }

  return rows;
}

function moveKeyboardFocusedPhotoVertical(rowDelta) {
  const rows = getVisiblePhotoCardRows();

  if (rows.length === 0) {
    return;
  }

  let currentRowIndex = -1;
  let currentColumnIndex = 0;

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const columnIndex = rows[rowIndex].items.findIndex(
      (entry) => entry.photoId === keyboardFocusedPhotoId
    );

    if (columnIndex !== -1) {
      currentRowIndex = rowIndex;
      currentColumnIndex = columnIndex;
      break;
    }
  }

  if (currentRowIndex === -1) {
    const fallbackPhotoId = rows[0].items[0]?.photoId;
    if (fallbackPhotoId) {
      setKeyboardFocusedPhoto(fallbackPhotoId);
    }
    return;
  }

  const nextRowIndex = Math.max(
    0,
    Math.min(rows.length - 1, currentRowIndex + rowDelta)
  );
  const nextRowItems = rows[nextRowIndex].items;
  const currentEntry = rows[currentRowIndex].items[currentColumnIndex];
  const currentCenterX =
    currentEntry.rect.left + currentEntry.rect.width / 2;
  const overlappingItems = nextRowItems.filter(
    (entry) =>
      currentCenterX >= entry.rect.left && currentCenterX <= entry.rect.right
  );

  let nextEntry = null;

  if (overlappingItems.length > 0) {
    nextEntry = overlappingItems.reduce((closest, entry) => {
      if (!closest) {
        return entry;
      }

      const closestCenterX =
        closest.rect.left + closest.rect.width / 2;
      const entryCenterX = entry.rect.left + entry.rect.width / 2;

      return Math.abs(entryCenterX - currentCenterX) <
        Math.abs(closestCenterX - currentCenterX)
        ? entry
        : closest;
    }, null);
  } else {
    nextEntry =
      rowDelta > 0
        ? nextRowItems[nextRowItems.length - 1]
        : nextRowItems[0];
  }

  const nextPhotoId = nextEntry?.photoId;

  if (nextPhotoId) {
    setKeyboardFocusedPhoto(nextPhotoId);
  }
}

function activateKeyboardFocusedPhoto() {
  if (!Number.isInteger(keyboardFocusedPhotoId) || keyboardFocusedPhotoId <= 0) {
    return;
  }

  if (isSelectionMode) {
    togglePhotoSelection(keyboardFocusedPhotoId);
    return;
  }

  const photo = getLatestKnownPhotoById(keyboardFocusedPhotoId);

  if (photo) {
    openImageModal(photo);
  }
}

function selectPhotoRange(anchorPhotoId, targetPhotoId) {
  const visiblePhotoIds = currentPhotos.map((photo) => photo.id);
  const startIndex = visiblePhotoIds.indexOf(anchorPhotoId);
  const endIndex = visiblePhotoIds.indexOf(targetPhotoId);

  if (startIndex === -1 || endIndex === -1) {
    return false;
  }

  const rangeStart = Math.min(startIndex, endIndex);
  const rangeEnd = Math.max(startIndex, endIndex);

  for (let index = rangeStart; index <= rangeEnd; index += 1) {
    selectedPhotoIds.add(visiblePhotoIds[index]);
  }

  return true;
}

function setPhotoSelectionState(photoId, shouldSelect) {
  const isSelected = selectedPhotoIds.has(photoId);

  if (shouldSelect === isSelected) {
    return false;
  }

  if (shouldSelect) {
    selectedPhotoIds.add(photoId);
  } else {
    selectedPhotoIds.delete(photoId);
  }

  lastSelectionAnchorPhotoId = photoId;
  syncRenderedPhotoSelectionState(photoId);
  return true;
}

function togglePhotoSelection(photoId, { rangeSelect = false } = {}) {
  if (
    rangeSelect &&
    Number.isInteger(lastSelectionAnchorPhotoId) &&
    lastSelectionAnchorPhotoId > 0 &&
    lastSelectionAnchorPhotoId !== photoId &&
    selectPhotoRange(lastSelectionAnchorPhotoId, photoId)
  ) {
    syncRenderedSelectionState();
    syncSelectionUi();
    return;
  }

  if (selectedPhotoIds.has(photoId)) {
    selectedPhotoIds.delete(photoId);
  } else {
    selectedPhotoIds.add(photoId);
  }

  lastSelectionAnchorPhotoId = photoId;
  syncRenderedPhotoSelectionState(photoId);
  syncSelectionUi();
}

function finishSelectionDrag() {
  if (!isSelectionDragActive) {
    return;
  }

  isSelectionDragActive = false;
  selectionDragTargetState = null;
  selectionDragPointerId = null;
  lastSelectionDragPhotoId = null;
  window.setTimeout(() => {
    suppressSelectionModeCardClickPhotoId = null;
  }, 0);
}

function applySelectionDragToPhoto(photoId) {
  if (
    !isSelectionDragActive ||
    !Number.isInteger(photoId) ||
    photoId <= 0 ||
    lastSelectionDragPhotoId === photoId
  ) {
    return;
  }

  lastSelectionDragPhotoId = photoId;
  suppressSelectionModeCardClickPhotoId = photoId;
  if (setPhotoSelectionState(photoId, selectionDragTargetState)) {
    syncSelectionUi();
  }
}

function beginSelectionDrag(photoId, pointerId) {
  if (!isSelectionMode || !Number.isInteger(photoId) || photoId <= 0) {
    return;
  }

  const targetState = !selectedPhotoIds.has(photoId);
  isSelectionDragActive = true;
  selectionDragTargetState = targetState;
  selectionDragPointerId = pointerId ?? null;
  suppressSelectionModeCardClickPhotoId = photoId;
  lastSelectionDragPhotoId = null;
  applySelectionDragToPhoto(photoId);
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

// Small UI-only helpers keep modal/card refresh logic readable: collection
// updates happen first, then the current month gallery decides whether a local
// card swap is enough or a full rerender is safer.
function isPhotoVisibleAfterCurrentFilters(photoId) {
  return currentPhotos.some((photo) => photo.id === photoId);
}

function updatePhotoInCurrentCollections(updatedPhoto) {
  setCurrentMonthPhotos(
    allCurrentMonthPhotos.map((photo) =>
      photo.id === updatedPhoto.id ? updatedPhoto : photo
    )
  );
}

function syncMonthGalleryAfterPhotoUpdate(updatedPhoto) {
  const isVisibleAfterFilters = isPhotoVisibleAfterCurrentFilters(updatedPhoto.id);
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

  return {
    rerenderedMonthGallery,
    isVisibleAfterFilters,
  };
}

function syncMonthGalleryAfterPhotoBatchUpdate(normalizedUpdates) {
  let rerenderedMonthGallery = false;

  for (const photo of normalizedUpdates) {
    const isVisibleAfterFilters = isPhotoVisibleAfterCurrentFilters(photo.id);

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

  return rerenderedMonthGallery;
}

// Single-photo updates from modal actions should keep collections, cards, and
// the open modal in sync without each caller repeating the same flow.
function syncSinglePhotoUpdate(updatedPhoto, { refreshModal = true } = {}) {
  if (!updatedPhoto) {
    return null;
  }

  updatePhotoInCurrentCollections(updatedPhoto);
  syncMonthGalleryAfterPhotoUpdate(updatedPhoto);

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

  setCurrentMonthPhotos(
    allCurrentMonthPhotos.map(
      (photo) => updatedPhotoMap.get(photo.id) || photo
    )
  );
  syncMonthGalleryAfterPhotoBatchUpdate(normalizedUpdates);

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

  setCurrentMonthPhotos(
    allCurrentMonthPhotos.map((photo) => ({
      ...photo,
      thumbnailPath: null,
      thumbnailUrl: null,
    }))
  );
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
  setCurrentMonthPhotos(
    allCurrentMonthPhotos.filter((photo) => photo.id !== photoId)
  );
}

function removePhotosFromCurrentCollections(photoIds) {
  const targetIdSet = new Set(photoIds);

  setCurrentMonthPhotos(
    allCurrentMonthPhotos.filter((photo) => !targetIdSet.has(photo.id))
  );
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

  syncKeyboardFocusedPhotoCard();
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

  syncKeyboardFocusedPhotoCard();

  return true;
}

function syncSidebarSelectionState() {
  if (!sidebarTree || sidebarTree.children.length === 0) {
    return false;
  }

  if (currentSidebarMode === 'world') {
    let hasWorldEntries = false;

    for (const worldButton of sidebarTree.querySelectorAll('.world-sidebar-item')) {
      const isActive =
        isWorldSelection(currentSelection) &&
        currentSelection.worldKey === worldButton.dataset.worldKey;

      hasWorldEntries = true;
      worldButton.classList.toggle('active', Boolean(isActive));
    }

    return hasWorldEntries;
  }

  let hasSidebarEntries = false;

  for (const yearBlock of sidebarTree.querySelectorAll('.year-block')) {
    const year = Number(yearBlock.dataset.year);
    const yearButton = yearBlock.querySelector('.year-button');
    const toggle = yearBlock.querySelector('.year-toggle');
    const monthList = yearBlock.querySelector('.month-list');
    const isExpanded = expandedYears.has(year);
    const isActiveYear =
      Number.isFinite(year) &&
      isYearSelection(currentSelection) &&
      currentSelection.year === year;

    hasSidebarEntries = true;
    monthList?.classList.toggle('hidden', !isExpanded);
    yearButton?.classList.toggle('active', Boolean(isActiveYear));

    if (toggle) {
      toggle.textContent = isExpanded ? '▾' : '▸';
    }

    for (const monthButton of yearBlock.querySelectorAll('.month-button')) {
      const month = Number(monthButton.dataset.month);
      const isActive =
        Number.isFinite(year) &&
        Number.isFinite(month) &&
        isMonthSelection(currentSelection) &&
        currentSelection.year === year &&
        currentSelection.month === month;

      monthButton.classList.toggle('active', Boolean(isActive));
    }
  }

  return hasSidebarEntries;
}

function setCurrentSelectionValue(selection) {
  currentSelection = normalizeSelection(selection);

  if (!currentSelection) {
    return;
  }

  if (currentSelection.mode === 'world') {
    lastWorldSelection = currentSelection;
    return;
  }

  lastTimelineSelection = currentSelection;

  if (currentSelection.year) {
    expandedYears.add(currentSelection.year);
  }
}

function ensureExpandedYearsForSidebar() {
  if (currentSidebarMode === 'world') {
    return;
  }

  if (expandedYears.size === 0) {
    for (const yearEntry of sidebarData) {
      expandedYears.add(yearEntry.year);
    }
  }

  if (currentSelection?.year) {
    expandedYears.add(currentSelection.year);
  }
}

// Month selection affects multiple surfaces at once, so keep sidebar/settings
// synchronization together instead of scattering it across each caller.
function syncSelectionLinkedUi({ forceSidebarRender = false } = {}) {
  ensureExpandedYearsForSidebar();

  if (forceSidebarRender || !syncSidebarSelectionState()) {
    renderSidebar();
  }

  syncSelectionDependentSettingsUi();
}

function ensureSidebarWorldSortControls() {
  if (!sidebarHeader || sidebarHeaderControls) {
    return;
  }

  sidebarHeaderControls = document.createElement('div');
  sidebarHeaderControls.className = 'sidebar-header-controls';

  const sortToggleGroup = document.createElement('div');
  sortToggleGroup.className = 'sidebar-sort-toggle-group';

  sidebarSortCountButton = document.createElement('button');
  sidebarSortCountButton.type = 'button';
  sidebarSortCountButton.className = 'sidebar-sort-toggle';
  sidebarSortCountButton.textContent = '撮影枚数順';

  sidebarSortNameButton = document.createElement('button');
  sidebarSortNameButton.type = 'button';
  sidebarSortNameButton.className = 'sidebar-sort-toggle';
  sidebarSortNameButton.textContent = '名前順';

  sortToggleGroup.append(sidebarSortCountButton, sidebarSortNameButton);
  sidebarHeaderControls.appendChild(sortToggleGroup);
  sidebarHeader.appendChild(sidebarHeaderControls);
}

function syncWorldSidebarSortButtons() {
  sidebarSortCountButton?.classList.toggle(
    'is-active',
    currentWorldSidebarSort === 'count'
  );
  sidebarSortNameButton?.classList.toggle(
    'is-active',
    currentWorldSidebarSort === 'name'
  );
}

function setSidebarHeaderControlsVisible(isVisible) {
  if (!sidebarHeaderControls) {
    return;
  }

  if (sidebarHeaderControlsHideTimer) {
    clearTimeout(sidebarHeaderControlsHideTimer);
    sidebarHeaderControlsHideTimer = null;
  }

  if (sidebar?.classList.contains('is-mode-switching')) {
    sidebarHeaderControls.hidden = !isVisible;
    sidebarHeaderControls.classList.toggle('is-visible', isVisible);
    return;
  }

  if (isVisible) {
    sidebarHeaderControls.hidden = false;
    requestAnimationFrame(() => {
      sidebarHeaderControls?.classList.add('is-visible');
    });
    return;
  }

  sidebarHeaderControls.classList.remove('is-visible');
  sidebarHeaderControlsHideTimer = setTimeout(() => {
    if (sidebarHeaderControls) {
      sidebarHeaderControls.hidden = true;
    }
    sidebarHeaderControlsHideTimer = null;
  }, 920);
}

async function runSidebarModeSwitchTransition(action) {
  if (!sidebar) {
    await action();
    return;
  }

  sidebar.classList.add('is-mode-switching');
  await new Promise((resolve) => setTimeout(resolve, 260));
  await action();
  await new Promise((resolve) => setTimeout(resolve, 60));
  await new Promise((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(resolve);
    });
  });
  sidebar.classList.remove('is-mode-switching');
}

async function runSidebarTreeRefreshTransition(action) {
  if (!sidebar) {
    await action();
    return;
  }

  sidebar.classList.add('is-tree-switching');
  // Wait until the old sidebar list has faded out enough before swapping
  // the sorted content, otherwise both states appear to overlap.
  await new Promise((resolve) => setTimeout(resolve, 400));
  await action();
  await new Promise((resolve) => setTimeout(resolve, 70));
  await new Promise((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(resolve);
    });
  });
  sidebar.classList.remove('is-tree-switching');
}

function syncSidebarModeUi() {
  if (sidebarHeaderTitle) {
    sidebarHeaderTitle.textContent =
      currentSidebarMode === 'world' ? 'ワールド' : '年月';
  }

  if (sidebarHeaderDescription) {
    sidebarHeaderDescription.hidden = currentSidebarMode === 'world';
  }

  setSidebarHeaderControlsVisible(currentSidebarMode === 'world');

  if (worldLibraryModeButton) {
    worldLibraryModeButton.classList.toggle(
      'is-active',
      currentSidebarMode === 'world'
    );
    worldLibraryModeButton.innerHTML =
      currentSidebarMode === 'world'
        ? '<span class="material-symbols-outlined">calendar_month</span><span>年月一覧</span>'
        : '<span class="material-symbols-outlined">public</span><span>ワールド一覧</span>';
    worldLibraryModeButton.setAttribute(
      'aria-label',
      currentSidebarMode === 'world' ? '年月一覧へ戻る' : 'ワールド一覧を表示'
    );
    worldLibraryModeButton.setAttribute(
      'title',
      currentSidebarMode === 'world' ? '年月一覧へ戻る' : 'ワールド一覧を表示'
    );
  }

  syncWorldSidebarSortButtons();
}

function applySidebarDeletionLocally(targetSelection, removedCount) {
  const normalizedSelection = normalizeSelection(targetSelection);

  if (
    !normalizedSelection ||
    normalizedSelection.mode !== 'month' ||
    !Number.isFinite(removedCount) ||
    removedCount <= 0 ||
    sidebarData.length === 0
  ) {
    return false;
  }

  let changed = false;

  sidebarData = sidebarData
    .map((yearEntry) => {
      if (yearEntry.year !== normalizedSelection.year) {
        return yearEntry;
      }

      const nextMonths = yearEntry.months
        .map((monthEntry) => {
          if (monthEntry.month !== normalizedSelection.month) {
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

  return createMonthSelection(latestYearEntry.year, latestMonthEntry.month);
}

function getLatestWorldSelectionFromSidebarData() {
  const latestWorldEntry = worldSidebarData[0];

  if (!latestWorldEntry) {
    return null;
  }

  return createWorldSelection(
    latestWorldEntry.worldKey,
    latestWorldEntry.worldName,
    latestWorldEntry.worldId
  );
}

// Sidebar restoration now supports both "year" and "month" selections.
function hasSidebarSelection(selection) {
  const normalizedSelection = normalizeSelection(selection);

  if (!normalizedSelection) {
    return false;
  }

  if (normalizedSelection.mode === 'world') {
    return worldSidebarData.some(
      (worldEntry) => worldEntry.worldKey === normalizedSelection.worldKey
    );
  }

  const matchingYearEntry = sidebarData.find(
    (yearEntry) => yearEntry.year === normalizedSelection.year
  );

  if (!matchingYearEntry) {
    return false;
  }

  if (normalizedSelection.mode === 'year') {
    return true;
  }

  return matchingYearEntry.months.some(
    (monthEntry) => monthEntry.month === normalizedSelection.month
  );
}

async function selectSidebarSelectionIfAvailable(selection) {
  const normalizedSelection = normalizeSelection(selection);

  if (!hasSidebarSelection(normalizedSelection)) {
    return false;
  }

  if (normalizedSelection.mode === 'world') {
    const matchingWorldEntry = worldSidebarData.find(
      (worldEntry) => worldEntry.worldKey === normalizedSelection.worldKey
    );

    if (!matchingWorldEntry) {
      return false;
    }

    await selectWorld(matchingWorldEntry);
  } else if (normalizedSelection.mode === 'year') {
    await selectYear(normalizedSelection.year);
  } else {
    await selectMonth(normalizedSelection.year, normalizedSelection.month);
  }

  return true;
}

// Data-changing operations often need to restore the same selection, fall back
// to a caller-provided selection, or finally show the latest available month.
async function restoreMonthViewAfterDataChange({
  preferredSelection = null,
  fallbackSelection = null,
  clearWhenEmpty = true,
} = {}) {
  if (currentSidebarMode === 'world') {
    if (await selectSidebarSelectionIfAvailable(preferredSelection)) {
      return true;
    }

    if (await selectSidebarSelectionIfAvailable(fallbackSelection)) {
      return true;
    }

    const latestWorldSelection = getLatestWorldSelectionFromSidebarData();

    if (latestWorldSelection) {
      await selectSidebarSelectionIfAvailable(latestWorldSelection);
      return true;
    }

    if (clearWhenEmpty) {
      clearMainContent();
    }

    return false;
  }

  if (await selectSidebarSelectionIfAvailable(preferredSelection)) {
    return true;
  }

  if (await selectSidebarSelectionIfAvailable(fallbackSelection)) {
    return true;
  }

  const latestMonth = getLatestSelectionFromSidebarData();

  if (latestMonth) {
    await selectMonth(latestMonth.year, latestMonth.month);
    return true;
  }

  if (clearWhenEmpty) {
    clearMainContent();
  }

  return false;
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

  renderModalPrintNote(item);

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

  setModalPhotoMemoStatus('');
  setModalPhotoMemoSaveButtonBusy(false);

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

function getNormalizedModalPrintNoteText(item = currentModalPhoto) {
  if (typeof item?.printNoteText !== 'string') {
    return '';
  }

  const normalized = item.printNoteText.trim();
  return /^\[object\s+[^\]]+\]$/i.test(normalized) ? '' : normalized;
}

function renderModalPrintNote(item = currentModalPhoto) {
  const printNoteText = getNormalizedModalPrintNoteText(item);
  const hasPrintNote = printNoteText.length > 0;

  if (modalPrintNoteValue) {
    modalPrintNoteValue.textContent = hasPrintNote ? printNoteText : '';
  }

  if (modalPrintNoteBlock) {
    modalPrintNoteBlock.classList.toggle('is-hidden', !hasPrintNote);
  }

  if (modalPrintNoteHeroBadge) {
    modalPrintNoteHeroBadge.textContent = 'プリント';
    modalPrintNoteHeroBadge.classList.toggle('is-hidden', !hasPrintNote);
  }
}

function setPhotoLabelCatalogMenuOpen(isOpen) {
  const nextOpen = Boolean(isOpen);
  if (nextOpen) {
    closeManagedDropdownsExcept(photoLabelCatalogDropdown);
  }
  isPhotoLabelCatalogMenuOpen = nextOpen;
  setAnimatedDropdownOpenState({
    dropdown: photoLabelCatalogDropdown,
    button: photoLabelCatalogButton,
    menu: photoLabelCatalogMenu,
    isOpen: nextOpen,
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
  const updatedModalPhoto = {
    ...currentModalPhoto,
    photoLabels: currentModalPhotoLabels,
  };
  syncSinglePhotoUpdate(updatedModalPhoto, { refreshModal: false });
  setCurrentModalPhotoState(updatedModalPhoto);
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

function invalidateCurrentModalAsyncRequests() {
  modalWorldMetadataRequestId += 1;
  modalPhotoLabelsRequestId += 1;
  modalImageRecoveryRequestId += 1;
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

function setCurrentModalPhotoState(photo, { resetScroll = false } = {}) {
  currentModalPhoto = photo;

  if (resetScroll && imageModalInfo) {
    imageModalInfo.scrollTop = 0;
  }

  syncImageModalPhotoLayout(photo);
  updateImageModalNavigationState();
  syncWorldMetadataSyncUi();
}

function clearCurrentModalPhotoState() {
  modalImage.src = '';
  currentModalPhoto = null;
  currentModalPhotoLabels = [];
  setWorldNameEditStatus('');
  setModalPhotoMemoStatus('');
  setModalPhotoMemoSaveButtonBusy(false);
  renderModalPrintNote(null);

  if (imageModalInfo) {
    imageModalInfo.scrollTop = 0;
  }

  updateImageModalNavigationState();
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

  invalidateCurrentModalAsyncRequests();
  setCurrentModalPhotoState(latestPhoto, { resetScroll: true });
  populateModal(latestPhoto);
  void loadModalWorldMetadata(latestPhoto);
  void loadModalPhotoLabels(latestPhoto);

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
  // Keep the hook in place for future experiments, but do not animate the
  // sticky shell today. The current effect reads as a reload/flash.
  if (modalShellRestoreTimer) {
    clearTimeout(modalShellRestoreTimer);
    modalShellRestoreTimer = null;
  }
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
  invalidateCurrentModalAsyncRequests();
  setImageModalScrollLock(false);

  const finalizeImageModalClose = () => {
    if (imageModalSwitchTimer) {
      clearTimeout(imageModalSwitchTimer);
      imageModalSwitchTimer = null;
    }

    imageModalBody?.classList.remove('is-switching-prev', 'is-switching-next');
    imageModal.classList.remove('is-closing');
    imageModal.classList.add('hidden');
    clearCurrentModalPhotoState();
    imageModalAnimationTimer = null;
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

  setWorldNameEditStatus('');
  syncWorldMetadataSyncUi();
  openSubModalElement(worldNameEditModal);
}

function closeWorldNameEditModal() {
  closeSubModalElement(worldNameEditModal, {
    onClosed: () => {
      setWorldNameEditStatus('');
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

function buildAppUpdatePromptConfig(payload) {
  const kind =
    typeof payload?.kind === 'string' ? payload.kind.trim().toLowerCase() : '';
  const version =
    typeof payload?.version === 'string' ? payload.version.trim() : '';
  const versionLabel = version || '最新バージョン';

  if (kind === 'downloaded') {
    return {
      title: 'アップデートの準備ができました',
      message: `${versionLabel} のダウンロードが完了しました。再起動して更新しますか？`,
      confirmText: '再起動して更新',
    };
  }

  if (kind === 'available') {
    return {
      title: 'アップデートがあります',
      message: `新しいバージョン ${versionLabel} が利用できます。今すぐダウンロードしますか？`,
      confirmText: '今すぐ更新',
    };
  }

  return null;
}

async function handleAppUpdateAction(payload) {
  const config = buildAppUpdatePromptConfig(payload);

  if (!config) {
    return;
  }

  const confirmed = await openConfirmModal(config);

  if (!confirmed) {
    if (payload?.kind === 'available') {
      showToast('アップデートは保留しました');
    } else if (payload?.kind === 'downloaded') {
      showToast('アップデートは準備済みです。あとで再起動して適用できます');
    }
    return;
  }

  try {
    if (payload?.kind === 'available') {
      const result = await window.electronAPI.startAppUpdateDownload?.();

      if (!result?.ok) {
        showToast(result?.message || 'アップデートを開始できませんでした');
      }
      return;
    }

    if (payload?.kind === 'downloaded') {
      const result = await window.electronAPI.installDownloadedAppUpdate?.();

      if (result && result.ok === false) {
        showToast(result.message || 'アップデートを適用できませんでした');
      }
    }
  } catch (error) {
    showToast(
      `アップデート処理に失敗しました: ${
        error instanceof Error ? error.message : '不明なエラー'
      }`
    );
  }
}

function queueAppUpdatePrompt(payload) {
  appUpdatePromptQueue = appUpdatePromptQueue
    .then(() => handleAppUpdateAction(payload))
    .catch((error) => {
      showToast(
        `アップデート確認の表示に失敗しました: ${
          error instanceof Error ? error.message : '不明なエラー'
        }`
      );
    });
}

function renderTrackedFolderList() {
  if (!trackedFolderList) {
    return;
  }

  syncTrackedFolderSettingsMeta();

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

  syncTrackedFolderListActionButtonsUi();
}

function syncTrackedFolderSettingsMeta() {
  if (!trackedFolderSettingsMeta) {
    return;
  }

  const count = Array.isArray(trackedFolders) ? trackedFolders.length : 0;
  trackedFolderSettingsMeta.textContent =
    count > 0
      ? `登録済み ${count}件`
      : 'まだ登録されていません';
}

function syncTrackedFolderSettingsActionsUi() {
  const hasTrackedFolders =
    Array.isArray(trackedFolders) && trackedFolders.length > 0;

  if (openTrackedFolderListButton) {
    openTrackedFolderListButton.disabled = isImporting || !hasTrackedFolders;
    openTrackedFolderListButton.setAttribute(
      'title',
      hasTrackedFolders ? '更新対象フォルダ一覧を表示' : '登録されたフォルダがありません'
    );
  }

  if (addTrackedFolderButton) {
    addTrackedFolderButton.disabled = isImporting;
    addTrackedFolderButton.setAttribute(
      'title',
      isImporting ? '処理中はフォルダを追加できません' : '更新対象フォルダを追加'
    );
  }
}

function syncTrackedFolderListActionButtonsUi() {
  trackedFolderList
    ?.querySelectorAll('[data-tracked-folder-path]')
    .forEach((button) => {
      button.disabled = isImporting;
      button.setAttribute(
        'title',
        isImporting ? '処理中はフォルダを削除できません' : 'このフォルダを削除'
      );
    });
}

function ensureSettingsMaintenanceStatus() {
  if (!settingsMaintenanceSection) {
    return null;
  }

  if (!settingsMaintenanceStatus) {
    settingsMaintenanceStatus = document.createElement('p');
    settingsMaintenanceStatus.className = 'settings-maintenance-status';
    settingsMaintenanceStatus.setAttribute('aria-live', 'polite');
    settingsMaintenanceSection.appendChild(settingsMaintenanceStatus);
  }

  return settingsMaintenanceStatus;
}

function setSettingsMaintenanceStatus(message = '', tone = 'default') {
  const statusElement = ensureSettingsMaintenanceStatus();

  if (!statusElement) {
    return;
  }

  statusElement.textContent = message;
  statusElement.classList.remove('is-success', 'is-error', 'is-busy');

  if (tone === 'success') {
    statusElement.classList.add('is-success');
  } else if (tone === 'error') {
    statusElement.classList.add('is-error');
  } else if (tone === 'busy') {
    statusElement.classList.add('is-busy');
  }
}

// Settings modal maintenance buttons are enabled or disabled from the current
// sidebar state so destructive actions stay predictable during testing.
function syncSettingsMaintenanceUi() {
  if (deleteCurrentMonthRegistrationsButton) {
    const hasSelection = isMonthSelection(currentSelection) && sidebarData.length > 0;

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

function syncSettingsUninstallUi() {
  if (settingsUninstallLaunchButton) {
    settingsUninstallLaunchButton.disabled = isImporting;
    settingsUninstallLaunchButton.setAttribute(
      'title',
      isImporting
        ? '処理中はアンインストールを開始できません'
        : 'アンインストールの確認を開く'
    );
  }

  if (uninstallAppButton) {
    uninstallAppButton.disabled = isImporting;
  }

  if (uninstallAppAndDeleteDataButton) {
    uninstallAppAndDeleteDataButton.disabled = isImporting;
  }
}

function initializeModalCloseIcons() {
  [
    imageModalClose,
    worldNameEditClose,
    photoLabelClose,
    settingsModalClose,
    uninstallModalClose,
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

  if (closeButton === uninstallModalClose) {
    event.preventDefault();
    event.stopPropagation();
    closeUninstallModal();
    return;
  }

  if (closeButton === trackedFolderModalClose) {
    event.preventDefault();
    event.stopPropagation();
    closeTrackedFolderModal();
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
        syncSinglePhotoUpdate(result.photo);
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

    modalPrintNoteHeroBadge = document.createElement('span');
    modalPrintNoteHeroBadge.className =
      'modal-world-label modal-print-note-hero-badge is-hidden';
    modalPrintNoteHeroBadge.textContent = 'プリント';
    badgeRow.appendChild(modalPrintNoteHeroBadge);

    if (!modalFavoriteButton) {
      modalFavoriteButton = document.createElement('button');
      modalFavoriteButton.id = 'modal-favorite-btn';
      modalFavoriteButton.type = 'button';
      modalFavoriteButton.className =
        'favorite-toggle-button modal-hero-favorite-button';
      modalFavoriteButton.setAttribute('aria-label', 'お気に入り切り替え');
      modalFavoriteButton.setAttribute('title', 'お気に入り切り替え');

      modalFavoriteIcon = document.createElement('span');
      modalFavoriteIcon.id = 'modal-favorite-icon';
      modalFavoriteIcon.className = 'material-symbols-outlined';
      modalFavoriteIcon.textContent = 'star';
      modalFavoriteButton.appendChild(modalFavoriteIcon);
    } else {
      modalFavoriteButton.classList.add('modal-hero-favorite-button');
    }

    // Keep the favorite toggle at the far left of the badge row.
    badgeRow.insertBefore(modalFavoriteButton, modalWorldLabel);
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
    openWorldNameEditButton.textContent = 'カードを編集';
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

function setWorldNameEditStatus(message = '') {
  if (worldNameSaveStatus) {
    worldNameSaveStatus.textContent = message;
  }
}

function setModalPhotoMemoStatus(message = '') {
  if (modalPhotoMemoStatus) {
    modalPhotoMemoStatus.textContent = message;
  }
}

function setModalPhotoMemoSaveButtonBusy(isBusy) {
  if (!modalPhotoMemoSaveButton) {
    return;
  }

  modalPhotoMemoSaveButton.disabled =
    Boolean(isBusy) || !Boolean(currentModalPhoto?.id);
}

function buildActionFailureMessage(actionLabel, result) {
  return `${actionLabel}: ${result?.message || '不明なエラー'}`;
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
    pickerTitle.textContent = 'ラベルを設定する';
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
    newTitle.textContent = 'ラベルを作成する';
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
    photoLabelNewNameInput.placeholder = 'ラベルの名前を入力してください';
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

function initializeModalPrintNoteUi() {
  if (modalPrintNoteBlock || !modalPhotoLabelsBlock || !modalPhotoMemoBlock) {
    return;
  }

  modalPrintNoteBlock = document.createElement('div');
  modalPrintNoteBlock.className =
    'modal-world-meta-block modal-print-note-block is-hidden';

  const title = document.createElement('p');
  title.className = 'modal-world-meta-title';
  title.textContent = 'プリントのノート';
  modalPrintNoteBlock.appendChild(title);

  modalPrintNoteValue = document.createElement('p');
  modalPrintNoteValue.className =
    'modal-world-description modal-print-note-value';
  modalPrintNoteBlock.appendChild(modalPrintNoteValue);

  modalPhotoLabelsBlock.insertAdjacentElement('afterend', modalPrintNoteBlock);
}

async function loadTrackedFoldersForSettings() {
  trackedFolders = await window.electronAPI.getTrackedFolders();
  renderTrackedFolderList();
}

// Settings modal updates currently come from several entry points (open,
// folder add/remove, maintenance actions). Keep the refresh steps in one place
// so future UI changes only need to update a single helper.
async function refreshSettingsModalUi({
  loadTrackedFolders = false,
  loadOverview = true,
  resetScroll = false,
  resetMaintenanceStatus = false,
} = {}) {
  ensureSettingsOverviewSection();
  ensureSettingsBackgroundSection();
  initializeSettingsTrackedFolderUi();
  ensureSettingsMaintenanceStatus();

  if (loadTrackedFolders) {
    await loadTrackedFoldersForSettings();
  } else {
    renderTrackedFolderList();
  }

  syncTrackedFolderSettingsActionsUi();
  syncSettingsBackgroundUi();
  syncSelectionDependentSettingsUi();

  if (loadOverview) {
    await loadSettingsOverview();
  }

  if (resetMaintenanceStatus) {
    setSettingsMaintenanceStatus('');
  }

  if (resetScroll && settingsModalBody) {
    settingsModalBody.scrollTop = 0;
  }
}

function initializeTopToolbarLayout() {
  // Toolbar layout is assembled at runtime so the static HTML can stay simple
  // and fragile header IDs do not need to move around in index.html.
  if (refreshTrackedFoldersButton && pageHeaderActions && settingsButton) {
    refreshTrackedFoldersButton.classList.add('theme-toggle-btn');
    refreshTrackedFoldersButton.innerHTML =
      '<span class="material-symbols-outlined">sync</span>';
    refreshTrackedFoldersButton.setAttribute('aria-label', '更新');
    refreshTrackedFoldersButton.setAttribute('title', '更新');
    pageHeaderActions.insertBefore(refreshTrackedFoldersButton, settingsButton);
  }

  if (worldNameFilterDropdown && toolbar && toolbarRight) {
    // The world-name filter behaves like a persistent toolbar input instead of
    // a transient dropdown. We relocate the existing block rather than
    // duplicating markup so all renderer bindings keep working.
    let toolbarLeftGroup = toolbar.querySelector('.toolbar-left-group');

    if (!toolbarLeftGroup) {
      toolbarLeftGroup = document.createElement('div');
      toolbarLeftGroup.className = 'toolbar-left-group';
      toolbar.insertBefore(toolbarLeftGroup, toolbarRight);
    }

    worldNameFilterDropdown.classList.add('toolbar-world-filter');
    worldNameFilterDropdown.classList.add('is-static-toolbar-filter');
    toolbarLeftGroup.appendChild(worldNameFilterDropdown);

    const inputPanel = worldNameFilterMenu?.querySelector(
      '.header-dropdown-input-panel'
    );

    if (inputPanel && !toolbarSearchScopeDropdown) {
      toolbarSearchScopeDropdown = document.createElement('div');
      toolbarSearchScopeDropdown.className =
        'header-dropdown toolbar-search-scope-dropdown';
      inputPanel.insertBefore(toolbarSearchScopeDropdown, worldNameFilterInput);

      toolbarSearchScopeButton = document.createElement('button');
      toolbarSearchScopeButton.type = 'button';
      toolbarSearchScopeButton.className =
        'header-dropdown-clear-button toolbar-search-scope-button';
      toolbarSearchScopeButton.setAttribute('aria-haspopup', 'menu');
      toolbarSearchScopeButton.setAttribute('aria-expanded', 'false');
      toolbarSearchScopeDropdown.appendChild(toolbarSearchScopeButton);

      toolbarSearchScopeMenu = document.createElement('div');
      toolbarSearchScopeMenu.className =
        'header-dropdown-menu toolbar-search-scope-menu';
      toolbarSearchScopeMenu.setAttribute('role', 'menu');
      toolbarSearchScopeMenu.hidden = true;
      toolbarSearchScopeDropdown.appendChild(toolbarSearchScopeMenu);
    }

    if (worldNameFilterSearchButton) {
      worldNameFilterSearchButton.textContent = '検索';
      worldNameFilterSearchButton.classList.add('toolbar-search-submit-button');
    }

    if (inputPanel && !toolbarSearchClearButton) {
      toolbarSearchClearButton = document.createElement('button');
      toolbarSearchClearButton.type = 'button';
      toolbarSearchClearButton.className =
        'header-dropdown-clear-button toolbar-search-clear-button';
      toolbarSearchClearButton.textContent = 'クリア';

      if (worldNameFilterSearchButton?.nextSibling) {
        inputPanel.insertBefore(
          toolbarSearchClearButton,
          worldNameFilterSearchButton.nextSibling
        );
      } else {
        inputPanel.appendChild(toolbarSearchClearButton);
      }
    }

    ensureSidebarWorldSortControls();

    if (!worldLibraryModeButton) {
      worldLibraryModeButton = document.createElement('button');
      worldLibraryModeButton.type = 'button';
      worldLibraryModeButton.className = 'small-action-button world-library-mode-btn';
    }

    if (worldLibraryModeButton && sidebarHeader) {
      const sidebarModeButtonAnchor = sidebarHeaderTitle || sidebarHeader.firstChild;

      if (worldLibraryModeButton.parentElement !== sidebarHeader) {
        if (sidebarModeButtonAnchor) {
          sidebarHeader.insertBefore(worldLibraryModeButton, sidebarModeButtonAnchor);
        } else {
          sidebarHeader.appendChild(worldLibraryModeButton);
        }
      }
    }

    ensureStaticToolbarWorldFilterVisible();
    renderToolbarSearchScopeMenu();
    syncToolbarSearchInputUi();
    syncSidebarModeUi();
  }

  if (regenerateThumbnailsButton && settingsMaintenanceSection) {
    // Settings owns thumbnail regeneration so the toolbar stays focused on
    // browsing/searching. The month selector is created once and reused.
    const utilityActions = ensureSettingsUtilityActionsContainer('thumbnails');
    ensureRegenerateThumbnailMonthDropdown(utilityActions);

    regenerateThumbnailsButton.classList.remove('secondary-toolbar-button');
    regenerateThumbnailsButton.classList.add('small-action-button');
    utilityActions.appendChild(regenerateThumbnailsButton);
    renderRegenerateThumbnailMonthOptions();
  }

  if (reimportRegisteredPhotosButton && settingsMaintenanceSection) {
    const utilityActions = ensureSettingsUtilityActionsContainer('reimport');
    ensureReimportRegisteredPhotoMonthDropdown(utilityActions);

    reimportRegisteredPhotosButton.classList.remove(
      'secondary',
      'settings-maintenance-button'
    );
    reimportRegisteredPhotosButton.classList.add('small-action-button');
    utilityActions.appendChild(reimportRegisteredPhotosButton);
    renderReimportRegisteredPhotoMonthOptions();
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

function openUninstallModal() {
  if (!uninstallModal || isImporting) {
    return;
  }

  syncSettingsUninstallUi();
  openSubModalElement(uninstallModal);
}

function closeUninstallModal() {
  closeSubModalElement(uninstallModal);
}

async function runUninstallFlow({ deleteData = false } = {}) {
  if (isImporting) {
    return;
  }

  const confirmed = await openConfirmModal({
    title: deleteData ? 'データも削除してアンインストール' : 'アンインストール',
    message: '本当に削除しますか？',
    confirmText: deleteData ? '削除してアンインストール' : 'アンインストール',
  });

  if (!confirmed) {
    return;
  }

  const result = deleteData
    ? await window.electronAPI.uninstallAppAndDeleteData()
    : await window.electronAPI.uninstallApp();

  if (!result?.ok) {
    showToast(result?.message || 'アンインストールを開始できませんでした');
    return;
  }

  showToast(
    deleteData
      ? 'データ削除とアンインストールを開始します'
      : 'アンインストールを開始します'
  );

  closeUninstallModal();
  closeSettingsModal();
}

async function openSettingsModal() {
  if (!settingsModal) {
    return;
  }

  await refreshSettingsModalUi({
    loadTrackedFolders: true,
    loadOverview: true,
    resetScroll: true,
    resetMaintenanceStatus: true,
  });

  openSubModalElement(settingsModal);
}

function closeSettingsModal() {
  closeSubModalElement(settingsModal, {
    onClosed: () => {
      closeUninstallModal();
      closeTrackedFolderModal();
      closeRegenerateThumbnailMonthMenu();
      closeReimportRegisteredPhotoMonthMenu();
      if (settingsModalBody) {
        settingsModalBody.scrollTop = 0;
      }
    },
  });
}

function syncSelectionDependentSettingsUi() {
  renderRegenerateThumbnailMonthOptions();
  renderReimportRegisteredPhotoMonthOptions();
  syncSettingsUtilityActionsUi();
  syncSettingsMaintenanceUi();
  syncSettingsUninstallUi();
}

function resetCurrentMonthState() {
  setCurrentSelectionValue(null);
  currentPhotos = [];
  allCurrentMonthPhotos = [];
}

function clearMainContent() {
  resetCurrentMonthState();
  setAnimatedMonthLabelText('写真一覧', { animate: false });
  setAnimatedMonthCountText('0枚', { animate: false });
  monthGalleryList.innerHTML = '';
  monthGalleryEmpty.style.display = 'block';
  monthGalleryEmpty.textContent =
    sidebarData.length > 0
      ? getDefaultSelectionEmptyMessage()
      : getDefaultMonthGalleryEmptyMessage();
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
  card.tabIndex = -1;

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

    togglePhotoSelection(item.id, { rangeSelect: event.shiftKey });
  });

  card.addEventListener('pointerdown', (event) => {
    if (
      !isSelectionMode ||
      event.button !== 0 ||
      event.shiftKey ||
      event.target.closest('.photo-card-selection-btn, .photo-card-favorite-btn')
    ) {
      return;
    }

    event.preventDefault();
    beginSelectionDrag(item.id, event.pointerId);
  });

  card.addEventListener('pointerenter', () => {
    applySelectionDragToPhoto(item.id);
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

  card.addEventListener('click', (event) => {
    if (isSelectionMode) {
      if (suppressSelectionModeCardClickPhotoId === item.id) {
        suppressSelectionModeCardClickPhotoId = null;
        return;
      }

      togglePhotoSelection(item.id, {
        rangeSelect: event.shiftKey,
      });
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
    currentSelection.mode || 'month',
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
  syncKeyboardFocusedPhotoCard();
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
    scheduleMainHeaderResponsiveLayout();
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
    monthGalleryEmpty.textContent = getDefaultSelectionEmptyMessage();
    return;
  }

  setAnimatedMonthLabelText(getSelectionLabelText());
  // This is the canonical render path for the active selection view:
  // header count, empty state, and filter button text are all synchronized here.
  syncFavoriteFilterUi();

  if (currentPhotos.length === 0) {
    resetMonthGalleryRenderState();
    monthGalleryEmpty.style.display = 'block';
    monthGalleryEmpty.textContent =
      allCurrentMonthPhotos.length > 0 && isAnyPhotoFilterActive()
        ? buildFilteredEmptyMessage()
        : getDefaultSelectionEmptyMessage();

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
  syncKeyboardFocusedPhotoCard();
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
  const appScrollTop = appRoot?.scrollTop || 0;
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

  if (appScrollTop > 0) {
    targets.push({
      start: appScrollTop,
      apply(value) {
        if (appRoot) {
          appRoot.scrollTop = value;
        }
      },
    });
  } else if (documentScrollTop > 0) {
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
  syncSidebarModeUi();

  if (currentSidebarMode === 'world') {
    if (worldSidebarData.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'sidebar-empty';
      empty.textContent = 'ワールド情報付きの写真はまだありません';
      sidebarTree.appendChild(empty);
      return;
    }

    for (const worldEntry of worldSidebarData) {
      const worldButton = document.createElement('button');
      worldButton.type = 'button';
      worldButton.className = 'world-sidebar-item';
      worldButton.dataset.worldKey = worldEntry.worldKey;

      const worldName = document.createElement('span');
      worldName.className = 'world-sidebar-item-name';
      worldName.textContent = worldEntry.worldName;

      const worldCount = document.createElement('span');
      worldCount.className = 'world-sidebar-item-count';
      worldCount.textContent = `${worldEntry.count}枚`;

      if (
        isWorldSelection(currentSelection) &&
        currentSelection.worldKey === worldEntry.worldKey
      ) {
        worldButton.classList.add('active');
      }

      worldButton.append(worldName, worldCount);
      worldButton.addEventListener('click', async () => {
        await selectWorld(worldEntry);
      });

      sidebarTree.appendChild(worldButton);
    }

    return;
  }

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
    label.textContent = String(yearEntry.year);

    yearLeft.appendChild(toggle);
    yearLeft.appendChild(label);

    const yearCount = document.createElement('span');
    yearCount.className = 'year-count';
    yearCount.textContent = `${yearEntry.totalCount}枚`;

    yearButton.appendChild(yearLeft);
    yearButton.appendChild(yearCount);

    toggle.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();

      if (expandedYears.has(yearEntry.year)) {
        expandedYears.delete(yearEntry.year);
      } else {
        expandedYears.add(yearEntry.year);
      }
      renderSidebar();
    });

    yearButton.addEventListener('click', async () => {
      await selectYear(yearEntry.year);
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
      monthName.textContent = pad2(monthEntry.month);

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
  const [nextSidebarData, nextWorldSidebarData] = await Promise.all([
    window.electronAPI.getSidebarData(),
    window.electronAPI.getWorldSidebarData(currentWorldSidebarSort),
  ]);

  sidebarData = nextSidebarData;
  worldSidebarData = Array.isArray(nextWorldSidebarData)
    ? nextWorldSidebarData
    : [];
  syncSelectionLinkedUi({ forceSidebarRender: true });
}

// Month/year switching shares the same fetch -> selection -> render pipeline.
// Only the fetcher and resulting normalized selection differ.
async function selectPhotoScope(fetchPhotos, nextSelection) {
  const requestId = ++monthSelectionRequestId;
  clearMonthHeaderAnimationState();
  const monthSwitchOverlay = createMonthSwitchOverlay({
    includeHeader: false,
    viewportFixed: true,
  });
  const photosPromise = fetchPhotos();
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

  setCurrentSelectionValue(nextSelection);

  setCurrentMonthPhotos(photos);

  syncSelectionLinkedUi();
  stopScrollToTopAnimation();
  scrollGalleryViewToTop({ animated: false });
  renderMonthGallery({ resetProgressive: true });
  playMonthSwitchAnimation({ includeHeader: false });
  requestAnimationFrame(() => {
    fadeOutMonthSwitchOverlay(monthSwitchOverlay);
  });
}

// Year selection reuses the same gallery/filter pipeline as month selection,
// but fetches a broader photo set and updates the header label accordingly.
async function selectYear(year) {
  await selectPhotoScope(
    () => window.electronAPI.getPhotosByYear(year),
    createYearSelection(year)
  );
}

async function selectWorld(worldEntry) {
  const nextSelection = createWorldSelection(
    worldEntry.worldKey,
    worldEntry.worldName,
    worldEntry.worldId
  );

  await selectPhotoScope(
    () => window.electronAPI.getPhotosByWorldSelection(nextSelection),
    nextSelection
  );
}

async function selectMonth(year, month) {
  await selectPhotoScope(
    () => window.electronAPI.getPhotosByMonth(year, month),
    createMonthSelection(year, month)
  );
}

async function selectCurrentSelection(selection = currentSelection) {
  const normalizedSelection = normalizeSelection(selection);

  if (!normalizedSelection) {
    return false;
  }

  if (normalizedSelection.mode === 'world') {
    const matchingWorldEntry = worldSidebarData.find(
      (worldEntry) => worldEntry.worldKey === normalizedSelection.worldKey
    );

    if (!matchingWorldEntry) {
      return false;
    }

    await selectWorld(matchingWorldEntry);
  } else if (normalizedSelection.mode === 'year') {
    await selectYear(normalizedSelection.year);
  } else {
    await selectMonth(normalizedSelection.year, normalizedSelection.month);
  }

  return true;
}

async function handleImportResult(result, modeLabel) {
  importStatus.textContent = buildImportStatusMessage(result, modeLabel);

  if (!result || result.canceled) {
    return;
  }

  if (result.failedCount > 0) {
    showToast(`${modeLabel}: ${result.failedCount}件失敗しました`);
  }

  await restorePhotoDataSelectionFromResult(result);
}

// Import / refresh flows all converge here so the same selection-restore rules
// are used regardless of whether data changed via import, tracked-folder
// refresh, or another foreground sync path.
async function queueWorldMetadataSyncForResult(result) {
  await startBackgroundWorldMetadataSync(result?.worldMetadataTargets);
}

async function restorePhotoDataSelectionFromResult(
  result,
  fallbackSelection = null
) {
  await restoreSidebarAndMonthSelection({
    preferredSelection: result?.selectedMonth || null,
    fallbackSelection,
  });
}

async function restoreSidebarAndMonthSelection({
  preferredSelection = null,
  fallbackSelection = null,
} = {}) {
  await refreshSidebar();
  await restoreMonthViewAfterDataChange({
    preferredSelection,
    fallbackSelection,
  });
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

// Refresh results have more branches than import results, so toast selection is
// separated from sidebar/month restoration to keep the main handler readable.
function showTrackedFoldersRefreshResultToast(result) {
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
    return;
  }

  if (result.missingFolderPaths?.length > 0) {
    showToast(`見つからないフォルダが${result.missingFolderPaths.length}件あります`);
    return;
  }

  if (!result.emptyRefresh && (result.importedCount || 0) > 0) {
    showToast('追跡フォルダを更新しました');
    return;
  }

  showToast('追跡フォルダを確認しました');
}

async function handleTrackedFoldersRefreshResult(result, fallbackSelection) {
  importStatus.textContent = buildTrackedFoldersRefreshMessage(result);

  if (!result || result.canceled) {
    return;
  }

  showTrackedFoldersRefreshResultToast(result);

  if (result.ok === false || result.noTrackedFolders) {
    return;
  }

  if (result.emptyRefresh) {
    return;
  }

  await restorePhotoDataSelectionFromResult(result, fallbackSelection);
}

async function runRegenerateThumbnailsFlow(targetYear, targetMonth) {
  setSettingsMaintenanceStatus('サムネイルを再生成しています...', 'busy');
  await runForegroundAsyncAction({
    statusMessage: 'サムネイルを再生成しています...',
    progressMessage: 'サムネイルを再生成しています...',
    run: () =>
      window.electronAPI.regenerateThumbnails({
        year: targetYear,
        month: targetMonth,
      }),
    handleResult: async (result) => {
      const successStatus = buildScopedRegenerateThumbnailsMessage(result);
      importStatus.textContent = successStatus;
      setSettingsMaintenanceStatus(successStatus, 'success');

      if (result?.failedCount > 0) {
        showToast(`サムネイル再生成 ${result.failedCount}件失敗しました`);
      } else if (result?.ok) {
        showToast('サムネイル再生成が完了しました');
      }

      await selectCurrentSelection();
    },
    buildErrorStatus: (message) => {
      const errorStatus = `サムネイル再生成に失敗しました: ${message}`;
      setSettingsMaintenanceStatus(errorStatus, 'error');
      return errorStatus;
    },
  });
}

async function runTrackedFoldersRefreshFlow() {
  const fallbackSelection = currentSelection ? { ...currentSelection } : null;
  const result = await runForegroundAsyncAction({
    statusMessage: '追跡フォルダを更新中...',
    progressMessage: '追跡フォルダを更新中...',
    run: () => window.electronAPI.refreshTrackedFolders(),
    handleResult: (currentResult) =>
      handleTrackedFoldersRefreshResult(currentResult, fallbackSelection),
    buildErrorStatus: (message) => `更新に失敗しました: ${message}`,
    releaseBusyBeforeHandleResult: true,
  });

  await queueWorldMetadataSyncForResult(result);
}

function setElementDisabledState(element, disabled) {
  if (element) {
    element.disabled = Boolean(disabled);
  }
}

function syncBusyAffectedPrimaryActions(isBusy) {
  setElementDisabledState(refreshTrackedFoldersButton, isBusy);
  setElementDisabledState(settingsButton, isBusy);
  syncSettingsUtilityActionsUi();
  syncTrackedFolderSettingsActionsUi();
  syncTrackedFolderListActionButtonsUi();
}

function syncBusyAffectedFilterActions(isBusy) {
  const selectionDependentDisabled = isBusy || !currentSelection;

  setElementDisabledState(favoriteFilterButton, selectionDependentDisabled);
  setElementDisabledState(orientationFilterButton, selectionDependentDisabled);
  setElementDisabledState(worldNameFilterButton, selectionDependentDisabled);
  setElementDisabledState(worldNameFilterInput, selectionDependentDisabled);
  setElementDisabledState(worldNameFilterSearchButton, selectionDependentDisabled);
  setElementDisabledState(toolbarSearchClearButton, selectionDependentDisabled);
  setElementDisabledState(toolbarSearchScopeButton, selectionDependentDisabled);
  setElementDisabledState(selectionModeButton, selectionDependentDisabled);
}

function syncBusyAffectedSelectionActions(isBusy) {
  setElementDisabledState(
    bulkFavoriteButton,
    isBusy || !isSelectionMode || selectedPhotoIds.size === 0
  );
  setElementDisabledState(
    bulkDeleteButton,
    isBusy || !isSelectionMode || selectedPhotoIds.size === 0
  );
}

function syncBusyAffectedMaintenanceActions(isBusy) {
  setElementDisabledState(
    clearThumbnailCacheButton,
    isBusy || sidebarData.length === 0
  );
  setElementDisabledState(
    resetDatabaseButton,
    isBusy || (sidebarData.length === 0 && trackedFolders.length === 0)
  );
}

function closeMenusBlockedByBusyState() {
  closeOrientationFilterMenu();
  closeWorldNameFilterMenu();
  closeRegenerateThumbnailMonthMenu();
}

function setImportUiBusy(isBusy) {
  isImporting = isBusy;

  if (!isBusy) {
    resetProcessingProgress();
  }

  syncBusyAffectedPrimaryActions(isBusy);
  syncBusyAffectedFilterActions(isBusy);
  syncBusyAffectedSelectionActions(isBusy);
  syncBusyAffectedMaintenanceActions(isBusy);

  if (isBusy && typeof resetDropOverlay === 'function') {
    resetDropOverlay();
  }

  if (isBusy) {
    closeMenusBlockedByBusyState();
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

async function runForegroundAsyncAction({
  guardMessage,
  statusMessage,
  progressMessage = '',
  showProgress = true,
  run,
  handleResult = null,
  buildErrorStatus,
  releaseBusyBeforeHandleResult = false,
}) {
  if (isImporting) {
    if (guardMessage) {
      showToast(guardMessage);
    }
    return null;
  }

  let result = null;
  let didReleaseBusyEarly = false;

  beginForegroundProgressOperation({
    statusMessage,
    progressMessage,
    showProgress,
  });

  try {
    result = await run();

    if (releaseBusyBeforeHandleResult) {
      setImportUiBusy(false);
      didReleaseBusyEarly = true;
    }

    if (typeof handleResult === 'function') {
      await handleResult(result);
    }

    return result;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    importStatus.textContent =
      typeof buildErrorStatus === 'function'
        ? buildErrorStatus(message)
        : message;
    return null;
  } finally {
    if (!didReleaseBusyEarly) {
      setImportUiBusy(false);
    }
  }
}


async function runImportFlow(modeLabel, startMessage, importRunner) {
  const result = await runForegroundAsyncAction({
    guardMessage: '取り込み中です。処理が終わってから次の取り込みを開始してください',
    statusMessage: startMessage,
    progressMessage: startMessage,
    run: importRunner,
    handleResult: (currentResult) => handleImportResult(currentResult, modeLabel),
    buildErrorStatus: (message) => `取り込みに失敗しました: ${message}`,
  });

  await queueWorldMetadataSyncForResult(result);
}

// Manual world editing stays isolated from the read-only modal so save/reload
// behavior can evolve without changing the primary photo modal bindings.
async function saveManualWorldEditForm({
  worldNameManual,
  worldUrl,
} = {}) {
  if (!currentModalPhoto) {
    return;
  }

  setWorldNameEditStatus('保存中...');

  const result = await window.electronAPI.updateWorldSettings(
    currentModalPhoto.id,
    {
      worldNameManual,
      worldUrl,
    }
  );

  if (!result?.ok) {
    setWorldNameEditStatus(buildActionFailureMessage('保存に失敗しました', result));
    return;
  }

  syncSinglePhotoUpdate(result.photo);
  closeWorldNameEditModal();
  showToast('World設定を保存しました');
}

// Memo saves only touch the currently open photo, but they still go through the
// shared single-photo sync path so cards, filters, and modal state stay aligned.
async function savePhotoMemo() {
  if (!currentModalPhoto || !modalPhotoMemoInput) {
    return;
  }

  setModalPhotoMemoStatus('保存中...');
  setModalPhotoMemoSaveButtonBusy(true);

  const result = await window.electronAPI.updatePhotoMemo(
    currentModalPhoto.id,
    modalPhotoMemoInput.value
  );

  if (!result?.ok) {
    setModalPhotoMemoStatus(
      buildActionFailureMessage('保存に失敗しました', result)
    );
    setModalPhotoMemoSaveButtonBusy(false);
    return;
  }

  syncSinglePhotoUpdate(result.photo);

  setModalPhotoMemoStatus('保存しました');
  setModalPhotoMemoSaveButtonBusy(false);

  showToast('メモを保存しました');
}

// Manual reload uses the edit modal's current World URL input so users do not
// need to save first just to test a corrected URL.
async function rereadWorldName() {
  if (!currentModalPhoto) {
    return;
  }

  setWorldNameEditStatus('再取得中...');

  const result = await window.electronAPI.rereadWorldName({
    photoId: currentModalPhoto.id,
    worldUrl: modalWorldUrlInput?.value || currentModalPhoto.worldUrl || '',
  });

  if (!result || !result.ok) {
    setWorldNameEditStatus(
      buildActionFailureMessage('再取得に失敗しました', result)
    );
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

  setCurrentMonthPhotos(
    allCurrentMonthPhotos.map((photo) => updatedPhotoMap.get(photo.id) || photo)
  );
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
    if (hasSidebarSelection(targetSelection)) {
      const isCurrentTarget = isSameSelection(currentSelection, targetSelection);

      if (preferLocalRender && isCurrentTarget) {
        const removedLocally = removeRenderedPhotoCards(removedPhotoIds);

        if (!removedLocally) {
          renderMonthGallery({ resetProgressive: true });
        } else {
          syncFavoriteFilterUi();
        }
        return;
      }

      await restoreMonthViewAfterDataChange({
        preferredSelection: targetSelection,
        clearWhenEmpty: false,
      });
      return;
    }
  }

  resetCurrentMonthState();
  clearSelectionState();
  await restoreMonthViewAfterDataChange();
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
  setSettingsMaintenanceStatus(busyStatus, 'busy');

  try {
    const result = await run();

    if (!result?.ok) {
      throw new Error(result?.message || '処理に失敗しました');
    }

    if (typeof onSuccess === 'function') {
      await onSuccess(result);
    }

    if (typeof buildSuccessStatus === 'function') {
      const successStatus = buildSuccessStatus(result);
      importStatus.textContent = successStatus;
      setSettingsMaintenanceStatus(successStatus, 'success');
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
      const errorStatus = buildErrorStatus(message);
      importStatus.textContent = errorStatus;
      setSettingsMaintenanceStatus(errorStatus, 'error');
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
    void refreshSettingsModalUi({
      loadTrackedFolders: false,
      loadOverview: true,
    });
  }
}

// These maintenance actions intentionally reuse the normal delete / refresh
// flows so verification work exercises the same data paths as day-to-day use.
async function deleteCurrentMonthRegistrationsFromSettings() {
  const targetSelection = { ...currentSelection };

  await runSettingsMaintenanceAction({
    isBlocked: () =>
      isImporting ||
      !isMonthSelection(currentSelection),
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

async function reimportRegisteredPhotosFromSettings(targetYear, targetMonth) {
  const targetSelection = createMonthSelection(targetYear, targetMonth);
  const fallbackSelection = currentSelection ? { ...currentSelection } : null;

  const result = await runSettingsMaintenanceAction({
    isBlocked: () =>
      isImporting ||
      sidebarData.length === 0 ||
      !window.electronAPI.reimportRegisteredPhotos,
    confirmOptions: {
      title: '既存画像の情報を再取り込み',
      message:
        `${targetSelection.year}年${targetSelection.month}月の登録済み画像から現在の解析ロジックで画像情報を再取得します。World情報、プリントのノート、解像度などは更新されますが、メモ・ラベル・手動のWorld名は保持されます。続行しますか？`,
      confirmText: '再取り込みする',
    },
    busyStatus: `${targetSelection.year}年${targetSelection.month}月の情報を再取り込み中...`,
    progressMessage: '既存画像の情報を再取り込み中...',
    run: () => window.electronAPI.reimportRegisteredPhotos(targetSelection),
    onSuccess: async (currentResult) => {
      await restorePhotoDataSelectionFromResult(currentResult, fallbackSelection);
    },
    buildSuccessStatus: (currentResult) => {
      if (currentResult.emptyReimport) {
        return `${targetSelection.year}年${targetSelection.month}月: 再取り込み対象の登録画像はありません`;
      }

      return [
        `${targetSelection.year}年${targetSelection.month}月: ${
          currentResult.importedCount || 0
        }件反映`,
        currentResult.updatedCount > 0
          ? `更新 ${currentResult.updatedCount}件`
          : null,
        currentResult.failedCount > 0
          ? `失敗 ${currentResult.failedCount}件`
          : null,
      ]
        .filter(Boolean)
        .join(' / ');
    },
    buildSuccessToast: (currentResult) => {
      if (currentResult.failedCount > 0) {
        return `再取り込み: ${currentResult.failedCount}件失敗しました`;
      }

      return `${targetSelection.year}年${targetSelection.month}月の情報を再取り込みしました`;
    },
    buildErrorStatus: (message) => `再取り込みに失敗しました: ${message}`,
    buildErrorToast: (message) => `再取り込みに失敗しました: ${message}`,
  });

  await queueWorldMetadataSyncForResult(result);
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

function finishAppInitialization() {
  clearTimeout(appInitializationFailsafeTimer);
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      appRoot?.classList.remove('is-app-initializing');
      appRoot?.setAttribute('aria-busy', 'false');
    });
  });
}

async function initializeApp() {
  await refreshSidebar();
  const restored = await restoreMonthViewAfterDataChange();

  if (!restored) {
    renderSidebar();
  }

  syncFavoriteFilterUi();
}

async function runRendererStartupStep(label, step) {
  try {
    return await step();
  } catch (error) {
    console.error(`[renderer startup] ${label} failed`, error);
    return null;
  }
}

async function bootstrapRenderer() {
  const initializationTimeoutMs = 8000;
  let timeoutId = null;

  try {
    await runRendererStartupStep('initializeRendererUi', async () => {
      initializeRendererUi();
    });
    await runRendererStartupStep('initializeRendererBindings', async () => {
      initializeRendererBindings();
    });
    await runRendererStartupStep('syncFavoriteFilterUi', async () => {
      syncFavoriteFilterUi();
    });
    await runRendererStartupStep('initializeApp', async () => {
      await Promise.race([
        initializeApp(),
        new Promise((_, reject) => {
          timeoutId = setTimeout(() => {
            reject(new Error('renderer app initialization timed out'));
          }, initializationTimeoutMs);
        }),
      ]);
    });
  } finally {
    if (timeoutId) {
      clearTimeout(timeoutId);
    }
    finishAppInitialization();
  }
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
      await runRegenerateThumbnailsFlow(targetYear, targetMonth);
    },
    { capture: true }
  );

  refreshTrackedFoldersButton?.addEventListener('click', async () => {
    if (isImporting) {
      showToast('別の処理中です。完了してから更新してください');
      return;
    }

    await runTrackedFoldersRefreshFlow();
  });
}

// Header filters stay interactive while month content changes underneath them.
function bindHeaderFilterControls() {
  favoriteFilterButton?.addEventListener('click', async () => {
    if (!currentSelection) {
      return;
    }

    isFavoriteFilterOnly = !isFavoriteFilterOnly;
    await syncCurrentPhotoFilterPresentation();
  });

  photoSortButton?.addEventListener('click', async () => {
    if (!currentSelection || isImporting) {
      return;
    }

    currentPhotoSortOrder = currentPhotoSortOrder === 'asc' ? 'desc' : 'asc';
    setCurrentMonthPhotos(allCurrentMonthPhotos);
    await syncCurrentPhotoFilterPresentation({ animate: false });
  });

  photoDensityButton?.addEventListener('click', () => {
    playPhotoCardDensityTransition();
    applyPhotoCardDensityPreference(
      currentPhotoCardDensity === 'compact' ? 'default' : 'compact'
    );
    syncPhotoCardDensityUi();
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

  toolbarSearchScopeButton?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();

    if (!currentSelection || isImporting) {
      return;
    }

    setToolbarSearchScopeMenuOpen(!isToolbarSearchScopeMenuOpen);
  });

  toolbarSearchScopeMenu?.addEventListener('click', async (event) => {
    const target = event.target.closest('[data-toolbar-search-scope]');

    if (!target) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    await setToolbarSearchScope(target.dataset.toolbarSearchScope || 'world');
  });

  worldNameFilterInput?.addEventListener('input', () => {
    clearWorldNameFilterInputTimer();
    syncToolbarSearchInputUi();
  });

  worldNameFilterInput?.addEventListener('keydown', async (event) => {
    if (event.key !== 'Enter' || event.isComposing || event.keyCode === 229) {
      return;
    }

    event.preventDefault();
    await submitWorldNameFilter({ focusCards: true });
  });

  worldNameFilterSearchButton?.addEventListener('click', async (event) => {
    event.preventDefault();
    event.stopPropagation();
    await submitWorldNameFilter({ focusCards: true });
  });

  toolbarSearchClearButton?.addEventListener('click', async (event) => {
    event.preventDefault();
    event.stopPropagation();
    await clearWorldNameFilter({ keepFocus: true });
  });

  worldLibraryModeButton?.addEventListener('click', async () => {
    await runSidebarModeSwitchTransition(async () => {
      currentSidebarMode = currentSidebarMode === 'world' ? 'timeline' : 'world';
      await refreshSidebar();

      if (currentSidebarMode === 'world') {
        const targetSelection =
          lastWorldSelection || getLatestWorldSelectionFromSidebarData();

        if (targetSelection) {
          await selectSidebarSelectionIfAvailable(targetSelection);
        } else {
          resetCurrentMonthState();
          clearSelectionState();
          clearMainContent();
          renderSidebar();
        }
        return;
      }

      await restoreMonthViewAfterDataChange({
        preferredSelection: lastTimelineSelection,
        fallbackSelection: getLatestSelectionFromSidebarData(),
      });
    });
  });

  sidebarSortCountButton?.addEventListener('click', async () => {
    if (currentWorldSidebarSort === 'count') {
      return;
    }

    currentWorldSidebarSort = 'count';
    await runSidebarTreeRefreshTransition(async () => {
      await refreshSidebar();

      if (currentSidebarMode === 'world') {
        await selectSidebarSelectionIfAvailable(
          lastWorldSelection || currentSelection || getLatestWorldSelectionFromSidebarData()
        );
      }
    });
  });

  sidebarSortNameButton?.addEventListener('click', async () => {
    if (currentWorldSidebarSort === 'name') {
      return;
    }

    currentWorldSidebarSort = 'name';
    await runSidebarTreeRefreshTransition(async () => {
      await refreshSidebar();

      if (currentSidebarMode === 'world') {
        await selectSidebarSelectionIfAvailable(
          lastWorldSelection || currentSelection || getLatestWorldSelectionFromSidebarData()
        );
      }
    });
  });
}

async function openCurrentModalWorldUrl() {
  if (!currentModalPhoto?.worldUrl) {
    return;
  }

  const result = await window.electronAPI.openExternalUrl(
    currentModalPhoto.worldUrl
  );

  if (!result?.ok) {
    showToast(`リンクを開けませんでした: ${result?.message || '不明なエラー'}`);
  }
}

function handleRecoveredModalFileActionResult(result, failureMessage) {
  if (result?.photo) {
    syncSinglePhotoUpdate(result.photo);
  }

  if (!result?.ok) {
    showToast(`${failureMessage}: ${result?.message || '不明なエラー'}`);
  }

  if (result?.recovered) {
    showToast('画像の保存場所を更新しました');
  }
}

async function openCurrentModalOriginalFile() {
  if (!currentModalPhoto?.filePath) {
    return;
  }

  const result = await window.electronAPI.openLocalFile({
    photoId: currentModalPhoto.id,
    filePath: currentModalPhoto.filePath,
  });

  handleRecoveredModalFileActionResult(result, '画像を開けませんでした');
}

async function openCurrentModalContainingFolder() {
  if (!currentModalPhoto?.filePath) {
    return;
  }

  const result = await window.electronAPI.openContainingFolder({
    photoId: currentModalPhoto.id,
    filePath: currentModalPhoto.filePath,
  });

  handleRecoveredModalFileActionResult(result, '保存先フォルダを開けませんでした');
}

async function handleSaveWorldSettingsClick() {
  await saveManualWorldEditForm({
    worldNameManual: modalWorldNameInput?.value || '',
    worldUrl: modalWorldUrlInput?.value || '',
  });
}

function handleWorldEditFormInput() {
  setWorldNameEditStatus('');
}

async function handlePhotoMemoSaveClick() {
  await savePhotoMemo();
}

function handlePhotoMemoInput() {
  resizeModalPhotoMemoInput();
  setModalPhotoMemoStatus('');
}

async function handlePhotoMemoKeydown(event) {
  if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
    event.preventDefault();
    await savePhotoMemo();
  }
}

async function handleClearWorldNameClick() {
  await saveManualWorldEditForm({
    worldNameManual: '',
    worldUrl: currentModalPhoto?.worldUrl || '',
  });
}

async function handleRereadWorldNameClick() {
  await rereadWorldName();
}

// Modal-level editors and detail actions are grouped here so photo-specific
// behavior can be traced without scanning the entire file bottom. The handlers
// above keep bind-time code declarative while preserving named entry points for
// future debugging.
function bindPhotoAndEditModalControls() {
  saveWorldNameButton?.addEventListener('click', handleSaveWorldSettingsClick);
  modalWorldNameInput?.addEventListener('input', handleWorldEditFormInput);
  modalWorldUrlInput?.addEventListener('input', handleWorldEditFormInput);
  modalPhotoMemoSaveButton?.addEventListener('click', handlePhotoMemoSaveClick);
  modalPhotoMemoInput?.addEventListener('input', handlePhotoMemoInput);
  modalPhotoMemoInput?.addEventListener('keydown', handlePhotoMemoKeydown);
  clearWorldNameButton?.addEventListener('click', handleClearWorldNameClick);
  rereadWorldNameButton?.addEventListener('click', handleRereadWorldNameClick);

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
    await openCurrentModalWorldUrl();
  });

  modalOpenWorldButton?.addEventListener('click', async () => {
    await openCurrentModalWorldUrl();
  });

  modalOpenOriginalButton?.addEventListener('click', async () => {
    await openCurrentModalOriginalFile();
  });

  modalOpenFolderButton?.addEventListener('click', async () => {
    await openCurrentModalContainingFolder();
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
  bindSubModalCloseTriggers(
    trackedFolderModalBackdrop,
    trackedFolderModalClose,
    closeTrackedFolderModal
  );
  bindSubModalCloseTriggers(
    uninstallModalBackdrop,
    uninstallModalClose,
    closeUninstallModal
  );
  settingsModalContent?.addEventListener('wheel', handleSettingsModalWheel, {
    passive: false,
  });

  openTrackedFolderListButton?.addEventListener('click', () => {
    openTrackedFolderModal();
  });

  settingsUninstallLaunchButton?.addEventListener('click', () => {
    openUninstallModal();
  });

  uninstallAppButton?.addEventListener('click', async () => {
    await runUninstallFlow({ deleteData: false });
  });

  uninstallAppAndDeleteDataButton?.addEventListener('click', async () => {
    await runUninstallFlow({ deleteData: true });
  });

  selectBackgroundImageButton?.addEventListener('click', async () => {
    await selectBackgroundImageFromSettings();
  });

  clearBackgroundImageButton?.addEventListener('click', async () => {
    await clearBackgroundImageFromSettings();
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
    await refreshSettingsModalUi({
      loadTrackedFolders: false,
      loadOverview: true,
      resetMaintenanceStatus: true,
    });

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
    await refreshSettingsModalUi({
      loadTrackedFolders: false,
      loadOverview: true,
      resetMaintenanceStatus: true,
    });
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

  reimportRegisteredPhotosButton?.addEventListener('click', async () => {
    const selectedMonthValue = reimportRegisteredPhotoMonthSelect?.value || '';

    if (!/^\d{4}-\d{2}$/.test(selectedMonthValue)) {
      showToast('再取り込みする月を選択してください');
      return;
    }

    const [targetYear, targetMonth] = selectedMonthValue.split('-').map(Number);
    await reimportRegisteredPhotosFromSettings(targetYear, targetMonth);
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
      lastSelectionAnchorPhotoId = null;
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

  document.addEventListener('pointerup', () => {
    finishSelectionDrag();
  });

  document.addEventListener('pointercancel', () => {
    finishSelectionDrag();
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

    if (trackedFolderModal && !trackedFolderModal.classList.contains('hidden')) {
      closeTrackedFolderModal();
      return;
    }

    if (uninstallModal && !uninstallModal.classList.contains('hidden')) {
      closeUninstallModal();
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

  document.addEventListener('keydown', (event) => {
    if (
      imageModal?.classList.contains('hidden') === false ||
      worldNameEditModal?.classList.contains('hidden') === false ||
      photoLabelModal?.classList.contains('hidden') === false ||
      trackedFolderModal?.classList.contains('hidden') === false ||
      settingsModal?.classList.contains('hidden') === false ||
      confirmModal?.classList.contains('hidden') === false ||
      isEditableKeyboardTarget(event.target)
    ) {
      return;
    }

    if (!currentSelection || getRenderedVisiblePhotoCards().length === 0) {
      return;
    }

    if (event.key === 'ArrowLeft') {
      event.preventDefault();
      moveKeyboardFocusedPhoto(-1);
      return;
    }

    if (event.key === 'ArrowRight') {
      event.preventDefault();
      moveKeyboardFocusedPhoto(1);
      return;
    }

    if (event.key === 'ArrowUp') {
      event.preventDefault();
      moveKeyboardFocusedPhotoVertical(-1);
      return;
    }

    if (event.key === 'ArrowDown') {
      event.preventDefault();
      moveKeyboardFocusedPhotoVertical(1);
      return;
    }

    if (event.key === 'Enter') {
      event.preventDefault();
      activateKeyboardFocusedPhoto();
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

  window.electronAPI.onAppUpdateStatus?.((payload) => {
    const message =
      typeof payload?.message === 'string' ? payload.message.trim() : '';

    if (!message) {
      return;
    }

    showToast(message);
  });

  window.electronAPI.onAppUpdateAction?.((payload) => {
    queueAppUpdatePrompt(payload);
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
  initializePhotoCardDensityPreference();
  initializeBackgroundImagePreference();
  initializeImageModalUi();
  initializeWorldNameEditUi();
  initializePhotoLabelUi();
  initializeModalPrintNoteUi();
  initializeModalCloseIcons();
  ensureSettingsBackgroundSection();
  initializeSettingsTrackedFolderUi();
  syncSettingsBackgroundUi();
  initializeTopToolbarLayout();
  initializeDragAndDropImport();
  initializeProgressiveMonthGalleryLoading();
  initializeScrollToTopAnimationInterrupts();
  scheduleMainHeaderResponsiveLayout();
}

bootstrapRenderer();


