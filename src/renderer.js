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

const RENDERER_PERFORMANCE_LOG_THRESHOLD_MS = 180;

function isRendererPerformanceLogForced() {
  try {
    return window.localStorage?.getItem('worldshotPerfLog') === '1';
  } catch {
    return false;
  }
}

function roundRendererDuration(durationMs) {
  return Math.round((Number(durationMs) || 0) * 10) / 10;
}

function logRendererPerformance(label, durationMs, details = {}) {
  if (
    !isRendererPerformanceLogForced() &&
    (!Number.isFinite(durationMs) ||
      durationMs < RENDERER_PERFORMANCE_LOG_THRESHOLD_MS)
  ) {
    return;
  }

  const normalizedDetails = Object.fromEntries(
    Object.entries(details || {}).map(([key, value]) => [
      key,
      typeof value === 'number' ? roundRendererDuration(value) : value,
    ])
  );

  console.info('[renderer performance]', label, {
    durationMs: roundRendererDuration(durationMs),
    ...normalizedDetails,
  });
}

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
const modalEditPhotoButton = document.getElementById('modal-edit-photo-btn');
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

// Photo editor modal.
const photoEditorModal = document.getElementById('photo-editor-modal');
const photoEditorBackdrop = document.getElementById('photo-editor-backdrop');
const photoEditorClose = document.getElementById('photo-editor-close');
const photoEditorFileName = document.getElementById('photo-editor-file-name');
const photoEditorResetButton = document.getElementById('photo-editor-reset-btn');
const photoEditorSaveButton = document.getElementById('photo-editor-save-btn');
const photoEditorUndoButton = document.getElementById('photo-editor-undo-btn');
const photoEditorRedoButton = document.getElementById('photo-editor-redo-btn');
const photoEditorCompareButton = document.getElementById(
  'photo-editor-compare-btn'
);
const photoEditorCanvas = document.getElementById('photo-editor-canvas');
const photoEditorCanvasWrap = document.getElementById('photo-editor-canvas-wrap');
const photoEditorRuleGridButton = document.getElementById(
  'photo-editor-rule-grid-btn'
);
const photoEditorRulerButton = document.getElementById('photo-editor-ruler-btn');
const photoEditorStatus = document.getElementById('photo-editor-status');
const photoEditorPresetResetButton = document.getElementById(
  'photo-editor-preset-reset-btn'
);
const photoEditorPresetList = document.getElementById('photo-editor-presets');
const photoEditorAutoStrengthInput = document.getElementById(
  'photo-editor-auto-strength'
);
const photoEditorAutoStrengthValue = document.getElementById(
  'photo-editor-auto-strength-value'
);
const photoEditorPresetNameInput = document.getElementById(
  'photo-editor-preset-name'
);
const photoEditorSavePresetButton = document.getElementById(
  'photo-editor-save-preset-btn'
);
const photoEditorCropPresetList = document.getElementById(
  'photo-editor-crop-presets'
);
const photoEditorCropResetButton = document.getElementById(
  'photo-editor-crop-reset-btn'
);
const photoEditorCropZoomInput = document.getElementById(
  'photo-editor-crop-zoom'
);
const photoEditorCropZoomValue = document.getElementById(
  'photo-editor-crop-zoom-value'
);
const photoEditorCropXInput = document.getElementById('photo-editor-crop-x');
const photoEditorCropXValue = document.getElementById(
  'photo-editor-crop-x-value'
);
const photoEditorCropYInput = document.getElementById('photo-editor-crop-y');
const photoEditorCropYValue = document.getElementById(
  'photo-editor-crop-y-value'
);
const photoEditorCropTiltInput = document.getElementById(
  'photo-editor-crop-tilt'
);
const photoEditorCropTiltValue = document.getElementById(
  'photo-editor-crop-tilt-value'
);
const photoEditorCropRotateLeftButton = document.getElementById(
  'photo-editor-crop-rotate-left'
);
const photoEditorCropRotateRightButton = document.getElementById(
  'photo-editor-crop-rotate-right'
);
const photoEditorCropFlipXButton = document.getElementById(
  'photo-editor-crop-flip-x'
);
const photoEditorCropFlipYButton = document.getElementById(
  'photo-editor-crop-flip-y'
);
const photoEditorExportResetButton = document.getElementById(
  'photo-editor-export-reset-btn'
);
const photoEditorExportFormatSelect = document.getElementById(
  'photo-editor-export-format'
);
const photoEditorExportMaxEdgeSelect = document.getElementById(
  'photo-editor-export-max-edge'
);
const photoEditorExportQualityInput = document.getElementById(
  'photo-editor-export-quality'
);
const photoEditorExportQualityValue = document.getElementById(
  'photo-editor-export-quality-value'
);
const photoEditorExportQualityRow = document.getElementById(
  'photo-editor-export-quality-row'
);
const photoEditorTextResetButton = document.getElementById(
  'photo-editor-text-reset-btn'
);
const photoEditorTextAddButton = document.getElementById(
  'photo-editor-text-add-btn'
);
const photoEditorTextDeleteButton = document.getElementById(
  'photo-editor-text-delete-btn'
);
const photoEditorTextList = document.getElementById(
  'photo-editor-text-list'
);
const photoEditorTextControls = document.getElementById(
  'photo-editor-text-controls'
);
const photoEditorTextContentInput = document.getElementById(
  'photo-editor-text-content'
);
const photoEditorTextFontSelect = document.getElementById(
  'photo-editor-text-font'
);
const photoEditorTextSizeInput = document.getElementById(
  'photo-editor-text-size'
);
const photoEditorTextSizeValue = document.getElementById(
  'photo-editor-text-size-value'
);
const photoEditorTextColorInput = document.getElementById(
  'photo-editor-text-color'
);
const photoEditorTextWeightSelect = document.getElementById(
  'photo-editor-text-weight'
);
const photoEditorTextStrokeTypeSelect = document.getElementById(
  'photo-editor-text-stroke-type'
);
const photoEditorTextStrokeWidthInput = document.getElementById(
  'photo-editor-text-stroke-width'
);
const photoEditorTextStrokeWidthValue = document.getElementById(
  'photo-editor-text-stroke-width-value'
);
const photoEditorTextStrokeColorInput = document.getElementById(
  'photo-editor-text-stroke-color'
);
const photoEditorTextFillTransparentInput = document.getElementById(
  'photo-editor-text-fill-transparent'
);
const photoEditorTextMaskModeSelect = document.getElementById(
  'photo-editor-text-mask-mode'
);
const photoEditorTextLetterSpacingInput = document.getElementById(
  'photo-editor-text-letter-spacing'
);
const photoEditorTextLetterSpacingValue = document.getElementById(
  'photo-editor-text-letter-spacing-value'
);
const photoEditorImageOverlayResetButton = document.getElementById(
  'photo-editor-image-overlay-reset-btn'
);
const photoEditorImageOverlayAddButton = document.getElementById(
  'photo-editor-image-overlay-add-btn'
);
const photoEditorImageOverlayLibrary = document.getElementById(
  'photo-editor-image-overlay-library'
);
const photoEditorImageOverlayList = document.getElementById(
  'photo-editor-image-overlay-list'
);
const photoEditorImageOverlayControls = document.getElementById(
  'photo-editor-image-overlay-controls'
);
const photoEditorImageOverlayOpacityInput = document.getElementById(
  'photo-editor-image-overlay-opacity'
);
const photoEditorImageOverlayOpacityValue = document.getElementById(
  'photo-editor-image-overlay-opacity-value'
);
const photoEditorImageOverlayBlendModeSelect = document.getElementById(
  'photo-editor-image-overlay-blend-mode'
);
const photoEditorImageOverlayMaskModeSelect = document.getElementById(
  'photo-editor-image-overlay-mask-mode'
);
const photoEditorImageOverlayForwardButton = document.getElementById(
  'photo-editor-image-overlay-forward-btn'
);
const photoEditorImageOverlayBackwardButton = document.getElementById(
  'photo-editor-image-overlay-backward-btn'
);
const photoEditorImageOverlayDeleteButton = document.getElementById(
  'photo-editor-image-overlay-delete-btn'
);
const photoEditorBlurResetButton = document.getElementById(
  'photo-editor-blur-reset-btn'
);
const photoEditorBlurModeGroup = document.getElementById(
  'photo-editor-blur-mode'
);
const photoEditorBlurModeButtons = Array.from(
  document.querySelectorAll('[data-photo-editor-blur-mode]')
);
const photoEditorBlurAmountInput = document.getElementById(
  'photo-editor-blur-amount'
);
const photoEditorBlurAmountValue = document.getElementById(
  'photo-editor-blur-amount-value'
);
const photoEditorBlurConfirmButton = document.getElementById(
  'photo-editor-blur-confirm-btn'
);
const photoEditorBlurConfirmLabel = document.getElementById(
  'photo-editor-blur-confirm-label'
);
const photoEditorAdjustmentTargetSelect = document.getElementById(
  'photo-editor-adjustment-target'
);
const photoEditorSubjectResetButton = document.getElementById(
  'photo-editor-subject-reset-btn'
);
const photoEditorSubjectStatus = document.getElementById(
  'photo-editor-subject-status'
);
const photoEditorSubjectAutoButton = document.getElementById(
  'photo-editor-subject-auto-btn'
);
const photoEditorSubjectTransparentSaveButton = document.getElementById(
  'photo-editor-subject-transparent-save-btn'
);
const photoEditorSubjectStandardDownloadButton = document.getElementById(
  'photo-editor-subject-standard-download-btn'
);
const photoEditorSubjectHighQualityDownloadButton = document.getElementById(
  'photo-editor-subject-high-quality-download-btn'
);
const photoEditorSubjectHighQualityButton = document.getElementById(
  'photo-editor-subject-high-quality-btn'
);
const photoEditorSubjectImportButton = document.getElementById(
  'photo-editor-subject-import-btn'
);
const photoEditorSubjectDeleteButton = document.getElementById(
  'photo-editor-subject-delete-btn'
);
const photoEditorSubjectFileInput = document.getElementById(
  'photo-editor-subject-file'
);
const photoEditorSubjectShowMaskInput = document.getElementById(
  'photo-editor-subject-show-mask'
);
const photoEditorSubjectInvertInput = document.getElementById(
  'photo-editor-subject-invert'
);
const photoEditorMaskToolGroup = document.getElementById(
  'photo-editor-mask-tools'
);
const photoEditorMaskToolButtons = Array.from(
  document.querySelectorAll('[data-photo-editor-mask-tool]')
);
const photoEditorMaskOptions = document.getElementById(
  'photo-editor-mask-options'
);
const photoEditorMaskShapeButtons = Array.from(
  document.querySelectorAll('[data-photo-editor-mask-shape]')
);
const photoEditorMaskShapeGroup = document.getElementById(
  'photo-editor-mask-shapes'
);
const photoEditorMaskBlurStrengthRow = document.getElementById(
  'photo-editor-mask-blur-strength-row'
);
const photoEditorMaskBlurStrengthInput = document.getElementById(
  'photo-editor-mask-blur-strength'
);
const photoEditorMaskBlurStrengthValue = document.getElementById(
  'photo-editor-mask-blur-strength-value'
);
const photoEditorMaskStrengthLabel = document.getElementById(
  'photo-editor-mask-strength-label'
);
const photoEditorFillColorInput = document.getElementById(
  'photo-editor-fill-color'
);
const photoEditorFillColorRow = document.getElementById(
  'photo-editor-fill-color-row'
);
const photoEditorMaskConfirmButton = document.getElementById(
  'photo-editor-mask-confirm-btn'
);
const photoEditorMaskUndoButton = document.getElementById(
  'photo-editor-mask-undo-btn'
);
const photoEditorMaskClearButton = document.getElementById(
  'photo-editor-mask-clear-btn'
);
const photoEditorAdjustmentResetButton = document.getElementById(
  'photo-editor-adjustment-reset-btn'
);
const photoEditorAdjustmentList = document.getElementById(
  'photo-editor-adjustments'
);
const photoEditorCurveResetButton = document.getElementById(
  'photo-editor-curve-reset-btn'
);
const photoEditorCurveModeGroup = document.getElementById(
  'photo-editor-curve-mode'
);
const photoEditorCurveModeButtons = Array.from(
  document.querySelectorAll('[data-photo-editor-curve-mode]')
);
const photoEditorCurveChannelList = document.getElementById(
  'photo-editor-curve-channels'
);
const photoEditorCurveCanvas = document.getElementById(
  'photo-editor-curve-canvas'
);
const photoEditorAccordionToggles = Array.from(
  document.querySelectorAll('[data-photo-editor-accordion-toggle]')
);

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
const settingsDataSection = settingsModal?.querySelector('.settings-data-section');
const settingsDataToggleButton = document.getElementById('settings-data-toggle');
const settingsDataPanel = document.getElementById('settings-data-panel');
const settingsDataStatus = document.getElementById('settings-data-status');
const settingsAiModelList = document.getElementById('settings-ai-model-list');
const settingsAiModelStatus = document.getElementById('settings-ai-model-status');
let trackedFolderSettingsMeta = null;
let settingsMaintenanceStatus = null;
let isSettingsDataSectionOpen = false;
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
const createAppDataBackupButton = document.getElementById(
  'create-app-data-backup-btn'
);
const checkAppDataHealthButton = document.getElementById(
  'check-app-data-health-btn'
);
const showMissingOriginalFilesButton = document.getElementById(
  'show-missing-original-files-btn'
);
const showMissingThumbnailsButton = document.getElementById(
  'show-missing-thumbnails-btn'
);
const showMissingWorldInfoButton = document.getElementById(
  'show-missing-world-info-btn'
);
const showWorldMetadataIssuesButton = document.getElementById(
  'show-world-metadata-issues-btn'
);
const regenerateMissingThumbnailsButton = document.getElementById(
  'regenerate-missing-thumbnails-btn'
);
const refreshWorldMetadataIssuesButton = document.getElementById(
  'refresh-world-metadata-issues-btn'
);
const restoreAppDataBackupButton = document.getElementById(
  'restore-app-data-backup-btn'
);
const exportPhotoCatalogCsvButton = document.getElementById(
  'export-photo-catalog-csv-btn'
);
const exportPhotoCatalogJsonButton = document.getElementById(
  'export-photo-catalog-json-btn'
);
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
let currentPhotoGroupIndexMap = new Map();
let currentSidebarMode = 'timeline';
let currentWorldSidebarSort = 'count';
let currentPhotoSortOrder = 'desc';
let currentPhotoCardDensity = 'default';
let expandedMonthDayKey = '';
let activeSidebarGroupDate = '';
let activeGalleryDateSyncFrame = 0;
let galleryDateJumpTimer = 0;
let galleryDateJumpRenderFrame = 0;
let galleryDateJumpRequestId = 0;
let activeGalleryJumpTarget = null;
let galleryJumpAnimationTimer = 0;
let currentModalPhoto = null;
let photoEditorState = null;
let imageModalAnimationTimer = null;
let imageModalSwitchTimer = null;
let photoEditorRenderFrame = 0;
let photoEditorRenderDebounceTimer = 0;
let photoEditorPreviewSettleTimer = 0;
let photoEditorPreviewRenderToken = 0;
let photoEditorPreviewWorkCanvas = null;
let photoEditorPreviewOverlayCanvas = null;
let photoEditorPreviewOverlayMeta = null;
let photoEditorPreviewCommittedValues = null;
let photoEditorWorker = null;
let photoEditorWorkerRequestId = 0;
let photoEditorWorkerUnavailable = false;
const photoEditorWorkerRequests = new Map();
let photoEditorTextRecentFonts = [];
let photoEditorOverlayAssets = [];
const photoEditorOverlayImageCache = new Map();
const photoEditorSubjectMaskImageCache = new Map();
let photoEditorAiSubjectModelStatus = null;
const photoEditorSubjectModelSessionCache = new Map();
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
let photoEditorUserPresets = [];

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
let monthGalleryAppendFrame = 0;
let renderedGalleryGroupMap = new Map();
let renderedGalleryGroupList = [];
let activeGalleryGroupIndex = -1;
const subModalAnimationTimers = new WeakMap();

const GALLERY_CARD_MIN_WIDTH = 220;
const GALLERY_GRID_GAP = 16;
const GALLERY_GROUP_HORIZONTAL_PADDING = 36;
const GALLERY_INITIAL_PREFETCH_ROWS = 1;
const GALLERY_INCREMENT_ROWS = 2;
const GALLERY_MAX_CARDS_PER_APPEND = 42;
const GALLERY_EAGER_THUMBNAIL_COUNT = 18;
const GALLERY_CARD_EXTRA_HEIGHT = 110;
const GALLERY_LOAD_AHEAD_PX = 560;
const GALLERY_LOAD_CHECK_THROTTLE_MS = 72;
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
const PHOTO_EDITOR_PREVIEW_MAX_EDGE = 3200;
const PHOTO_EDITOR_INTERACTIVE_PREVIEW_MAX_EDGE = 1200;
const PHOTO_EDITOR_CROP_INTERACTIVE_PREVIEW_MAX_EDGE = 520;
const PHOTO_EDITOR_CURVE_DRAG_PREVIEW_MAX_EDGE = 820;
const PHOTO_EDITOR_HEAVY_INTERACTIVE_PREVIEW_MAX_EDGE = 760;
const PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS = 16;
const PHOTO_EDITOR_HEAVY_INTERACTIVE_PREVIEW_DEBOUNCE_MS = 42;
const PHOTO_EDITOR_EFFECT_DRAG_PREVIEW_DEBOUNCE_MS = 58;
const PHOTO_EDITOR_INTERACTIVE_PREVIEW_SETTLE_MS = 140;
const PHOTO_EDITOR_MIN_MASK_SIZE = 0.006;
const PHOTO_EDITOR_HISTORY_LIMIT = 40;
const PHOTO_EDITOR_HISTORY_COMMIT_DELAY_MS = 420;
const PHOTO_EDITOR_MASK_OUTSIDE_MARGIN = 1.25;
const PHOTO_EDITOR_MASK_HANDLE_HIT_RADIUS = 22;
const PHOTO_EDITOR_MASK_ROTATE_HANDLE_OFFSET = 34;
const PHOTO_EDITOR_USER_PRESET_LIMIT = 24;
const PHOTO_EDITOR_USER_PRESETS_STORAGE_KEY =
  'vrchat-world-photo-manager-photo-editor-user-presets';
const PHOTO_EDITOR_TEXT_RECENT_FONTS_STORAGE_KEY =
  'vrchat-world-photo-manager-photo-editor-recent-text-fonts';
const PHOTO_EDITOR_TEXT_RECENT_FONT_LIMIT = 5;
const PHOTO_EDITOR_TEXT_REFERENCE_EDGE = 900;
const PHOTO_EDITOR_TEXT_CENTER_SNAP_THRESHOLD = 0.006;
const PHOTO_EDITOR_IMAGE_OVERLAY_LIMIT = 24;
const PHOTO_EDITOR_IMAGE_OVERLAY_MIN_SIZE = 0.018;
const PHOTO_EDITOR_IMAGE_OVERLAY_DEFAULT_EDGE = 0.28;
const PHOTO_EDITOR_IMAGE_OVERLAY_HANDLE_HIT_RADIUS = 24;
const PHOTO_EDITOR_IMAGE_OVERLAY_BLEND_MODES = Object.freeze([
  'source-over',
  'multiply',
  'screen',
  'overlay',
  'soft-light',
  'hard-light',
  'darken',
  'lighten',
]);
const PHOTO_EDITOR_MASK_MODES = Object.freeze([
  'normal',
  'subject-only',
  'background-only',
]);
const PHOTO_EDITOR_INTERNAL_MASK_MODES = PHOTO_EDITOR_MASK_MODES;
const PHOTO_EDITOR_ADJUSTMENT_TARGETS = Object.freeze([
  'whole',
  'subject',
  'background',
]);
const PHOTO_EDITOR_RULER_GUIDE_LIMIT = 40;
const PHOTO_EDITOR_RULER_GUIDE_MIN_DRAG_PX = 8;
const PHOTO_EDITOR_TEXT_FONT_OPTIONS = Object.freeze([
  {
    key: 'bebasNeue',
    label: 'Bebas Neue',
    family: '"Bebas Neue", "Segoe UI", sans-serif',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'bodoniModa',
    label: 'Bodoni Moda',
    family: '"Bodoni Moda", Georgia, serif',
    weights: ['400', '500', '700', '900'],
    defaultWeight: '700',
  },
  {
    key: 'delaGothicOne',
    label: 'Dela Gothic One',
    family: '"Dela Gothic One", "Yu Gothic UI", "Segoe UI", sans-serif',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'dmSerifDisplay',
    label: 'DM Serif Display',
    family: '"DM Serif Display", Georgia, serif',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'doHyeon',
    label: 'Do Hyeon',
    family: '"Do Hyeon", "Noto Sans KR", "Malgun Gothic", sans-serif',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'dongle',
    label: 'Dongle',
    family: '"Dongle", "Noto Sans KR", "Malgun Gothic", sans-serif',
    weights: ['300', '400', '700'],
    defaultWeight: '700',
  },
  {
    key: 'dotGothic16',
    label: 'DotGothic16',
    family: '"DotGothic16", "Yu Gothic UI", "Segoe UI", monospace',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'fraunces',
    label: 'Fraunces',
    family: '"Fraunces", Georgia, serif',
    weights: ['400', '700', '900'],
    defaultWeight: '700',
  },
  {
    key: 'gaegu',
    label: 'Gaegu',
    family: '"Gaegu", "Noto Sans KR", "Malgun Gothic", cursive',
    weights: ['300', '400', '700'],
    defaultWeight: '700',
  },
  {
    key: 'gowunDodum',
    label: 'Gowun Dodum',
    family: '"Gowun Dodum", "Noto Sans KR", "Malgun Gothic", sans-serif',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'hachiMaruPop',
    label: 'Hachi Maru Pop',
    family: '"Hachi Maru Pop", "Yu Gothic UI", "Segoe UI", cursive',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'ibmPlexSansKr',
    label: 'IBM Plex Sans KR',
    family: '"IBM Plex Sans KR", "Noto Sans KR", "Malgun Gothic", sans-serif',
    weights: ['400', '500', '700'],
    defaultWeight: '700',
  },
  {
    key: 'jua',
    label: 'Jua',
    family: '"Jua", "Noto Sans KR", "Malgun Gothic", sans-serif',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'kaiseiDecol',
    label: 'Kaisei Decol',
    family: '"Kaisei Decol", "Yu Gothic UI", "Segoe UI", serif',
    weights: ['400', '500', '700'],
    defaultWeight: '700',
  },
  {
    key: 'kaiseiOpti',
    label: 'Kaisei Opti',
    family: '"Kaisei Opti", "Yu Gothic UI", "Segoe UI", serif',
    weights: ['400', '500', '700'],
    defaultWeight: '700',
  },
  {
    key: 'kleeOne',
    label: 'Klee One',
    family: '"Klee One", "Yu Gothic UI", "Segoe UI", cursive',
    weights: ['400', '600'],
    defaultWeight: '600',
  },
  {
    key: 'kiwimaru',
    label: 'Kiwi Maru',
    family: '"Kiwi Maru", "Segoe UI", "Yu Gothic UI", sans-serif',
    weights: ['400', '500', '700'],
    defaultWeight: '700',
  },
  {
    key: 'lobster',
    label: 'Lobster',
    family: '"Lobster", "Segoe UI", cursive',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'mplus',
    label: 'M PLUS 1',
    family: '"M PLUS 1", "Segoe UI", "Yu Gothic UI", sans-serif',
    weights: ['400', '500', '700'],
    defaultWeight: '700',
  },
  {
    key: 'monomaniacOne',
    label: 'Monomaniac One',
    family: '"Monomaniac One", "Yu Gothic UI", "Segoe UI", sans-serif',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'notoSansKr',
    label: 'Noto Sans KR',
    family: '"Noto Sans KR", "Malgun Gothic", "Segoe UI", sans-serif',
    weights: ['400', '500', '700'],
    defaultWeight: '700',
  },
  {
    key: 'notoSerifJp',
    label: 'Noto Serif JP',
    family: '"Noto Serif JP", "Yu Mincho", "Yu Gothic UI", serif',
    weights: ['400', '500', '700', '900'],
    defaultWeight: '700',
  },
  {
    key: 'rampartOne',
    label: 'Rampart One',
    family: '"Rampart One", "Yu Gothic UI", "Segoe UI", cursive',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'rocknrollOne',
    label: 'RocknRoll One',
    family: '"RocknRoll One", "Yu Gothic UI", "Segoe UI", sans-serif',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'sacramento',
    label: 'Sacramento',
    family: '"Sacramento", "Segoe UI", cursive',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'sawarabimincho',
    label: 'Sawarabi Mincho',
    family: '"Sawarabi Mincho", "Yu Gothic UI", "Segoe UI", serif',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'shipporiMincho',
    label: 'Shippori Mincho',
    family: '"Shippori Mincho", "Yu Mincho", "Yu Gothic UI", serif',
    weights: ['400', '500', '700'],
    defaultWeight: '700',
  },
  {
    key: 'singleDay',
    label: 'Single Day',
    family: '"Single Day", "Noto Sans KR", "Malgun Gothic", cursive',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'spaceGrotesk',
    label: 'Space Grotesk',
    family: '"Space Grotesk", "Segoe UI", sans-serif',
    weights: ['400', '500', '700'],
    defaultWeight: '700',
  },
  {
    key: 'syne',
    label: 'Syne',
    family: '"Syne", "Segoe UI", sans-serif',
    weights: ['400', '500', '700', '800'],
    defaultWeight: '700',
  },
  {
    key: 'system',
    label: 'システム',
    family: '"Segoe UI", "Yu Gothic UI", "Malgun Gothic", sans-serif',
    weights: ['400', '600', '700'],
    defaultWeight: '700',
  },
  {
    key: 'tsukimiRounded',
    label: 'Tsukimi Rounded',
    family: '"Tsukimi Rounded", "Yu Gothic UI", "Segoe UI", sans-serif',
    weights: ['400', '500', '700'],
    defaultWeight: '700',
  },
  {
    key: 'unbounded',
    label: 'Unbounded',
    family: '"Unbounded", "Segoe UI", sans-serif',
    weights: ['400', '500', '700', '900'],
    defaultWeight: '700',
  },
  {
    key: 'walterTurncoat',
    label: 'Walter Turncoat',
    family: '"Walter Turncoat", "Segoe UI", cursive',
    weights: ['400'],
    defaultWeight: '400',
  },
  {
    key: 'zenmaru',
    label: 'Zen Maru Gothic',
    family: '"Zen Maru Gothic", "Segoe UI", "Yu Gothic UI", sans-serif',
    weights: ['400', '500', '700'],
    defaultWeight: '700',
  },
]);
const PHOTO_EDITOR_TEXT_STROKE_TYPES = Object.freeze([
  'none',
  'outline',
  'shadow',
  'glow',
]);
const PHOTO_EDITOR_TEXT_WEIGHTS = Object.freeze(['400', '500', '700', '900']);
const PHOTO_EDITOR_EXPORT_FORMATS = Object.freeze({
  png: {
    label: 'PNG',
    mimeType: 'image/png',
    extension: 'png',
    supportsQuality: false,
  },
  jpeg: {
    label: 'JPEG',
    mimeType: 'image/jpeg',
    extension: 'jpg',
    supportsQuality: true,
  },
  webp: {
    label: 'WebP',
    mimeType: 'image/webp',
    extension: 'webp',
    supportsQuality: true,
  },
});
const PHOTO_EDITOR_EXPORT_MAX_EDGES = Object.freeze([0, 3840, 2560, 2048, 1600, 1200, 1024, 800]);
const PHOTO_EDITOR_AUTO_ENHANCE_ANALYSIS_MAX_EDGE = 512;
const PHOTO_EDITOR_AUTO_ENHANCE_DEFAULT_STRENGTH = 50;
const PHOTO_EDITOR_CROP_ZOOM_MAX = 300;
const PHOTO_EDIT_SLIDERS = [
  { key: 'brightness', label: '明るさ', min: -100, max: 100, defaultValue: 0 },
  { key: 'exposure', label: '露出', min: -100, max: 100, defaultValue: 0 },
  { key: 'contrast', label: 'コントラスト', min: -60, max: 60, defaultValue: 0 },
  { key: 'highlights', label: 'ハイライト', min: -60, max: 60, defaultValue: 0 },
  { key: 'shadows', label: 'シャドウ', min: -100, max: 100, defaultValue: 0 },
  { key: 'whites', label: 'ホワイト', min: -100, max: 100, defaultValue: 0 },
  { key: 'blacks', label: 'ブラック', min: -100, max: 100, defaultValue: 0 },
  { key: 'gamma', label: 'ガンマ', min: -100, max: 100, defaultValue: 0 },
  { key: 'temperature', label: '色温度', min: -100, max: 100, defaultValue: 0 },
  { key: 'tint', label: '色合い', min: -100, max: 100, defaultValue: 0 },
  { key: 'saturation', label: '彩度', min: -100, max: 100, defaultValue: 0 },
  { key: 'vibrance', label: '自然な彩度', min: -100, max: 100, defaultValue: 0 },
  { key: 'clarity', label: '明瞭度', min: -100, max: 100, defaultValue: 0 },
  { key: 'texture', label: 'テクスチャ', min: -100, max: 100, defaultValue: 0 },
  { key: 'sharpness', label: 'シャープ', min: 0, max: 100, defaultValue: 0 },
  { key: 'denoise', label: 'ノイズ低減', min: 0, max: 100, defaultValue: 0 },
  { key: 'fade', label: 'フェード', min: 0, max: 100, defaultValue: 0 },
  { key: 'grain', label: '粒子', min: 0, max: 100, defaultValue: 0 },
  { key: 'vignette', label: 'ビネット', min: -100, max: 100, defaultValue: 0 },
];
const PHOTO_EDIT_DEFAULT_VALUES = Object.freeze(
  Object.fromEntries(
    PHOTO_EDIT_SLIDERS.map((slider) => [slider.key, slider.defaultValue])
  )
);
const PHOTO_EDITOR_MASK_STRENGTH_LABELS = Object.freeze({
  blur: 'ぼかしの濃さ',
  mosaic: 'モザイクの濃さ',
  fill: '塗りつぶしの濃さ',
});
const PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS = Object.freeze({
  blur: 45,
  mosaic: 60,
  fill: 100,
});
const PHOTO_EDITOR_BLUR_MODES = Object.freeze(['radial', 'full', 'background']);
const PHOTO_EDITOR_RADIAL_BLUR_MIN_RADIUS = 0.08;
const PHOTO_EDITOR_RADIAL_BLUR_MAX_RADIUS = 1.12;
const PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER = 0.04;
const PHOTO_EDITOR_CURVE_MODES = Object.freeze(['rgb', 'hsv']);
const PHOTO_EDITOR_CURVE_CHANNELS = Object.freeze({
  rgb: [
    { key: 'master', label: '全体' },
    { key: 'r', label: 'R' },
    { key: 'g', label: 'G' },
    { key: 'b', label: 'B' },
  ],
  hsv: [
    { key: 'master', label: '全体' },
    { key: 'h', label: 'H' },
    { key: 's', label: 'S' },
    { key: 'v', label: 'V' },
  ],
});
const PHOTO_EDITOR_CURVE_DEFAULT_POINTS = Object.freeze([0, 0.25, 0.5, 0.75, 1]);
const PHOTO_EDITOR_CURVE_HISTOGRAM_BINS = 64;
const PHOTO_EDITOR_CURVE_PREVIEW_DEBOUNCE_MS = 96;
const PHOTO_EDIT_PRESETS = {
  auto: {
    label: '✨ 自動補正',
    isAuto: true,
    values: {
      brightness: 0,
      exposure: 0,
      contrast: 0,
      highlights: 0,
      shadows: 0,
      whites: 0,
      blacks: 0,
      temperature: 0,
      saturation: 0,
      vibrance: 0,
      clarity: 0,
      texture: 0,
      sharpness: 0,
      denoise: 0,
      fade: 0,
      grain: 0,
      vignette: 0,
    },
  },
  learningAuto: {
    label: '学習補正',
    isLearningAuto: true,
    values: {
      brightness: 0,
      exposure: 0,
      contrast: 0,
      highlights: 0,
      shadows: 0,
      whites: 0,
      blacks: 0,
      temperature: 0,
      saturation: 0,
      vibrance: 0,
      clarity: 0,
      texture: 0,
      sharpness: 0,
      denoise: 0,
      fade: 0,
      grain: 0,
      vignette: 0,
    },
  },
  vrchatPost: {
    label: '投稿クリア',
    values: {
      brightness: 5,
      exposure: 4,
      contrast: 12,
      highlights: -12,
      shadows: 10,
      whites: 12,
      blacks: -6,
      temperature: 2,
      saturation: 10,
      vibrance: 12,
      clarity: 8,
      texture: 6,
      sharpness: 10,
      denoise: 4,
      fade: 0,
      grain: 0,
      vignette: 8,
    },
  },
  naturalClear: {
    label: '自然クリア',
    values: {
      brightness: 4,
      exposure: 2,
      contrast: 8,
      highlights: -8,
      shadows: 8,
      whites: 8,
      blacks: -4,
      temperature: 1,
      saturation: 6,
      vibrance: 10,
      clarity: 6,
      texture: 4,
      sharpness: 8,
      denoise: 5,
      fade: 0,
      grain: 0,
      vignette: 6,
    },
  },
  night: {
    label: '夜景強調',
    values: {
      brightness: 8,
      exposure: 5,
      contrast: 18,
      highlights: -24,
      shadows: 26,
      whites: -8,
      blacks: -12,
      temperature: -4,
      saturation: 8,
      vibrance: 10,
      clarity: 12,
      texture: 8,
      sharpness: 8,
      denoise: 6,
      fade: 0,
      grain: 5,
      vignette: 16,
    },
  },
  neon: {
    label: 'ネオン強調',
    values: {
      brightness: 3,
      exposure: 2,
      contrast: 18,
      highlights: -10,
      shadows: 8,
      whites: 8,
      blacks: -10,
      temperature: -6,
      saturation: 22,
      vibrance: 18,
      clarity: 14,
      texture: 10,
      sharpness: 10,
      denoise: 3,
      fade: 0,
      grain: 3,
      vignette: 12,
    },
  },
  soft: {
    label: 'ふんわり1',
    values: {
      brightness: 7,
      exposure: 3,
      contrast: -8,
      highlights: -12,
      shadows: 16,
      whites: 6,
      blacks: 4,
      temperature: 7,
      saturation: 4,
      vibrance: 8,
      clarity: -12,
      texture: -10,
      sharpness: 0,
      denoise: 8,
      fade: 7,
      grain: 0,
      vignette: 4,
    },
  },
  soft2: {
    label: 'ふんわり2',
    values: {
      brightness: 9,
      exposure: 4,
      contrast: -12,
      highlights: -18,
      shadows: 20,
      whites: 8,
      blacks: 8,
      temperature: 10,
      saturation: 2,
      vibrance: 12,
      clarity: -18,
      texture: -16,
      sharpness: 0,
      denoise: 12,
      fade: 10,
      grain: 0,
      vignette: 0,
    },
  },
  soft3: {
    label: 'ふんわり3',
    values: {
      brightness: 5,
      exposure: 2,
      contrast: -4,
      highlights: -8,
      shadows: 12,
      whites: 12,
      blacks: 2,
      temperature: -2,
      saturation: 8,
      vibrance: 16,
      clarity: -10,
      texture: -8,
      sharpness: 2,
      denoise: 8,
      fade: 5,
      grain: 3,
      vignette: 6,
    },
  },
  film: {
    label: 'フィルム風',
    values: {
      brightness: 2,
      exposure: 0,
      contrast: 8,
      highlights: -14,
      shadows: 10,
      whites: -6,
      blacks: -6,
      temperature: 8,
      saturation: -6,
      vibrance: 2,
      clarity: 2,
      texture: -2,
      sharpness: 3,
      denoise: 0,
      fade: 12,
      grain: 18,
      vignette: 14,
    },
  },
  highContrast: {
    label: '高コントラスト1',
    values: {
      brightness: 0,
      exposure: 1,
      contrast: 30,
      highlights: -12,
      shadows: -6,
      whites: 16,
      blacks: -18,
      temperature: 0,
      saturation: 10,
      vibrance: 8,
      clarity: 18,
      texture: 12,
      sharpness: 14,
      denoise: 0,
      fade: 0,
      grain: 0,
      vignette: 10,
    },
  },
  highContrast2: {
    label: '高コントラスト2',
    values: {
      brightness: -2,
      exposure: 0,
      contrast: 42,
      highlights: -20,
      shadows: -10,
      whites: 8,
      blacks: -30,
      gamma: -4,
      temperature: -2,
      tint: 0,
      saturation: 6,
      vibrance: 12,
      clarity: 24,
      texture: 16,
      sharpness: 16,
      denoise: 0,
      fade: 0,
      grain: 0,
      vignette: 16,
    },
  },
  coolBlue: {
    label: 'クールブルー1',
    values: {
      brightness: 1,
      exposure: 0,
      contrast: 12,
      highlights: -18,
      shadows: 8,
      whites: -4,
      blacks: -8,
      gamma: -2,
      temperature: -22,
      tint: -4,
      saturation: 2,
      vibrance: 12,
      clarity: 8,
      texture: 6,
      sharpness: 8,
      denoise: 3,
      fade: 0,
      grain: 0,
      vignette: 8,
    },
  },
  deepBlue: {
    label: 'クールブルー2',
    values: {
      brightness: -1,
      exposure: -2,
      contrast: 18,
      highlights: -24,
      shadows: 12,
      whites: -8,
      blacks: -14,
      gamma: -6,
      temperature: -34,
      tint: 6,
      saturation: -2,
      vibrance: 18,
      clarity: 12,
      texture: 8,
      sharpness: 10,
      denoise: 4,
      fade: 2,
      grain: 4,
      vignette: 14,
    },
  },
  sweetPink: {
    label: 'スイートピンク1',
    values: {
      brightness: 6,
      exposure: 3,
      contrast: -4,
      highlights: -12,
      shadows: 14,
      whites: 10,
      blacks: 4,
      gamma: 4,
      temperature: 8,
      tint: 22,
      saturation: 8,
      vibrance: 16,
      clarity: -8,
      texture: -10,
      sharpness: 2,
      denoise: 8,
      fade: 8,
      grain: 0,
      vignette: 4,
    },
  },
  pastelPink: {
    label: 'スイートピンク2',
    values: {
      brightness: 9,
      exposure: 4,
      contrast: -12,
      highlights: -18,
      shadows: 20,
      whites: 12,
      blacks: 8,
      gamma: 6,
      temperature: 4,
      tint: 30,
      saturation: -2,
      vibrance: 20,
      clarity: -16,
      texture: -18,
      sharpness: 0,
      denoise: 12,
      fade: 14,
      grain: 2,
      vignette: 0,
    },
  },
  vividPink: {
    label: 'スイートピンク3',
    values: {
      brightness: 3,
      exposure: 1,
      contrast: 14,
      highlights: -10,
      shadows: 8,
      whites: 14,
      blacks: -10,
      gamma: 0,
      temperature: 2,
      tint: 28,
      saturation: 14,
      vibrance: 24,
      clarity: 8,
      texture: 4,
      sharpness: 8,
      denoise: 4,
      fade: 2,
      grain: 4,
      vignette: 10,
    },
  },
  shadowLift: {
    label: '暗部クリア',
    values: {
      brightness: 6,
      exposure: 5,
      contrast: 10,
      highlights: -18,
      shadows: 30,
      whites: 4,
      blacks: -10,
      temperature: 0,
      saturation: 6,
      vibrance: 10,
      clarity: 12,
      texture: 8,
      sharpness: 10,
      denoise: 10,
      fade: 0,
      grain: 0,
      vignette: 4,
    },
  },
  thumbnailPop: {
    label: 'サムネ強調',
    values: {
      brightness: 3,
      exposure: 2,
      contrast: 22,
      highlights: -12,
      shadows: 6,
      whites: 18,
      blacks: -18,
      temperature: 1,
      saturation: 14,
      vibrance: 14,
      clarity: 16,
      texture: 12,
      sharpness: 16,
      denoise: 2,
      fade: 0,
      grain: 0,
      vignette: 8,
    },
  },
  monochrome: {
    label: 'モノクロ',
    values: {
      brightness: 4,
      exposure: 2,
      contrast: 20,
      highlights: -10,
      shadows: 8,
      whites: 10,
      blacks: -12,
      temperature: 0,
      saturation: -100,
      vibrance: 0,
      clarity: 10,
      texture: 8,
      sharpness: 8,
      denoise: 0,
      fade: 4,
      grain: 12,
      vignette: 12,
    },
  },
};
const PHOTO_EDITOR_CROP_PRESETS = [
  { key: 'original', label: 'オリジナル', ratio: null },
  { key: 'square', label: '1:1', ratio: 1 },
  { key: 'wide', label: '16:9', ratio: 16 / 9 },
  { key: 'portrait', label: '9:16', ratio: 9 / 16 },
  { key: 'fiveFour', label: '5:4', ratio: 5 / 4 },
  { key: 'booth', label: '4:5', ratio: 4 / 5 },
  { key: 'classic', label: '3:2', ratio: 3 / 2 },
  { key: 'twoThree', label: '2:3', ratio: 2 / 3 },
  {
    key: 'avatarThumbnail',
    label: 'アバターサムネイル',
    ratio: 4 / 3,
    exportWidth: 800,
    exportHeight: 600,
    exportMaxEdge: 800,
  },
  {
    key: 'vrcGallery',
    label: 'VRCギャラリー',
    ratio: 1,
    transparentPadding: true,
    exportWidth: 2048,
    exportHeight: 2048,
    exportMaxEdge: 2048,
    exportFormat: 'png',
  },
  {
    key: 'vrcSticker',
    label: '絵文字・ステッカー',
    ratio: 1,
    transparentPadding: true,
    exportWidth: 1024,
    exportHeight: 1024,
    exportMaxEdge: 1024,
    exportFormat: 'png',
  },
];

function pad2(value) {
  return String(value).padStart(2, '0');
}

function getMonthDayKey(year, month) {
  return `${year}-${month}`;
}

function compareSidebarNumberDescending(leftValue, rightValue) {
  return Number(rightValue) - Number(leftValue);
}

function compareSidebarDayNumber(leftDay, rightDay) {
  const diff = Number(leftDay) - Number(rightDay);
  return currentPhotoSortOrder === 'asc' ? diff : -diff;
}

function getOrderedTimelineSidebarData() {
  return sidebarData
    .map((yearEntry) => ({
      ...yearEntry,
      months: [...(yearEntry.months || [])]
        .map((monthEntry) => ({
          ...monthEntry,
          days: [...(monthEntry.days || [])].sort((leftDay, rightDay) =>
            compareSidebarDayNumber(leftDay.day, rightDay.day)
          ),
        }))
        .sort((leftMonth, rightMonth) =>
          compareSidebarNumberDescending(leftMonth.month, rightMonth.month)
        ),
    }))
    .sort((leftYear, rightYear) =>
      compareSidebarNumberDescending(leftYear.year, rightYear.year)
    );
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

const HEALTH_ISSUE_VIEW_META = Object.freeze({
  'missing-original': {
    label: '元画像なし',
    busyStatus: '元画像なしの写真を抽出中...',
    progressMessage: '元画像ファイルが見つからない写真を集めています...',
    successPrefix: '元画像なし',
    emptyToast: '元画像なしの写真はありません',
    successToast: (count) => `元画像なしの写真を${count}件表示しました`,
    errorPrefix: '元画像なし画像の抽出',
  },
  'missing-thumbnail': {
    label: 'サムネイルなし',
    busyStatus: 'サムネイルなしの写真を抽出中...',
    progressMessage: 'サムネイルが欠損している写真を集めています...',
    successPrefix: 'サムネイルなし',
    emptyToast: 'サムネイルなしの写真はありません',
    successToast: (count) => `サムネイルなしの写真を${count}件表示しました`,
    errorPrefix: 'サムネイルなし画像の抽出',
  },
  'missing-world-info': {
    label: 'World情報未取得',
    busyStatus: 'World情報未取得の写真を抽出中...',
    progressMessage: 'World情報が未取得の写真を集めています...',
    successPrefix: 'World情報未取得',
    emptyToast: 'World情報未取得の写真はありません',
    successToast: (count) =>
      `World情報未取得の写真を${count}件表示しました`,
    errorPrefix: 'World情報未取得画像の抽出',
  },
  'world-metadata': {
    label: 'Worldメタデータ要確認',
    busyStatus: 'World要確認画像を抽出中...',
    progressMessage: 'Worldメタデータ要確認の写真を集めています...',
    successPrefix: 'Worldメタデータ要確認',
    emptyToast: 'Worldメタデータ要確認の写真はありません',
    successToast: (count) =>
      `Worldメタデータ要確認の写真を${count}件表示しました`,
    errorPrefix: 'World要確認画像の抽出',
  },
});

function createHealthSelection(kind, label) {
  return {
    mode: 'health',
    kind,
    label,
  };
}

function normalizeSelection(selection) {
  if (selection?.mode === 'health') {
    const rawKind =
      typeof selection.kind === 'string' && selection.kind.trim().length > 0
        ? selection.kind.trim()
        : null;
    const normalizedKind =
      rawKind === 'world-metadata-issues' ? 'world-metadata' : rawKind;
    const normalizedLabel =
      typeof selection.label === 'string' && selection.label.trim().length > 0
        ? selection.label.trim()
        : '状態チェック結果';

    if (!normalizedKind) {
      return null;
    }

    return createHealthSelection(normalizedKind, normalizedLabel);
  }

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

function isHealthSelection(selection = currentSelection) {
  return normalizeSelection(selection)?.mode === 'health';
}

function isSameSelection(leftSelection, rightSelection) {
  const left = normalizeSelection(leftSelection);
  const right = normalizeSelection(rightSelection);

  if (!left || !right) {
    return false;
  }

  if (left.mode === 'health' || right.mode === 'health') {
    return left.mode === right.mode && left.kind === right.kind;
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

  if (normalizedSelection.mode === 'health') {
    return normalizedSelection.label;
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

  if (normalizedSelection.mode === 'health') {
    return '該当する写真はありません';
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

function translateUiText(text) {
  return window.WorldShotI18n?.t
    ? window.WorldShotI18n.t(text)
    : String(text);
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

function getPhotoGroupDate(photo) {
  const groupDate =
    typeof photo?.groupDate === 'string' ? photo.groupDate.trim() : '';

  return groupDate || '日付不明';
}

function rebuildCurrentPhotoGroupIndexMap() {
  const nextGroupIndexMap = new Map();

  currentPhotos.forEach((photo, index) => {
    const groupDate = getPhotoGroupDate(photo);
    const groupIndex = nextGroupIndexMap.get(groupDate);

    if (groupIndex) {
      groupIndex.endIndex = index + 1;
      return;
    }

    nextGroupIndexMap.set(groupDate, {
      startIndex: index,
      endIndex: index + 1,
    });
  });

  currentPhotoGroupIndexMap = nextGroupIndexMap;
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
    const optionName = button.dataset.fontOption || 'standard';
    const isAvailable = isAppFontAvailableForCurrentLanguage(optionName);
    const isActive = isAvailable && optionName === fontName;
    button.hidden = !isAvailable;
    button.disabled = !isAvailable;
    button.setAttribute('aria-hidden', isAvailable ? 'false' : 'true');
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  });
}

const KOREAN_ONLY_APP_FONT_OPTIONS = new Set([
  'notoSansKr',
  'ibmPlexSansKr',
  'dongle',
  'gaegu',
  'gowunDodum',
]);

function getCurrentLanguagePreference() {
  return window.WorldShotI18n?.getLanguage
    ? window.WorldShotI18n.getLanguage()
    : document.body.dataset.language || 'ja';
}

function isAppFontAvailableForCurrentLanguage(fontName) {
  return (
    !KOREAN_ONLY_APP_FONT_OPTIONS.has(fontName) ||
    getCurrentLanguagePreference() === 'ko'
  );
}

function applyFontPreference(fontName) {
  const allowedFonts = new Set([
    'standard',
    'zenmaru',
    'mplus',
    'kiwimaru',
    'sawarabimincho',
    'notoSansKr',
    'ibmPlexSansKr',
    'dongle',
    'gaegu',
    'gowunDodum',
  ]);
  const nextFont =
    typeof fontName === 'string' &&
    allowedFonts.has(fontName) &&
    isAppFontAvailableForCurrentLanguage(fontName)
      ? fontName
      : 'standard';

  document.body.setAttribute('data-font', nextFont);
  localStorage.setItem(FONT_STORAGE_KEY, nextFont);
  syncFontOptionButtons(nextFont);
}

function syncFontPreferenceForLanguage() {
  const currentFont = document.body.getAttribute('data-font') || 'standard';
  if (!isAppFontAvailableForCurrentLanguage(currentFont)) {
    applyFontPreference('standard');
    return;
  }

  syncFontOptionButtons(currentFont);
}

function initializeFontPreference() {
  const savedFont = localStorage.getItem(FONT_STORAGE_KEY);

  if (
    savedFont === 'standard' ||
    savedFont === 'zenmaru' ||
    savedFont === 'mplus' ||
    savedFont === 'kiwimaru' ||
    savedFont === 'sawarabimincho' ||
    savedFont === 'notoSansKr' ||
    savedFont === 'ibmPlexSansKr' ||
    savedFont === 'dongle' ||
    savedFont === 'gaegu' ||
    savedFont === 'gowunDodum'
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
  rebuildCurrentPhotoGroupIndexMap();
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

function setKeyboardFocusToFirstPhotoInGroup(groupSection) {
  const firstCard = groupSection?.querySelector?.('.photo-card:not([hidden])');
  const firstPhotoId = Number(firstCard?.dataset?.photoId);

  if (!Number.isInteger(firstPhotoId) || firstPhotoId <= 0) {
    return false;
  }

  return setKeyboardFocusedPhoto(firstPhotoId, { scroll: false });
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

function removeRenderedGalleryGroupState(groupDate, groupState) {
  renderedGalleryGroupMap.delete(groupDate);

  const groupIndex = renderedGalleryGroupList.indexOf(groupState);

  if (groupIndex === -1) {
    return;
  }

  renderedGalleryGroupList.splice(groupIndex, 1);

  if (activeGalleryGroupIndex === groupIndex) {
    activeGalleryGroupIndex = -1;
  } else if (activeGalleryGroupIndex > groupIndex) {
    activeGalleryGroupIndex -= 1;
  }
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
          removeRenderedGalleryGroupState(groupDate, groupState);
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
  scheduleActiveGalleryDateSync();

  return true;
}

function isSidebarDayButtonCurrent(button) {
  return (
    Boolean(activeSidebarGroupDate) &&
    isMonthSelection(currentSelection) &&
    button?.dataset?.groupDate === activeSidebarGroupDate
  );
}

function syncSidebarActiveDayButtons() {
  if (!sidebarTree) {
    return;
  }

  for (const dayButton of sidebarTree.querySelectorAll('.day-button')) {
    const isActive = isSidebarDayButtonCurrent(dayButton);
    dayButton.classList.toggle('active', isActive);

    if (isActive) {
      dayButton.setAttribute('aria-current', 'date');
    } else {
      dayButton.removeAttribute('aria-current');
    }
  }
}

function setActiveSidebarGroupDate(groupDate) {
  const normalizedGroupDate =
    typeof groupDate === 'string' ? groupDate.trim() : '';
  const nextActiveGalleryGroupIndex = normalizedGroupDate
    ? renderedGalleryGroupList.findIndex(
        (groupState) => groupState?.groupDate === normalizedGroupDate
      )
    : -1;

  if (activeSidebarGroupDate === normalizedGroupDate) {
    activeGalleryGroupIndex = nextActiveGalleryGroupIndex;
    syncSidebarActiveDayButtons();
    return;
  }

  activeSidebarGroupDate = normalizedGroupDate;
  activeGalleryGroupIndex = nextActiveGalleryGroupIndex;
  syncSidebarActiveDayButtons();
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

  syncSidebarActiveDayButtons();

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

  if (currentSelection.mode === 'health') {
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

  if (modalEditPhotoButton) {
    modalEditPhotoButton.disabled = !item.filePath;
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

  if (
    !normalized ||
    normalized === '{}' ||
    normalized === '[]' ||
    /^\[object\s+[^\]]+\]$/i.test(normalized)
  ) {
    return '';
  }

  try {
    const parsed = JSON.parse(normalized);

    if (Array.isArray(parsed) && parsed.length === 0) {
      return '';
    }

    if (
      parsed &&
      typeof parsed === 'object' &&
      Object.values(parsed).every((entry) => {
        if (entry == null) {
          return true;
        }

        if (typeof entry === 'string') {
          return entry.trim().length === 0;
        }

        if (Array.isArray(entry)) {
          return entry.length === 0;
        }

        return false;
      })
    ) {
      return '';
    }
  } catch {
    // Plain text notes are expected here.
  }

  return normalized;
}

function renderModalPrintNote(item = currentModalPhoto) {
  const printNoteText = getNormalizedModalPrintNoteText(item);
  const hasPrintNote = printNoteText.length > 0;
  const hasPrintBadge = Boolean(item?.isPrintPhoto || hasPrintNote);

  if (modalPrintNoteValue) {
    modalPrintNoteValue.textContent = hasPrintNote ? printNoteText : '';
  }

  if (modalPrintNoteBlock) {
    modalPrintNoteBlock.classList.toggle('is-hidden', !hasPrintNote);
  }

  if (modalPrintNoteHeroBadge) {
    modalPrintNoteHeroBadge.textContent = 'プリント';
    modalPrintNoteHeroBadge.classList.toggle('is-hidden', !hasPrintBadge);
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

function clonePhotoEditValues(values = {}) {
  return {
    ...PHOTO_EDIT_DEFAULT_VALUES,
    ...Object.fromEntries(
      PHOTO_EDIT_SLIDERS.map((slider) => [
        slider.key,
        clampNumber(values[slider.key], slider.min, slider.max, slider.defaultValue),
      ])
    ),
  };
}

function normalizePhotoEditorPresetName(name) {
  return String(name || '').replace(/\s+/g, ' ').trim().slice(0, 32);
}

function createPhotoEditorUserPresetId() {
  return `user-${Date.now().toString(36)}-${Math.random()
    .toString(36)
    .slice(2, 8)}`;
}

function normalizePhotoEditorUserPreset(rawPreset) {
  const label = normalizePhotoEditorPresetName(rawPreset?.label);

  if (!label) {
    return null;
  }

  const id =
    typeof rawPreset?.id === 'string' && rawPreset.id.trim()
      ? rawPreset.id.trim()
      : createPhotoEditorUserPresetId();

  return {
    id,
    label,
    values: clonePhotoEditValues(rawPreset?.values),
    crop: rawPreset?.crop ? normalizePhotoEditorCropState(rawPreset.crop) : null,
    blur: rawPreset?.blur ? normalizePhotoEditorBlurState(rawPreset.blur) : null,
    curve: rawPreset?.curve ? normalizePhotoEditorCurveState(rawPreset.curve) : null,
    exportSettings: rawPreset?.exportSettings
      ? normalizePhotoEditorExportSettings(rawPreset.exportSettings)
      : null,
  };
}

function getPhotoEditorUserPresetOrder(preset, fallbackOrder) {
  const match = /^user-([a-z0-9]+)-/i.exec(String(preset?.id || ''));

  if (!match) {
    return fallbackOrder;
  }

  const order = Number.parseInt(match[1], 36);
  return Number.isFinite(order) ? order : fallbackOrder;
}

function loadPhotoEditorUserPresets() {
  try {
    const parsed = JSON.parse(
      localStorage.getItem(PHOTO_EDITOR_USER_PRESETS_STORAGE_KEY) || '[]'
    );

    return (Array.isArray(parsed) ? parsed : [])
      .map(normalizePhotoEditorUserPreset)
      .filter(Boolean)
      .map((preset, index) => ({
        preset,
        order: getPhotoEditorUserPresetOrder(preset, index),
        index,
      }))
      .sort((left, right) => left.order - right.order || left.index - right.index)
      .map((entry) => entry.preset)
      .slice(-PHOTO_EDITOR_USER_PRESET_LIMIT);
  } catch {
    return [];
  }
}

function savePhotoEditorUserPresetsToStorage() {
  try {
    localStorage.setItem(
      PHOTO_EDITOR_USER_PRESETS_STORAGE_KEY,
      JSON.stringify(photoEditorUserPresets)
    );
  } catch {
    showToast('プリセットを保存できませんでした');
  }
}

function getPhotoEditorPreset(presetKey) {
  if (PHOTO_EDIT_PRESETS[presetKey]) {
    return PHOTO_EDIT_PRESETS[presetKey];
  }

  if (typeof presetKey === 'string' && presetKey.startsWith('user:')) {
    const userPresetId = presetKey.slice('user:'.length);
    return (
      photoEditorUserPresets.find((preset) => preset.id === userPresetId) ||
      null
    );
  }

  return null;
}

function createPhotoEditorPresetButton({ key, label, isUserPreset = false }) {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = isUserPreset
    ? 'photo-editor-preset-button is-user-preset'
    : 'photo-editor-preset-button';
  button.dataset.photoEditorPreset = key;
  button.textContent = label;
  button.addEventListener('click', () => {
    applyPhotoEditorPreset(key);
  });
  return button;
}

function deletePhotoEditorUserPreset(presetId) {
  const presetIndex = photoEditorUserPresets.findIndex(
    (preset) => preset.id === presetId
  );

  if (presetIndex < 0) {
    return;
  }

  const [deletedPreset] = photoEditorUserPresets.splice(presetIndex, 1);
  savePhotoEditorUserPresetsToStorage();
  renderPhotoEditorPresetButtons();
  setPhotoEditorStatus(`プリセットを削除しました: ${deletedPreset.label}`);
  showToast(`プリセットを削除しました: ${deletedPreset.label}`);
}

function createPhotoEditorUserPresetItem(preset) {
  const item = document.createElement('div');
  item.className = 'photo-editor-user-preset-item';

  item.appendChild(
    createPhotoEditorPresetButton({
      key: `user:${preset.id}`,
      label: preset.label,
      isUserPreset: true,
    })
  );

  const deleteButton = document.createElement('button');
  deleteButton.type = 'button';
  deleteButton.className = 'photo-editor-user-preset-delete';
  deleteButton.setAttribute('aria-label', `${preset.label}を削除`);
  deleteButton.title = `${preset.label}を削除`;
  deleteButton.innerHTML = '<span class="material-symbols-outlined">delete</span>';
  deleteButton.addEventListener('click', (event) => {
    event.stopPropagation();
    deletePhotoEditorUserPreset(preset.id);
  });
  item.appendChild(deleteButton);

  return item;
}

function renderPhotoEditorPresetButtons() {
  if (!photoEditorPresetList) {
    return;
  }

  photoEditorPresetList.innerHTML = '';

  for (const [presetKey, preset] of Object.entries(PHOTO_EDIT_PRESETS)) {
    photoEditorPresetList.appendChild(
      createPhotoEditorPresetButton({
        key: presetKey,
        label: preset.label,
      })
    );
  }

  for (const preset of photoEditorUserPresets) {
    photoEditorPresetList.appendChild(createPhotoEditorUserPresetItem(preset));
  }
}

function createPhotoEditorPresetStateFromEditor() {
  if (!photoEditorState) {
    return {
      values: clonePhotoEditValues(),
      crop: getDefaultPhotoEditorCropState(),
      blur: getDefaultPhotoEditorBlurState(),
      curve: getDefaultPhotoEditorCurveState(),
      exportSettings: getDefaultPhotoEditorExportSettings(),
    };
  }

  return {
    values: clonePhotoEditValues(photoEditorState.values),
    crop: normalizePhotoEditorCropState(photoEditorState.crop),
    blur: normalizePhotoEditorBlurState(photoEditorState.blur),
    curve: normalizePhotoEditorCurveState(photoEditorState.curve),
    exportSettings: normalizePhotoEditorExportSettings(photoEditorState.exportSettings),
  };
}

function saveCurrentPhotoEditorPreset() {
  if (!photoEditorState) {
    return;
  }

  const label = normalizePhotoEditorPresetName(photoEditorPresetNameInput?.value);

  if (!label) {
    setPhotoEditorStatus('プリセット名を入力してください');
    return;
  }

  const nextPreset = {
    id: createPhotoEditorUserPresetId(),
    label,
    ...createPhotoEditorPresetStateFromEditor(),
  };
  const existingIndex = photoEditorUserPresets.findIndex(
    (preset) => preset.label.toLowerCase() === label.toLowerCase()
  );

  if (existingIndex >= 0) {
    nextPreset.id = photoEditorUserPresets[existingIndex].id;
    photoEditorUserPresets.splice(existingIndex, 1, nextPreset);
  } else {
    photoEditorUserPresets.push(nextPreset);
  }

  photoEditorUserPresets = photoEditorUserPresets.slice(
    -PHOTO_EDITOR_USER_PRESET_LIMIT
  );
  savePhotoEditorUserPresetsToStorage();
  renderPhotoEditorPresetButtons();

  if (photoEditorPresetNameInput) {
    photoEditorPresetNameInput.value = '';
  }

  setPhotoEditorStatus(`プリセットを保存しました: ${label}`);
  showToast(`プリセットを保存しました: ${label}`);
}

function normalizePhotoEditorCropRotation(value) {
  const numericValue = Number(value);
  const roundedValue = Number.isFinite(numericValue)
    ? Math.round(numericValue / 90) * 90
    : 0;

  return ((roundedValue % 360) + 360) % 360;
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

  return clampNumber(nextZoom, 0, PHOTO_EDITOR_CROP_ZOOM_MAX, 0);
}

function getPhotoEditorCropZoomScale(crop = {}) {
  return 1 + normalizePhotoEditorCropZoomValue(crop) / 100;
}

function getDefaultPhotoEditorCropState() {
  return {
    preset: 'original',
    zoomMode: 'offset',
    zoom: 0,
    offsetX: 0,
    offsetY: 0,
    rotation: 0,
    flipX: false,
    flipY: false,
    tilt: 0,
  };
}

function normalizePhotoEditorCropState(crop = {}) {
  const preset = getPhotoEditorCropPreset(crop?.preset);

  return {
    preset: preset.key,
    zoomMode: 'offset',
    zoom: normalizePhotoEditorCropZoomValue(crop),
    offsetX: clampNumber(crop?.offsetX, -100, 100, 0),
    offsetY: clampNumber(crop?.offsetY, -100, 100, 0),
    rotation: normalizePhotoEditorCropRotation(crop?.rotation),
    flipX: Boolean(crop?.flipX),
    flipY: Boolean(crop?.flipY),
    tilt: clampNumber(crop?.tilt, -45, 45, 0),
  };
}

function getDefaultPhotoEditorExportSettings() {
  return {
    format: 'png',
    quality: 92,
    maxEdge: 0,
  };
}

function normalizePhotoEditorExportSettings(settings = {}) {
  const defaults = getDefaultPhotoEditorExportSettings();
  const format = PHOTO_EDITOR_EXPORT_FORMATS[settings?.format]
    ? settings.format
    : defaults.format;
  const maxEdge = clampNumber(
    settings?.maxEdge,
    0,
    Math.max(...PHOTO_EDITOR_EXPORT_MAX_EDGES),
    defaults.maxEdge
  );
  const normalizedMaxEdge = PHOTO_EDITOR_EXPORT_MAX_EDGES.includes(maxEdge)
    ? maxEdge
    : defaults.maxEdge;

  return {
    format,
    quality: clampNumber(settings?.quality, 60, 100, defaults.quality),
    maxEdge: normalizedMaxEdge,
  };
}

function getPhotoEditorTextFontOption(fontKey) {
  return (
    PHOTO_EDITOR_TEXT_FONT_OPTIONS.find((font) => font.key === fontKey) ||
    PHOTO_EDITOR_TEXT_FONT_OPTIONS[0]
  );
}

function getPhotoEditorTextFontWeights(fontKey) {
  const font = getPhotoEditorTextFontOption(fontKey);
  const weights = Array.isArray(font.weights) && font.weights.length > 0
    ? font.weights
    : PHOTO_EDITOR_TEXT_WEIGHTS;

  return weights.map((weight) => String(weight));
}

function createPhotoEditorTextFontOption(font) {
  const normalizedFont = getPhotoEditorTextFontOption(font?.key);
  const option = document.createElement('option');
  option.value = normalizedFont.key;
  option.textContent = normalizedFont.label;
  option.style.fontFamily = normalizedFont.family;
  option.style.fontWeight = normalizedFont.defaultWeight || '700';
  return option;
}

function syncPhotoEditorTextFontSelectPreview(fontKey) {
  if (!photoEditorTextFontSelect) {
    return;
  }

  const font = getPhotoEditorTextFontOption(fontKey);
  photoEditorTextFontSelect.style.fontFamily = font.family;
  photoEditorTextFontSelect.style.fontWeight = font.defaultWeight || '700';
}

function getPhotoEditorTextDefaultWeight(fontKey) {
  const font = getPhotoEditorTextFontOption(fontKey);
  const weights = getPhotoEditorTextFontWeights(font.key);
  const defaultWeight = String(font.defaultWeight || weights[weights.length - 1]);

  return weights.includes(defaultWeight) ? defaultWeight : weights[weights.length - 1];
}

function getClosestPhotoEditorTextWeight(fontKey, weight) {
  const weights = getPhotoEditorTextFontWeights(fontKey);
  const requestedWeight = String(weight || '');

  if (weights.includes(requestedWeight)) {
    return requestedWeight;
  }

  const numericWeight = Number(requestedWeight);

  if (!Number.isFinite(numericWeight)) {
    return getPhotoEditorTextDefaultWeight(fontKey);
  }

  return weights.reduce((closest, candidate) => {
    const closestDistance = Math.abs(Number(closest) - numericWeight);
    const candidateDistance = Math.abs(Number(candidate) - numericWeight);
    return candidateDistance < closestDistance ? candidate : closest;
  }, weights[0]);
}

function normalizePhotoEditorTextColor(value, fallback = '#ffffff') {
  const color = String(value || '').trim();
  return /^#[0-9a-f]{6}$/i.test(color) ? color : fallback;
}

function createPhotoEditorTextOverlayId() {
  return `text-${Date.now().toString(36)}-${Math.random()
    .toString(36)
    .slice(2, 8)}`;
}

function normalizePhotoEditorMaskMode(value) {
  return PHOTO_EDITOR_MASK_MODES.includes(value) ? value : 'normal';
}

function normalizePhotoEditorInternalMaskMode(value) {
  return PHOTO_EDITOR_INTERNAL_MASK_MODES.includes(value) ? value : 'normal';
}

function normalizePhotoEditorAdjustmentTarget(value) {
  return PHOTO_EDITOR_ADJUSTMENT_TARGETS.includes(value) ? value : 'whole';
}

function getDefaultPhotoEditorSubjectMaskState(overrides = {}) {
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

function normalizePhotoEditorSubjectMaskState(subjectMask = {}) {
  const defaults = getDefaultPhotoEditorSubjectMaskState();
  const maskDataUrl =
    typeof subjectMask?.maskDataUrl === 'string' &&
    subjectMask.maskDataUrl.startsWith('data:image/')
      ? subjectMask.maskDataUrl
      : '';
  const enabled = Boolean(subjectMask?.enabled && maskDataUrl);
  const statusValues = ['none', 'loading', 'ready', 'failed'];
  const sourceValues = [
    'none',
    'lightweight',
    'standard',
    'high-quality',
    'manual',
    'imported',
    'dummy',
  ];

  return {
    enabled,
    status: statusValues.includes(subjectMask?.status)
      ? subjectMask.status
      : enabled
        ? 'ready'
        : defaults.status,
    source: sourceValues.includes(subjectMask?.source)
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
      typeof subjectMask?.createdAt === 'string' ? subjectMask.createdAt : null,
    updatedAt:
      typeof subjectMask?.updatedAt === 'string' ? subjectMask.updatedAt : null,
    errorMessage:
      typeof subjectMask?.errorMessage === 'string'
        ? subjectMask.errorMessage
        : '',
  };
}

function getDefaultPhotoEditorTextState(overrides = {}) {
  const defaultFont = getPhotoEditorTextFontOption('system');

  return {
    id: createPhotoEditorTextOverlayId(),
    enabled: false,
    text: '',
    x: 0.5,
    y: 0.5,
    rotation: 0,
    fontKey: defaultFont.key,
    fontFamily: defaultFont.family,
    size: 64,
    color: '#ffffff',
    weight: getPhotoEditorTextDefaultWeight(defaultFont.key),
    strokeType: 'outline',
    strokeWidth: 4,
    strokeColor: '#111827',
    fillTransparent: false,
    maskMode: 'normal',
    letterSpacing: 0,
    ...overrides,
  };
}

function normalizePhotoEditorTextState(textState = {}) {
  const defaults = getDefaultPhotoEditorTextState();
  const text = String(textState?.text || '').slice(0, 160);
  const font = getPhotoEditorTextFontOption(textState?.fontKey);
  const strokeType = PHOTO_EDITOR_TEXT_STROKE_TYPES.includes(
    textState?.strokeType
  )
    ? textState.strokeType
    : defaults.strokeType;
  const weight = getClosestPhotoEditorTextWeight(font.key, textState?.weight);
  const id =
    typeof textState?.id === 'string' && textState.id.trim()
      ? textState.id.trim()
      : defaults.id;

  return {
    id,
    enabled: Boolean(textState?.enabled || text.trim()),
    text,
    x: clampNumber(textState?.x, 0, 1, defaults.x),
    y: clampNumber(textState?.y, 0, 1, defaults.y),
    rotation: normalizePhotoEditorMaskRotation(textState?.rotation),
    fontKey: font.key,
    fontFamily: font.family,
    size: clampNumber(textState?.size, 12, 220, defaults.size),
    color: normalizePhotoEditorTextColor(textState?.color, defaults.color),
    weight,
    strokeType,
    strokeWidth: clampNumber(
      textState?.strokeWidth,
      0,
      28,
      defaults.strokeWidth
    ),
    strokeColor: normalizePhotoEditorTextColor(
      textState?.strokeColor,
      defaults.strokeColor
    ),
    fillTransparent: Boolean(textState?.fillTransparent),
    maskMode: normalizePhotoEditorMaskMode(textState?.maskMode),
    letterSpacing: clampNumber(
      textState?.letterSpacing,
      -10,
      40,
      defaults.letterSpacing
    ),
  };
}

function normalizePhotoEditorTextOverlays(textOverlays = []) {
  return (Array.isArray(textOverlays) ? textOverlays : [])
    .map(normalizePhotoEditorTextState)
    .filter(
      (textOverlay) =>
        Boolean(textOverlay.id) || textOverlay.enabled || textOverlay.text.trim()
    )
    .slice(0, 20);
}

function normalizePhotoEditorRulerGuides(rulerGuides = {}) {
  const normalizeAxisGuides = (values) => {
    const sortedValues = (Array.isArray(values) ? values : [])
      .map((value) => Number(value))
      .filter((value) => Number.isFinite(value))
      .map((value) => clampNumber(value, 0, 1, 0))
      .sort((left, right) => left - right);
    const uniqueValues = [];

    for (const value of sortedValues) {
      if (
        uniqueValues.length === 0 ||
        Math.abs(uniqueValues[uniqueValues.length - 1] - value) > 0.001
      ) {
        uniqueValues.push(value);
      }
    }

    return uniqueValues.slice(-PHOTO_EDITOR_RULER_GUIDE_LIMIT);
  };

  return {
    x: normalizeAxisGuides(rulerGuides?.x),
    y: normalizeAxisGuides(rulerGuides?.y),
  };
}

function getPhotoEditorTextCollectionFromState(state = photoEditorState) {
  const textOverlays = normalizePhotoEditorTextOverlays(
    Array.isArray(state?.textOverlays)
      ? state.textOverlays
      : state?.textOverlay
        ? [state.textOverlay]
        : []
  );
  const activeTextId =
    state?.activeTextId &&
    textOverlays.some((textOverlay) => textOverlay.id === state.activeTextId)
      ? state.activeTextId
      : '';

  return {
    textOverlays,
    activeTextId,
  };
}

function setPhotoEditorTextCollection(textOverlays, activeTextId = '') {
  if (!photoEditorState) {
    return;
  }

  const normalizedOverlays = normalizePhotoEditorTextOverlays(textOverlays);
  photoEditorState.textOverlays = normalizedOverlays;
  photoEditorState.activeTextId =
    activeTextId &&
    normalizedOverlays.some((textOverlay) => textOverlay.id === activeTextId)
      ? activeTextId
      : '';
  photoEditorState.textOverlay =
    normalizedOverlays.find(
      (textOverlay) => textOverlay.id === photoEditorState.activeTextId
    ) ||
    getDefaultPhotoEditorTextState();
}

function getPhotoEditorActiveTextOverlay() {
  const collection = getPhotoEditorTextCollectionFromState();
  return (
    collection.textOverlays.find(
      (textOverlay) => textOverlay.id === collection.activeTextId
    ) || null
  );
}

function createPhotoEditorImageOverlayId() {
  return `image-${Date.now().toString(36)}-${Math.random()
    .toString(36)
    .slice(2, 8)}`;
}

function normalizePhotoEditorOverlayAsset(rawAsset = {}) {
  const id =
    typeof rawAsset?.id === 'string' && rawAsset.id.trim()
      ? rawAsset.id.trim()
      : '';
  const fileUrl =
    typeof rawAsset?.fileUrl === 'string' && rawAsset.fileUrl.trim()
      ? rawAsset.fileUrl.trim()
      : '';

  if (!id || !fileUrl) {
    return null;
  }

  return {
    id,
    fileName:
      typeof rawAsset?.fileName === 'string' && rawAsset.fileName.trim()
        ? rawAsset.fileName.trim()
        : id,
    fileUrl,
    width: Math.max(0, Number(rawAsset?.width) || 0),
    height: Math.max(0, Number(rawAsset?.height) || 0),
    sizeBytes: Math.max(0, Number(rawAsset?.sizeBytes) || 0),
    updatedAt: Math.max(0, Number(rawAsset?.updatedAt) || 0),
  };
}

function normalizePhotoEditorOverlayAssets(assets = []) {
  return (Array.isArray(assets) ? assets : [])
    .map(normalizePhotoEditorOverlayAsset)
    .filter(Boolean);
}

function getPhotoEditorImageOverlayAspectRatio(overlay = {}) {
  const width = Number(overlay?.naturalWidth || overlay?.width || overlay?.assetWidth);
  const height = Number(overlay?.naturalHeight || overlay?.height || overlay?.assetHeight);
  return width > 0 && height > 0 ? width / height : 1;
}

function normalizePhotoEditorImageOverlayBlendMode(value) {
  return PHOTO_EDITOR_IMAGE_OVERLAY_BLEND_MODES.includes(value)
    ? value
    : 'source-over';
}

function getDefaultPhotoEditorImageOverlayState(asset = {}, overrides = {}) {
  const normalizedAsset = normalizePhotoEditorOverlayAsset(asset) || {};
  const aspectRatio = getPhotoEditorImageOverlayAspectRatio({
    width: normalizedAsset.width,
    height: normalizedAsset.height,
  });
  let defaultWidth = PHOTO_EDITOR_IMAGE_OVERLAY_DEFAULT_EDGE;
  let defaultHeight = defaultWidth / Math.max(0.05, aspectRatio);

  if (defaultHeight > 0.72) {
    defaultHeight = 0.72;
    defaultWidth = defaultHeight * aspectRatio;
  }

  return {
    id: createPhotoEditorImageOverlayId(),
    assetId: normalizedAsset.id || '',
    fileName: normalizedAsset.fileName || 'overlay',
    fileUrl: normalizedAsset.fileUrl || '',
    x: 0.5,
    y: 0.5,
    width: clampNumber(defaultWidth, 0.08, 0.72, PHOTO_EDITOR_IMAGE_OVERLAY_DEFAULT_EDGE),
    height: clampNumber(defaultHeight, 0.08, 0.72, defaultWidth),
    opacity: 1,
    blendMode: 'source-over',
    maskMode: 'normal',
    naturalWidth: normalizedAsset.width || 0,
    naturalHeight: normalizedAsset.height || 0,
    ...overrides,
  };
}

function normalizePhotoEditorImageOverlayState(overlay = {}) {
  const defaults = getDefaultPhotoEditorImageOverlayState();
  const id =
    typeof overlay?.id === 'string' && overlay.id.trim()
      ? overlay.id.trim()
      : defaults.id;
  const fileUrl =
    typeof overlay?.fileUrl === 'string' && overlay.fileUrl.trim()
      ? overlay.fileUrl.trim()
      : '';
  const naturalWidth = Math.max(0, Number(overlay?.naturalWidth) || 0);
  const naturalHeight = Math.max(0, Number(overlay?.naturalHeight) || 0);

  return {
    id,
    assetId:
      typeof overlay?.assetId === 'string' && overlay.assetId.trim()
        ? overlay.assetId.trim()
        : '',
    fileName:
      typeof overlay?.fileName === 'string' && overlay.fileName.trim()
        ? overlay.fileName.trim()
        : 'overlay',
    fileUrl,
    x: clampNumber(overlay?.x, -0.5, 1.5, defaults.x),
    y: clampNumber(overlay?.y, -0.5, 1.5, defaults.y),
    width: clampNumber(
      overlay?.width,
      PHOTO_EDITOR_IMAGE_OVERLAY_MIN_SIZE,
      2,
      defaults.width
    ),
    height: clampNumber(
      overlay?.height,
      PHOTO_EDITOR_IMAGE_OVERLAY_MIN_SIZE,
      2,
      defaults.height
    ),
    opacity: clampNumber(overlay?.opacity, 0, 1, defaults.opacity),
    blendMode: normalizePhotoEditorImageOverlayBlendMode(
      overlay?.blendMode || defaults.blendMode
    ),
    maskMode: normalizePhotoEditorMaskMode(overlay?.maskMode),
    naturalWidth,
    naturalHeight,
  };
}

function normalizePhotoEditorImageOverlays(imageOverlays = []) {
  return (Array.isArray(imageOverlays) ? imageOverlays : [])
    .map(normalizePhotoEditorImageOverlayState)
    .filter((overlay) => Boolean(overlay.fileUrl))
    .slice(0, PHOTO_EDITOR_IMAGE_OVERLAY_LIMIT);
}

function getPhotoEditorImageOverlayCollectionFromState(state = photoEditorState) {
  const imageOverlays = normalizePhotoEditorImageOverlays(state?.imageOverlays);
  const activeImageOverlayId =
    state?.activeImageOverlayId &&
    imageOverlays.some((overlay) => overlay.id === state.activeImageOverlayId)
      ? state.activeImageOverlayId
      : '';

  return {
    imageOverlays,
    activeImageOverlayId,
  };
}

function setPhotoEditorImageOverlayCollection(
  imageOverlays,
  activeImageOverlayId = ''
) {
  if (!photoEditorState) {
    return;
  }

  const normalizedOverlays = normalizePhotoEditorImageOverlays(imageOverlays);
  photoEditorState.imageOverlays = normalizedOverlays;
  photoEditorState.activeImageOverlayId =
    activeImageOverlayId &&
    normalizedOverlays.some((overlay) => overlay.id === activeImageOverlayId)
      ? activeImageOverlayId
      : '';
}

function getPhotoEditorActiveImageOverlay() {
  const collection = getPhotoEditorImageOverlayCollectionFromState();
  return (
    collection.imageOverlays.find(
      (overlay) => overlay.id === collection.activeImageOverlayId
    ) || null
  );
}

function getPhotoEditorOverlayImageCacheEntry(overlay) {
  const normalizedOverlay = normalizePhotoEditorImageOverlayState(overlay);
  const cacheKey = normalizedOverlay.fileUrl;

  if (!cacheKey) {
    return null;
  }

  const cachedEntry = photoEditorOverlayImageCache.get(cacheKey);

  if (cachedEntry) {
    return cachedEntry;
  }

  const image = new Image();
  let resolveLoad = null;
  const entry = {
    image,
    loaded: false,
    failed: false,
    loadPromise: new Promise((resolve) => {
      resolveLoad = resolve;
    }),
  };

  image.onload = () => {
    entry.loaded = true;
    entry.failed = false;
    resolveLoad?.(entry);
    if (photoEditorState) {
      schedulePhotoEditorRender();
    }
  };
  image.onerror = () => {
    entry.failed = true;
    resolveLoad?.(entry);
  };
  image.decoding = 'async';
  image.src = normalizedOverlay.fileUrl;
  photoEditorOverlayImageCache.set(cacheKey, entry);

  return entry;
}

async function waitForPhotoEditorImageOverlaysToLoad(imageOverlays = []) {
  const loadPromises = normalizePhotoEditorImageOverlays(imageOverlays)
    .map((overlay) => getPhotoEditorOverlayImageCacheEntry(overlay)?.loadPromise)
    .filter(Boolean);

  if (loadPromises.length === 0) {
    return;
  }

  await Promise.all(loadPromises);
}

async function waitForPhotoEditorSubjectMaskToLoad(subjectMask = photoEditorState?.subjectMask) {
  const entry = getPhotoEditorSubjectMaskImageCacheEntry(subjectMask);

  if (!entry?.loadPromise) {
    return;
  }

  await entry.loadPromise;
}

function loadPhotoEditorRecentTextFonts() {
  try {
    const parsed = JSON.parse(
      localStorage.getItem(PHOTO_EDITOR_TEXT_RECENT_FONTS_STORAGE_KEY) || '[]'
    );
    return (Array.isArray(parsed) ? parsed : [])
      .filter((fontKey) => getPhotoEditorTextFontOption(fontKey).key === fontKey)
      .slice(0, PHOTO_EDITOR_TEXT_RECENT_FONT_LIMIT);
  } catch {
    return [];
  }
}

function savePhotoEditorRecentTextFonts() {
  try {
    localStorage.setItem(
      PHOTO_EDITOR_TEXT_RECENT_FONTS_STORAGE_KEY,
      JSON.stringify(photoEditorTextRecentFonts.slice(0, PHOTO_EDITOR_TEXT_RECENT_FONT_LIMIT))
    );
  } catch {
    // Font recents are a convenience only; rendering should continue normally.
  }
}

function rememberPhotoEditorTextFont(fontKey) {
  const font = getPhotoEditorTextFontOption(fontKey);
  photoEditorTextRecentFonts = [
    font.key,
    ...photoEditorTextRecentFonts.filter((key) => key !== font.key),
  ].slice(0, PHOTO_EDITOR_TEXT_RECENT_FONT_LIMIT);
  savePhotoEditorRecentTextFonts();
  renderPhotoEditorTextFontOptions();
}

async function loadPhotoEditorTextFont(textOverlay) {
  if (!document.fonts) {
    return;
  }

  const textState = normalizePhotoEditorTextState(textOverlay);
  const primaryFontFamily =
    getPhotoEditorTextFontOption(textState.fontKey).family.split(',')[0]?.trim() ||
    textState.fontFamily;

  try {
    await document.fonts.load(
      `${textState.weight} ${Math.round(textState.size)}px ${primaryFontFamily}`,
      textState.text || 'WorldShot'
    );
  } catch {
    // The canvas can still render with the browser fallback if loading fails.
  }
}

function getDefaultPhotoEditorBlurState() {
  return {
    mode: 'full',
    amount: 0,
    centerX: 0.5,
    centerY: 0.5,
    radius: 0.34,
    outerRadius: 0.52,
    isConfirmed: false,
  };
}

function normalizePhotoEditorBlurState(blur = {}) {
  const defaults = getDefaultPhotoEditorBlurState();
  const mode = PHOTO_EDITOR_BLUR_MODES.includes(blur?.mode)
    ? blur.mode
    : defaults.mode;
  const radius = clampNumber(
    blur?.radius,
    PHOTO_EDITOR_RADIAL_BLUR_MIN_RADIUS,
    PHOTO_EDITOR_RADIAL_BLUR_MAX_RADIUS - PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER,
    defaults.radius
  );
  const fallbackOuterRadius = clampNumber(
    Math.max(defaults.outerRadius, radius + 0.18),
    radius + PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER,
    PHOTO_EDITOR_RADIAL_BLUR_MAX_RADIUS,
    defaults.outerRadius
  );
  const outerRadius = clampNumber(
    blur?.outerRadius,
    radius + PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER,
    PHOTO_EDITOR_RADIAL_BLUR_MAX_RADIUS,
    fallbackOuterRadius
  );

  return {
    mode,
    amount: clampNumber(blur?.amount, 0, 100, defaults.amount),
    centerX: clampNumber(blur?.centerX, 0, 1, defaults.centerX),
    centerY: clampNumber(blur?.centerY, 0, 1, defaults.centerY),
    radius,
    outerRadius,
    isConfirmed: Boolean(blur?.isConfirmed),
  };
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
  const validChannels = PHOTO_EDITOR_CURVE_CHANNELS[mode].map(
    (channel) => channel.key
  );
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

function getDefaultPhotoEditorAutoEnhanceState() {
  return {
    enabled: false,
    originalValuesBeforeAuto: null,
    generatedValues: null,
    strength: PHOTO_EDITOR_AUTO_ENHANCE_DEFAULT_STRENGTH,
    presetKey: '',
  };
}

function normalizePhotoEditorAutoEnhanceState(autoEnhance = {}) {
  const defaults = getDefaultPhotoEditorAutoEnhanceState();
  const strength = clampNumber(
    autoEnhance?.strength,
    0,
    100,
    defaults.strength
  );
  const originalValues = autoEnhance?.originalValuesBeforeAuto
    ? clonePhotoEditValues(autoEnhance.originalValuesBeforeAuto)
    : null;
  const generatedValues = autoEnhance?.generatedValues
    ? clonePhotoEditValues(autoEnhance.generatedValues)
    : null;

  return {
    enabled: Boolean(autoEnhance?.enabled && originalValues && generatedValues),
    originalValuesBeforeAuto: originalValues,
    generatedValues,
    strength,
    presetKey:
      typeof autoEnhance?.presetKey === 'string'
        ? autoEnhance.presetKey
        : defaults.presetKey,
  };
}

function createPhotoEditorState(photo) {
  return {
    sourcePhoto: photo,
    sourceImage: null,
    loadToken: null,
    values: clonePhotoEditValues(),
    adjustmentTarget: 'whole',
    crop: getDefaultPhotoEditorCropState(),
    blur: getDefaultPhotoEditorBlurState(),
    curve: getDefaultPhotoEditorCurveState(),
    autoEnhance: getDefaultPhotoEditorAutoEnhanceState(),
    textOverlay: getDefaultPhotoEditorTextState(),
    textOverlays: [],
    activeTextId: '',
    imageOverlays: [],
    activeImageOverlayId: '',
    subjectMask: getDefaultPhotoEditorSubjectMaskState(),
    rulerGuides: normalizePhotoEditorRulerGuides(),
    draftRulerGuide: null,
    dragInitialRulerGuide: null,
    snapGuide: null,
    exportSettings: getDefaultPhotoEditorExportSettings(),
    masks: [],
    maskTool: 'none',
    maskShape: 'rect',
    maskStrengths: { ...PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS },
    blurStrength: PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS.blur,
    fillColor: '#111827',
    dragMode: null,
    dragStart: null,
    dragInitialCrop: null,
    dragInitialSourceRect: null,
    dragInitialBlur: null,
    dragInitialMask: null,
    dragInitialText: null,
    dragInitialImageOverlay: null,
    dragCurvePointIndex: null,
    draftMask: null,
    showOriginalPreview: false,
    showRuleOfThirdsGrid: false,
    showRulers: false,
    isInteractivePreview: false,
    isCropInteractivePreview: false,
    lastClipInfo: null,
    curveHistogram: null,
    curveHistogramKey: '',
    isSaving: false,
    historyUndoStack: [],
    historyRedoStack: [],
    pendingHistorySnapshot: null,
    historyCommitTimer: null,
    isRestoringHistory: false,
  };
}

function clonePhotoEditorHistoryData(value) {
  return JSON.parse(JSON.stringify(value || null));
}

function capturePhotoEditorHistorySnapshot() {
  if (!photoEditorState) {
    return null;
  }

  const textCollection = getPhotoEditorTextCollectionFromState();
  const imageOverlayCollection = getPhotoEditorImageOverlayCollectionFromState();

  return {
    values: clonePhotoEditValues(photoEditorState.values),
    adjustmentTarget: normalizePhotoEditorAdjustmentTarget(
      photoEditorState.adjustmentTarget
    ),
    crop: normalizePhotoEditorCropState(photoEditorState.crop),
    blur: normalizePhotoEditorBlurState(photoEditorState.blur),
    curve: normalizePhotoEditorCurveState(photoEditorState.curve),
    autoEnhance: normalizePhotoEditorAutoEnhanceState(
      photoEditorState.autoEnhance
    ),
    textOverlays: clonePhotoEditorHistoryData(textCollection.textOverlays),
    activeTextId: textCollection.activeTextId,
    imageOverlays: clonePhotoEditorHistoryData(
      imageOverlayCollection.imageOverlays
    ),
    activeImageOverlayId: imageOverlayCollection.activeImageOverlayId,
    subjectMask: normalizePhotoEditorSubjectMaskState(photoEditorState.subjectMask),
    rulerGuides: normalizePhotoEditorRulerGuides(photoEditorState.rulerGuides),
    exportSettings: normalizePhotoEditorExportSettings(
      photoEditorState.exportSettings
    ),
    masks: clonePhotoEditorHistoryData(photoEditorState.masks || []),
    maskShape: photoEditorState.maskShape || 'rect',
    maskStrengths: {
      ...PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS,
      ...(photoEditorState.maskStrengths || {}),
    },
    blurStrength: clampNumber(
      photoEditorState.blurStrength,
      0,
      100,
      PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS.blur
    ),
    fillColor: photoEditorState.fillColor || '#111827',
  };
}

function arePhotoEditorHistorySnapshotsEqual(left, right) {
  return JSON.stringify(left || null) === JSON.stringify(right || null);
}

function syncPhotoEditorHistoryControls() {
  if (photoEditorUndoButton) {
    photoEditorUndoButton.disabled = !(
      photoEditorState?.historyUndoStack?.length > 0 ||
      photoEditorState?.pendingHistorySnapshot
    );
  }

  if (photoEditorRedoButton) {
    photoEditorRedoButton.disabled = !(
      photoEditorState?.historyRedoStack?.length > 0
    );
  }
}

function beginPhotoEditorHistoryMutation() {
  if (!photoEditorState || photoEditorState.isRestoringHistory) {
    return;
  }

  if (!photoEditorState.pendingHistorySnapshot) {
    photoEditorState.pendingHistorySnapshot = capturePhotoEditorHistorySnapshot();
  }
}

function clearPhotoEditorHistoryCommitTimer() {
  if (photoEditorState?.historyCommitTimer) {
    clearTimeout(photoEditorState.historyCommitTimer);
    photoEditorState.historyCommitTimer = null;
  }
}

function commitPhotoEditorHistoryMutation() {
  if (!photoEditorState || photoEditorState.isRestoringHistory) {
    return;
  }

  clearPhotoEditorHistoryCommitTimer();

  const previousSnapshot = photoEditorState.pendingHistorySnapshot;

  if (!previousSnapshot) {
    syncPhotoEditorHistoryControls();
    return;
  }

  photoEditorState.pendingHistorySnapshot = null;
  const currentSnapshot = capturePhotoEditorHistorySnapshot();

  if (!arePhotoEditorHistorySnapshotsEqual(previousSnapshot, currentSnapshot)) {
    photoEditorState.historyUndoStack.push(previousSnapshot);
    if (photoEditorState.historyUndoStack.length > PHOTO_EDITOR_HISTORY_LIMIT) {
      photoEditorState.historyUndoStack.shift();
    }
    photoEditorState.historyRedoStack = [];
  }

  syncPhotoEditorHistoryControls();
}

function schedulePhotoEditorHistoryCommit(delayMs = PHOTO_EDITOR_HISTORY_COMMIT_DELAY_MS) {
  if (!photoEditorState || !photoEditorState.pendingHistorySnapshot) {
    return;
  }

  clearPhotoEditorHistoryCommitTimer();
  photoEditorState.historyCommitTimer = setTimeout(() => {
    if (photoEditorState) {
      photoEditorState.historyCommitTimer = null;
    }
    commitPhotoEditorHistoryMutation();
  }, delayMs);
  syncPhotoEditorHistoryControls();
}

function applyPhotoEditorHistorySnapshot(snapshot) {
  if (!photoEditorState || !snapshot) {
    return;
  }

  photoEditorState.isRestoringHistory = true;
  photoEditorState.values = clonePhotoEditValues(snapshot.values);
  photoEditorState.adjustmentTarget = normalizePhotoEditorAdjustmentTarget(
    snapshot.adjustmentTarget
  );
  photoEditorState.crop = normalizePhotoEditorCropState(snapshot.crop);
  photoEditorState.blur = normalizePhotoEditorBlurState(snapshot.blur);
  photoEditorState.curve = normalizePhotoEditorCurveState(snapshot.curve);
  photoEditorState.autoEnhance = normalizePhotoEditorAutoEnhanceState(
    snapshot.autoEnhance
  );
  setPhotoEditorTextCollection(
    Array.isArray(snapshot.textOverlays)
      ? snapshot.textOverlays
      : snapshot.textOverlay
        ? [snapshot.textOverlay]
        : [],
    snapshot.activeTextId || ''
  );
  setPhotoEditorImageOverlayCollection(
    snapshot.imageOverlays || [],
    snapshot.activeImageOverlayId || ''
  );
  photoEditorState.subjectMask = normalizePhotoEditorSubjectMaskState(
    snapshot.subjectMask
  );
  photoEditorState.rulerGuides = normalizePhotoEditorRulerGuides(
    snapshot.rulerGuides
  );
  photoEditorState.exportSettings = normalizePhotoEditorExportSettings(
    snapshot.exportSettings
  );
  photoEditorState.masks = clonePhotoEditorHistoryData(snapshot.masks || []);
  photoEditorState.maskShape = snapshot.maskShape || 'rect';
  photoEditorState.maskStrengths = {
    ...PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS,
    ...(snapshot.maskStrengths || {}),
  };
  photoEditorState.blurStrength = clampNumber(
    snapshot.blurStrength,
    0,
    100,
    PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS.blur
  );
  photoEditorState.fillColor = snapshot.fillColor || '#111827';
  photoEditorState.maskTool = 'none';
  photoEditorState.dragMode = null;
  photoEditorState.dragStart = null;
  photoEditorState.dragInitialCrop = null;
  photoEditorState.dragInitialSourceRect = null;
  photoEditorState.dragInitialBlur = null;
  photoEditorState.dragInitialMask = null;
  photoEditorState.dragInitialText = null;
  photoEditorState.dragInitialImageOverlay = null;
  photoEditorState.draftRulerGuide = null;
  photoEditorState.dragInitialRulerGuide = null;
  photoEditorState.snapGuide = null;
  photoEditorState.dragCurvePointIndex = null;
  photoEditorState.draftMask = null;
  photoEditorState.lastClipInfo = null;
  photoEditorState.isCropInteractivePreview = false;
  photoEditorCanvas?.classList.remove(
    'is-panning',
    'is-mask-draft-active',
    'is-text-tool-active',
    'is-image-overlay-tool-active'
  );
  clearPhotoEditorAdjustmentLivePreview();
  syncPhotoEditorUi();
  schedulePhotoEditorRender();
  photoEditorState.isRestoringHistory = false;
}

function undoPhotoEditorEdit() {
  if (!photoEditorState) {
    return;
  }

  commitPhotoEditorHistoryMutation();
  const previousSnapshot = photoEditorState.historyUndoStack.pop();

  if (!previousSnapshot) {
    syncPhotoEditorHistoryControls();
    return;
  }

  photoEditorState.historyRedoStack.push(capturePhotoEditorHistorySnapshot());
  applyPhotoEditorHistorySnapshot(previousSnapshot);
  syncPhotoEditorHistoryControls();
}

function redoPhotoEditorEdit() {
  if (!photoEditorState) {
    return;
  }

  commitPhotoEditorHistoryMutation();
  const nextSnapshot = photoEditorState.historyRedoStack.pop();

  if (!nextSnapshot) {
    syncPhotoEditorHistoryControls();
    return;
  }

  photoEditorState.historyUndoStack.push(capturePhotoEditorHistorySnapshot());
  applyPhotoEditorHistorySnapshot(nextSnapshot);
  syncPhotoEditorHistoryControls();
}

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

function getPhotoEditorAccordionToggle(key) {
  return photoEditorAccordionToggles.find(
    (toggle) => toggle.dataset.photoEditorAccordionToggle === key
  );
}

function isPhotoEditorAccordionOpen(key) {
  const toggle = getPhotoEditorAccordionToggle(key);
  return toggle?.getAttribute('aria-expanded') === 'true';
}

function setPhotoEditorAccordionOpen(key, isOpen) {
  const toggle = getPhotoEditorAccordionToggle(key);
  const panelId = toggle?.getAttribute('aria-controls');
  const panel = panelId ? document.getElementById(panelId) : null;

  if (!toggle || !panel) {
    return;
  }

  toggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
  panel.hidden = !isOpen;
}

function initializePhotoEditorAccordions() {
  for (const toggle of photoEditorAccordionToggles) {
    const key = toggle.dataset.photoEditorAccordionToggle;

    if (!key) {
      continue;
    }

    setPhotoEditorAccordionOpen(key, false);
  }
}

function setPhotoEditorStatus(message = '') {
  if (photoEditorStatus) {
    photoEditorStatus.textContent = message;
  }
}

function setPhotoEditorSaving(isSaving) {
  if (!photoEditorState) {
    return;
  }

  photoEditorState.isSaving = Boolean(isSaving);

  if (photoEditorSaveButton) {
    photoEditorSaveButton.disabled = photoEditorState.isSaving;
    photoEditorSaveButton.textContent = photoEditorState.isSaving
      ? '保存中...'
      : '別名で保存';
  }

  if (photoEditorSubjectTransparentSaveButton) {
    photoEditorSubjectTransparentSaveButton.disabled =
      photoEditorState.isSaving || !hasPhotoEditorReadySubjectMask();
    photoEditorSubjectTransparentSaveButton.title =
      hasPhotoEditorReadySubjectMask()
        ? '被写体マスクを使って背景を透過したPNGとして保存します'
        : 'AI被写体選択でマスクを作成すると使用できます';
  }
}

function getPhotoEditorSliderMeta(key) {
  return PHOTO_EDIT_SLIDERS.find((slider) => slider.key === key) || null;
}

function getPhotoEditorAdjustmentInput(key) {
  return photoEditorAdjustmentList?.querySelector(
    `[data-photo-editor-slider="${key}"]`
  );
}

function getPhotoEditorAdjustmentValueElement(key) {
  return photoEditorAdjustmentList?.querySelector(
    `[data-photo-editor-value="${key}"]`
  );
}

function syncPhotoEditorAdjustmentControls(targetKey = null) {
  if (!photoEditorState) {
    return;
  }

  photoEditorState.adjustmentTarget = normalizePhotoEditorAdjustmentTarget(
    photoEditorState.adjustmentTarget
  );

  if (photoEditorAdjustmentTargetSelect) {
    photoEditorAdjustmentTargetSelect.value = photoEditorState.adjustmentTarget;
  }

  const sliders = targetKey
    ? PHOTO_EDIT_SLIDERS.filter((slider) => slider.key === targetKey)
    : PHOTO_EDIT_SLIDERS;

  for (const slider of sliders) {
    const value = photoEditorState.values[slider.key];
    const input = getPhotoEditorAdjustmentInput(slider.key);
    const valueElement = getPhotoEditorAdjustmentValueElement(slider.key);

    if (input) {
      input.value = String(value);
    }

    if (valueElement) {
      valueElement.textContent = String(value);
    }
  }
}

function syncPhotoEditorSubjectMaskControls() {
  if (!photoEditorState) {
    return;
  }

  photoEditorState.subjectMask = normalizePhotoEditorSubjectMaskState(
    photoEditorState.subjectMask
  );
  const subjectMask = photoEditorState.subjectMask;
  const hasMask = Boolean(subjectMask.enabled && subjectMask.maskDataUrl);

  if (photoEditorSubjectStatus) {
    if (hasMask) {
      const sourceLabel =
        subjectMask.source === 'dummy'
          ? 'ダミーマスク'
          : subjectMask.source === 'lightweight'
            ? '軽量AIマスク'
          : subjectMask.source === 'standard'
            ? '標準AIマスク'
          : subjectMask.source === 'high-quality'
            ? '高精度AIマスク'
          : subjectMask.source === 'imported'
            ? '読み込みマスク'
            : 'マスク';
      const sizeText =
        subjectMask.width > 0 && subjectMask.height > 0
          ? ` / ${subjectMask.width} x ${subjectMask.height}`
          : '';
      photoEditorSubjectStatus.textContent = `${sourceLabel} 使用中${sizeText}`;
    } else if (subjectMask.status === 'failed' && subjectMask.errorMessage) {
      photoEditorSubjectStatus.textContent = subjectMask.errorMessage;
      } else {
        const autoModel = getPhotoEditorAutoSubjectModel();
        if (autoModel?.id === 'withoutbg-snap') {
          photoEditorSubjectStatus.textContent =
            '現在: 標準AIで実行されます';
        } else if (autoModel?.id === 'u2netp') {
          photoEditorSubjectStatus.textContent =
            '現在: 軽量AIで実行されます / 標準AIをダウンロードすると検出精度が向上します';
        } else {
          photoEditorSubjectStatus.textContent = 'AIモデルを利用できません';
        }
    }
  }

  if (photoEditorSubjectDeleteButton) {
    photoEditorSubjectDeleteButton.disabled = !hasMask;
  }

  if (photoEditorSubjectTransparentSaveButton) {
    photoEditorSubjectTransparentSaveButton.disabled =
      !hasMask || photoEditorState.isSaving || subjectMask.status === 'loading';
    photoEditorSubjectTransparentSaveButton.title = hasMask
      ? '被写体マスクを使って背景を透過したPNGとして保存します'
      : 'AI被写体選択でマスクを作成すると使用できます';
  }

  if (photoEditorSubjectAutoButton) {
    const model = getPhotoEditorAutoSubjectModel();
    const isLoading = subjectMask.status === 'loading';
    photoEditorSubjectAutoButton.disabled = !model?.ready || isLoading;
    photoEditorSubjectAutoButton.title = model?.ready
      ? `現在: ${model.displayName || model.label}で実行します`
      : model?.message || 'AIモデルを利用できません';
  }

  const standardModel = getPhotoEditorSubjectModel('withoutbg-snap');
  const highQualityModel = getPhotoEditorSubjectModel('withoutbg-focus');

  if (photoEditorSubjectStandardDownloadButton) {
    photoEditorSubjectStandardDownloadButton.hidden =
      Boolean(standardModel?.ready) || !standardModel?.canDownload;
    photoEditorSubjectStandardDownloadButton.disabled =
      subjectMask.status === 'loading' || !standardModel?.canDownload;
    photoEditorSubjectStandardDownloadButton.title =
      standardModel?.description || '標準AIをダウンロード';
  }

  if (photoEditorSubjectHighQualityDownloadButton) {
    photoEditorSubjectHighQualityDownloadButton.hidden =
      Boolean(highQualityModel?.ready) || !highQualityModel?.canDownload;
    photoEditorSubjectHighQualityDownloadButton.disabled =
      subjectMask.status === 'loading' || !highQualityModel?.canDownload;
    photoEditorSubjectHighQualityDownloadButton.title =
      highQualityModel?.description || '高精度AIをダウンロード';
  }

  if (photoEditorSubjectHighQualityButton) {
    const model = highQualityModel;
    const isLoading = subjectMask.status === 'loading';
    photoEditorSubjectHighQualityButton.disabled = !model?.ready || isLoading;
    photoEditorSubjectHighQualityButton.hidden = !model?.ready;
    photoEditorSubjectHighQualityButton.title = model?.ready
      ? `${model.modelName} / ${model.license} / ${model.sizeLabel || ''}`
      : model?.message || '設定から高精度AIモデルをダウンロードしてください';
  }

  if (photoEditorSubjectResetButton) {
    photoEditorSubjectResetButton.disabled = !hasMask;
  }

  if (photoEditorSubjectShowMaskInput) {
    photoEditorSubjectShowMaskInput.disabled = !hasMask;
    photoEditorSubjectShowMaskInput.classList.toggle(
      'is-active',
      Boolean(subjectMask.showOverlay)
    );
    photoEditorSubjectShowMaskInput.setAttribute(
      'aria-pressed',
      String(Boolean(subjectMask.showOverlay))
    );
  }

  if (photoEditorSubjectInvertInput) {
    photoEditorSubjectInvertInput.disabled = !hasMask;
    photoEditorSubjectInvertInput.classList.toggle(
      'is-active',
      Boolean(subjectMask.invert)
    );
    photoEditorSubjectInvertInput.setAttribute(
      'aria-pressed',
      String(Boolean(subjectMask.invert))
    );
  }

}

function updatePhotoEditorAdjustmentTarget(target) {
  if (!photoEditorState) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.adjustmentTarget = normalizePhotoEditorAdjustmentTarget(target);
  syncPhotoEditorAdjustmentControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function hasPhotoEditorReadySubjectMask(state = photoEditorState) {
  const subjectMask = normalizePhotoEditorSubjectMaskState(state?.subjectMask);
  return Boolean(subjectMask.enabled && subjectMask.maskDataUrl);
}

function updatePhotoEditorSubjectMask(nextValues = {}, { interactive = true } = {}) {
  if (!photoEditorState) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.subjectMask = normalizePhotoEditorSubjectMaskState({
    ...photoEditorState.subjectMask,
    ...nextValues,
  });
  syncPhotoEditorSubjectMaskControls();
  schedulePhotoEditorRender({
    debounceMs: interactive ? PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS : 0,
    interactive,
  });
  schedulePhotoEditorHistoryCommit();
}

function clearPhotoEditorSubjectMask() {
  if (!photoEditorState) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.subjectMask = getDefaultPhotoEditorSubjectMaskState();
  syncPhotoEditorSubjectMaskControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function getPhotoEditorSubjectModel(modelId) {
  const models = Array.isArray(photoEditorAiSubjectModelStatus?.models)
    ? photoEditorAiSubjectModelStatus.models
    : [];
  return models.find((model) => model?.id === modelId) || null;
}

function getPhotoEditorAutoSubjectModel() {
  const standardModel = getPhotoEditorSubjectModel('withoutbg-snap');

  if (standardModel?.ready) {
    return standardModel;
  }

  return getPhotoEditorSubjectModel('u2netp');
}

async function refreshPhotoEditorAiSubjectModelStatus({ silent = false } = {}) {
  if (!window.electronAPI.getAiSubjectModelStatus) {
    if (!silent) {
      setPhotoEditorStatus('AIモデル情報を取得できません');
    }
    return null;
  }

  try {
    const result = await window.electronAPI.getAiSubjectModelStatus();

    if (!result?.ok) {
      throw new Error(result?.message || 'AIモデル情報を取得できません');
    }

    photoEditorAiSubjectModelStatus = result;
    syncPhotoEditorSubjectMaskControls();
    renderSettingsAiSubjectModels();
    return result;
  } catch (error) {
    photoEditorAiSubjectModelStatus = {
      ok: false,
      message: error.message,
      models: [],
    };
    syncPhotoEditorSubjectMaskControls();
    renderSettingsAiSubjectModels();
    if (!silent) {
      setPhotoEditorStatus(error.message);
    }
    return null;
  }
}

async function getPhotoEditorSubjectSession(model, modelUrl = '') {
  const targetModelUrl = modelUrl || model?.modelUrl || '';

  if (!model?.ready || !targetModelUrl) {
    throw new Error(model?.message || 'AIモデルを利用できません');
  }

  if (!window.ort?.InferenceSession || !window.ort?.Tensor) {
    throw new Error('ONNX Runtimeを読み込めませんでした');
  }

  if (model.wasmBaseUrl && window.ort.env?.wasm) {
    window.ort.env.wasm.wasmPaths = model.wasmBaseUrl;
    window.ort.env.wasm.numThreads = 1;
  }

  const sessionKey = `${model.id}:${targetModelUrl}`;

  if (photoEditorSubjectModelSessionCache.has(sessionKey)) {
    return photoEditorSubjectModelSessionCache.get(sessionKey);
  }

  const session = await window.ort.InferenceSession.create(targetModelUrl, {
    executionProviders: ['wasm'],
    graphOptimizationLevel: 'all',
  });
  photoEditorSubjectModelSessionCache.set(sessionKey, session);
  return session;
}

function normalizePhotoEditorSubjectOutputTensor(result) {
  const outputData = result?.outputData;
  let data = outputData;

  if (outputData instanceof ArrayBuffer) {
    data = new Float32Array(outputData);
  } else if (Array.isArray(outputData)) {
    data = Float32Array.from(outputData);
  } else if (ArrayBuffer.isView(outputData)) {
    data = outputData;
  }

  if (!data || typeof data.length !== 'number') {
    throw new Error('AIマスクの出力を読み取れませんでした');
  }

  return {
    data,
    dims: Array.isArray(result?.outputShape) ? result.outputShape : [],
    type: result?.outputType || 'float32',
  };
}

function getPhotoEditorOrtInputChannels(session, fallback = 3) {
  const inputName = session?.inputNames?.[0] || '';
  const metadata = inputName ? session?.inputMetadata?.[inputName] : null;
  const dims = metadata?.dimensions || metadata?.dims || [];
  const channelCount = Number(dims?.[1]) || 0;

  if (channelCount > 0) {
    return channelCount;
  }

  if (/(rgbd_alpha|rgbdm)/i.test(inputName)) {
    return 5;
  }

  if (/(rgb_alpha|rgbd)/i.test(inputName)) {
    return 4;
  }

  return channelCount > 0 ? channelCount : fallback;
}

function createPhotoEditorCanvas(width, height) {
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(width));
  canvas.height = Math.max(1, Math.round(height));
  return canvas;
}

function drawPhotoEditorImageToCanvas(image, width, height) {
  const canvas = createPhotoEditorCanvas(width, height);
  const ctx = canvas.getContext('2d', { willReadFrequently: true });

  if (!ctx) {
    throw new Error('AI入力画像を作成できませんでした');
  }

  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
  return canvas;
}

function getPhotoEditorWithoutBgSize(
  sourceWidth,
  sourceHeight,
  targetWidth,
  targetHeight,
  ensureMultipleOf = 1
) {
  const scale = Math.max(
    targetWidth / Math.max(1, sourceWidth),
    targetHeight / Math.max(1, sourceHeight)
  );
  const constrain = (value, target) => {
    const multiple = Math.max(1, ensureMultipleOf);
    let next = Math.round(value / multiple) * multiple;

    if (next < target) {
      next = Math.ceil(value / multiple) * multiple;
    }

    return Math.max(multiple, next);
  };

  return {
    width: constrain(sourceWidth * scale, targetWidth),
    height: constrain(sourceHeight * scale, targetHeight),
  };
}

function createPhotoEditorRgbTensorFromCanvas(
  canvas,
  {
    mean = [0, 0, 0],
    std = [1, 1, 1],
  } = {}
) {
  const ctx = canvas.getContext('2d', { willReadFrequently: true });

  if (!ctx) {
    throw new Error('AI入力画像を作成できませんでした');
  }

  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const sourceData = imageData.data;
  const planeSize = canvas.width * canvas.height;
  const tensorData = new Float32Array(1 * 3 * planeSize);

  for (let index = 0; index < planeSize; index += 1) {
    const sourceIndex = index * 4;
    tensorData[index] = (sourceData[sourceIndex] / 255 - mean[0]) / std[0];
    tensorData[planeSize + index] =
      (sourceData[sourceIndex + 1] / 255 - mean[1]) / std[1];
    tensorData[planeSize * 2 + index] =
      (sourceData[sourceIndex + 2] / 255 - mean[2]) / std[2];
  }

  return new window.ort.Tensor('float32', tensorData, [
    1,
    3,
    canvas.height,
    canvas.width,
  ]);
}

function createPhotoEditorGrayCanvasFromValues(values, width, height) {
  const canvas = createPhotoEditorCanvas(width, height);
  const ctx = canvas.getContext('2d', { willReadFrequently: true });

  if (!ctx) {
    throw new Error('AIマスクを作成できませんでした');
  }

  const imageData = ctx.createImageData(canvas.width, canvas.height);
  const targetData = imageData.data;
  const pixelCount = canvas.width * canvas.height;

  for (let index = 0; index < pixelCount; index += 1) {
    const value = clampColorChannel(values[index] || 0);
    const targetIndex = index * 4;
    targetData[targetIndex] = value;
    targetData[targetIndex + 1] = value;
    targetData[targetIndex + 2] = value;
    targetData[targetIndex + 3] = 255;
  }

  ctx.putImageData(imageData, 0, 0);
  return canvas;
}

function getPhotoEditorTensorSpatialSize(outputTensor, fallbackWidth, fallbackHeight) {
  const dims = Array.isArray(outputTensor?.dims) ? outputTensor.dims : [];
  const dataLength = Number(outputTensor?.data?.length) || 0;
  let width = Number(fallbackWidth) || 0;
  let height = Number(fallbackHeight) || 0;

  if (dims.length >= 4 && Number(dims[dims.length - 1]) > 0 && Number(dims[dims.length - 2]) > 0) {
    width = Number(dims[dims.length - 1]);
    height = Number(dims[dims.length - 2]);
  } else if (dims.length >= 2 && Number(dims[dims.length - 1]) > 0 && Number(dims[dims.length - 2]) > 0) {
    width = Number(dims[dims.length - 1]);
    height = Number(dims[dims.length - 2]);
  }

  if (width > 0 && height > 0 && width * height <= dataLength) {
    return { width, height };
  }

  const squareSize = Math.round(Math.sqrt(dataLength));

  if (squareSize > 0 && squareSize * squareSize === dataLength) {
    return { width: squareSize, height: squareSize };
  }

  return {
    width: Math.max(1, width || squareSize || 1),
    height: Math.max(1, height || squareSize || 1),
  };
}

function createPhotoEditorAlphaCanvasFromTensor(outputTensor, width, height) {
  const outputData = outputTensor?.data;

  if (!outputData || outputData.length === 0) {
    throw new Error('AIマスクの出力が空でした');
  }

  const tensorSize = getPhotoEditorTensorSpatialSize(outputTensor, width, height);
  const outputWidth = tensorSize.width;
  const outputHeight = tensorSize.height;
  const pixelCount = outputWidth * outputHeight;
  const values = new Uint8ClampedArray(pixelCount);
  let minValue = Infinity;
  let maxValue = -Infinity;

  for (let index = 0; index < Math.min(pixelCount, outputData.length); index += 1) {
    const value = Number(outputData[index]);
    if (Number.isFinite(value)) {
      minValue = Math.min(minValue, value);
      maxValue = Math.max(maxValue, value);
    }
  }

  const shouldNormalize = maxValue > 1 || minValue < 0;
  const range = maxValue > minValue ? maxValue - minValue : 1;

  for (let index = 0; index < pixelCount; index += 1) {
    const rawValue = Number(outputData[index] || 0);
    const normalizedValue = shouldNormalize
      ? (rawValue - minValue) / range
      : rawValue;
    values[index] = clampColorChannel(normalizedValue * 255);
  }

  return createPhotoEditorGrayCanvasFromValues(values, outputWidth, outputHeight);
}

async function runPhotoEditorOrtSession(session, feeds) {
  const outputs = await session.run(feeds);
  const outputName = session.outputNames?.[0] || Object.keys(outputs)[0];
  return outputs[outputName];
}

async function runPhotoEditorWithoutBgDepthStage(model, sourceImage) {
  const sourceSize = getPhotoEditorImageSize(sourceImage);
  const depthSize = getPhotoEditorWithoutBgSize(
    sourceSize.width,
    sourceSize.height,
    518,
    518,
    14
  );
  const inputCanvas = drawPhotoEditorImageToCanvas(
    sourceImage,
    depthSize.width,
    depthSize.height
  );
  const tensor = createPhotoEditorRgbTensorFromCanvas(inputCanvas, {
    mean: [0.485, 0.456, 0.406],
    std: [0.229, 0.224, 0.225],
  });
  const session = await getPhotoEditorSubjectSession(
    model,
    model.modelFiles?.depth
  );
  const inputName = session.inputNames?.[0] || 'image';
  const outputTensor = await runPhotoEditorOrtSession(session, {
    [inputName]: tensor,
  });

  return createPhotoEditorAlphaCanvasFromTensor(
    outputTensor,
    depthSize.width,
    depthSize.height
  );
}

async function runPhotoEditorWithoutBgIsnetStage(model, sourceImage) {
  const inputCanvas = drawPhotoEditorImageToCanvas(sourceImage, 1024, 1024);
  const tensor = createPhotoEditorRgbTensorFromCanvas(inputCanvas, {
    mean: [0.5, 0.5, 0.5],
    std: [1, 1, 1],
  });
  const session = await getPhotoEditorSubjectSession(
    model,
    model.modelFiles?.isnet
  );
  const inputName = session.inputNames?.[0] || 'input';
  const outputTensor = await runPhotoEditorOrtSession(session, {
    [inputName]: tensor,
  });

  return createPhotoEditorAlphaCanvasFromTensor(outputTensor, 1024, 1024);
}

function getPhotoEditorCanvasChannel(canvas, width, height, channel = 0) {
  const resizedCanvas =
    canvas.width === width && canvas.height === height
      ? canvas
      : drawPhotoEditorImageToCanvas(canvas, width, height);
  const ctx = resizedCanvas.getContext('2d', { willReadFrequently: true });

  if (!ctx) {
    throw new Error('AI入力画像を作成できませんでした');
  }

  const imageData = ctx.getImageData(0, 0, width, height);
  const sourceData = imageData.data;
  const values = new Float32Array(width * height);

  for (let index = 0; index < values.length; index += 1) {
    values[index] = sourceData[index * 4 + channel] / 255;
  }

  return values;
}

function createPhotoEditorWithoutBgPipelineTensor({
  sourceImage,
  width,
  height,
  depthCanvas = null,
  isnetCanvas = null,
  alphaCanvas = null,
  channels = 4,
  alphaFirstForFourthChannel = false,
}) {
  const rgbCanvas = drawPhotoEditorImageToCanvas(sourceImage, width, height);
  const rgbCtx = rgbCanvas.getContext('2d', { willReadFrequently: true });

  if (!rgbCtx) {
    throw new Error('AI入力画像を作成できませんでした');
  }

  const rgbData = rgbCtx.getImageData(0, 0, width, height).data;
  const planeSize = width * height;
  const tensorData = new Float32Array(1 * channels * planeSize);
  const depthValues = depthCanvas
    ? getPhotoEditorCanvasChannel(depthCanvas, width, height, 0)
    : null;
  const isnetValues = isnetCanvas
    ? getPhotoEditorCanvasChannel(isnetCanvas, width, height, 0)
    : null;
  const alphaValues = alphaCanvas
    ? getPhotoEditorCanvasChannel(alphaCanvas, width, height, 0)
    : null;

  for (let index = 0; index < planeSize; index += 1) {
    const sourceIndex = index * 4;
    tensorData[index] = rgbData[sourceIndex] / 255;
    tensorData[planeSize + index] = rgbData[sourceIndex + 1] / 255;
    tensorData[planeSize * 2 + index] = rgbData[sourceIndex + 2] / 255;

    if (channels >= 4) {
      tensorData[planeSize * 3 + index] =
        alphaFirstForFourthChannel
          ? alphaValues?.[index] ?? depthValues?.[index] ?? 0
          : depthValues?.[index] ?? alphaValues?.[index] ?? 0;
    }

    if (channels >= 5) {
      tensorData[planeSize * 4 + index] =
        alphaValues?.[index] ?? isnetValues?.[index] ?? 0;
    }
  }

  return new window.ort.Tensor('float32', tensorData, [
    1,
    channels,
    height,
    width,
  ]);
}

async function runPhotoEditorWithoutBgMattingStage(
  model,
  sourceImage,
  depthCanvas,
  isnetCanvas
) {
  const session = await getPhotoEditorSubjectSession(
    model,
    model.modelFiles?.matting
  );
  const channels = getPhotoEditorOrtInputChannels(
    session,
    model.id === 'withoutbg-focus' ? 5 : 4
  );
  const tensor = createPhotoEditorWithoutBgPipelineTensor({
    sourceImage,
    width: 256,
    height: 256,
    depthCanvas,
    isnetCanvas,
    channels,
  });
  const inputName = session.inputNames?.[0] || 'input';
  const outputTensor = await runPhotoEditorOrtSession(session, {
    [inputName]: tensor,
  });

  return createPhotoEditorAlphaCanvasFromTensor(outputTensor, 256, 256);
}

async function runPhotoEditorWithoutBgRefinerStage(
  model,
  sourceImage,
  depthCanvas,
  alphaCanvas
) {
  const sourceSize = getPhotoEditorImageSize(sourceImage);
  const maxSize = 800;
  const scale = Math.min(
    1,
    maxSize / Math.max(1, sourceSize.width, sourceSize.height)
  );
  const refinerSize = {
    width: Math.max(1, Math.round(sourceSize.width * scale)),
    height: Math.max(1, Math.round(sourceSize.height * scale)),
  };
  const session = await getPhotoEditorSubjectSession(
    model,
    model.modelFiles?.refiner
  );
  const inputName = session.inputNames?.[0] || 'input';
  const channels = getPhotoEditorOrtInputChannels(session, 4);
  const tensor = createPhotoEditorWithoutBgPipelineTensor({
    sourceImage,
    width: refinerSize.width,
    height: refinerSize.height,
    depthCanvas,
    alphaCanvas,
    channels,
    alphaFirstForFourthChannel: /rgb_alpha/i.test(inputName) || channels === 4,
  });
  const outputTensor = await runPhotoEditorOrtSession(session, {
    [inputName]: tensor,
  });
  const refinedCanvas = createPhotoEditorAlphaCanvasFromTensor(
    outputTensor,
    refinerSize.width,
    refinerSize.height
  );
  const outputCanvas = drawPhotoEditorImageToCanvas(
    refinedCanvas,
    sourceSize.width,
    sourceSize.height
  );

  refinePhotoEditorSubjectMaskWithSourceImage(outputCanvas, sourceImage);
  return outputCanvas;
}

async function createPhotoEditorWithoutBgSubjectMaskDataUrl(model, sourceImage) {
  const depthCanvas = await runPhotoEditorWithoutBgDepthStage(model, sourceImage);
  const isnetCanvas =
    model.pipeline === 'withoutbg-focus'
      ? await runPhotoEditorWithoutBgIsnetStage(model, sourceImage)
      : null;
  const alphaCanvas = await runPhotoEditorWithoutBgMattingStage(
    model,
    sourceImage,
    depthCanvas,
    isnetCanvas
  );
  const refinedCanvas = await runPhotoEditorWithoutBgRefinerStage(
    model,
    sourceImage,
    depthCanvas,
    alphaCanvas
  );

  return refinedCanvas.toDataURL('image/png');
}

async function runPhotoEditorSubjectModel(model, geometry) {
  if (model.runtime === 'node') {
    if (!window.electronAPI.runAiSubjectModel) {
      throw new Error('標準AIの実行機能を利用できません');
    }

    const result = await window.electronAPI.runAiSubjectModel({
      modelId: model.id,
      inputShape: geometry.tensor.dims,
      inputData: geometry.tensor.data,
    });

    if (!result?.ok) {
      throw new Error(result?.message || 'AIモデルの実行に失敗しました');
    }

    return normalizePhotoEditorSubjectOutputTensor(result);
  }

  const session = await getPhotoEditorSubjectSession(model);
  const inputName = session.inputNames?.[0] || 'input';
  const feeds = {
    [inputName]: geometry.tensor,
  };
  const outputs = await session.run(feeds);
  const outputName = session.outputNames?.[0] || Object.keys(outputs)[0];
  return outputs[outputName];
}

function createPhotoEditorSubjectModelInput(
  image,
  inputSize = 320,
  { keepAspectRatio = true } = {}
) {
  const sourceSize = getPhotoEditorImageSize(image);

  if (sourceSize.width <= 0 || sourceSize.height <= 0) {
    throw new Error('画像サイズを取得できませんでした');
  }

  const scale = keepAspectRatio
    ? Math.min(inputSize / sourceSize.width, inputSize / sourceSize.height)
    : inputSize / Math.max(1, Math.max(sourceSize.width, sourceSize.height));
  const drawWidth = keepAspectRatio
    ? Math.max(1, Math.round(sourceSize.width * scale))
    : inputSize;
  const drawHeight = keepAspectRatio
    ? Math.max(1, Math.round(sourceSize.height * scale))
    : inputSize;
  const offsetX = keepAspectRatio ? Math.floor((inputSize - drawWidth) / 2) : 0;
  const offsetY = keepAspectRatio ? Math.floor((inputSize - drawHeight) / 2) : 0;
  const canvas = document.createElement('canvas');
  canvas.width = inputSize;
  canvas.height = inputSize;
  const ctx = canvas.getContext('2d', { willReadFrequently: true });

  if (!ctx) {
    throw new Error('AI入力画像を作成できませんでした');
  }

  ctx.fillStyle = '#000000';
  ctx.fillRect(0, 0, inputSize, inputSize);
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(image, offsetX, offsetY, drawWidth, drawHeight);

  const imageData = ctx.getImageData(0, 0, inputSize, inputSize);
  const data = imageData.data;
  const floatData = new Float32Array(1 * 3 * inputSize * inputSize);
  const mean = [0.485, 0.456, 0.406];
  const std = [0.229, 0.224, 0.225];
  const planeSize = inputSize * inputSize;

  for (let index = 0; index < planeSize; index += 1) {
    const pixelIndex = index * 4;
    floatData[index] = (data[pixelIndex] / 255 - mean[0]) / std[0];
    floatData[planeSize + index] =
      (data[pixelIndex + 1] / 255 - mean[1]) / std[1];
    floatData[planeSize * 2 + index] =
      (data[pixelIndex + 2] / 255 - mean[2]) / std[2];
  }

  return {
    tensor: new window.ort.Tensor('float32', floatData, [
      1,
      3,
      inputSize,
      inputSize,
    ]),
    inputSize,
    drawWidth,
    drawHeight,
    offsetX,
    offsetY,
    sourceWidth: sourceSize.width,
    sourceHeight: sourceSize.height,
    keepAspectRatio,
  };
}

function getPhotoEditorSubjectMaskHistogramPercentile(histogram, percentile) {
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

function getPhotoEditorSubjectMaskOtsuThreshold(histogram, pixelCount) {
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

function removeSmallPhotoEditorSubjectMaskComponents(
  alphaValues,
  width,
  height,
  threshold
) {
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

function fillSmallPhotoEditorSubjectMaskHoles(
  alphaValues,
  width,
  height,
  threshold
) {
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

function cleanupPhotoEditorSubjectMaskCanvas(maskCanvas) {
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

    const p04 = getPhotoEditorSubjectMaskHistogramPercentile(histogram, 0.04);
    const p35 = getPhotoEditorSubjectMaskHistogramPercentile(histogram, 0.35);
    const p58 = getPhotoEditorSubjectMaskHistogramPercentile(histogram, 0.58);
    const p82 = getPhotoEditorSubjectMaskHistogramPercentile(histogram, 0.82);
    const p96 = getPhotoEditorSubjectMaskHistogramPercentile(histogram, 0.96);

    if (p96 - p04 < 12) {
      return maskCanvas;
    }

    const otsuThreshold = getPhotoEditorSubjectMaskOtsuThreshold(
      histogram,
      pixelCount
    );
    const foregroundCount = alphaValues.reduce(
      (count, value) => count + (value >= otsuThreshold ? 1 : 0),
      0
    );
    const foregroundRatio = foregroundCount / Math.max(1, pixelCount);
    const adjustedThreshold =
      foregroundRatio > 0.68
        ? Math.max(otsuThreshold, p58)
        : clampNumber(otsuThreshold, p35, p82, otsuThreshold);
    const low = clampNumber(adjustedThreshold - 20, p04, 245, adjustedThreshold);
    const high = clampNumber(adjustedThreshold + 34, low + 1, p96, low + 36);
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

    removeSmallPhotoEditorSubjectMaskComponents(
      alphaValues,
      width,
      height,
      componentThreshold
    );
    fillSmallPhotoEditorSubjectMaskHoles(
      alphaValues,
      width,
      height,
      componentThreshold
    );

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

function refinePhotoEditorSubjectMaskWithSourceImage(maskCanvas, sourceImage) {
  const sourceSize = getPhotoEditorImageSize(sourceImage);

  if (
    !maskCanvas ||
    !sourceImage ||
    sourceSize.width !== maskCanvas.width ||
    sourceSize.height !== maskCanvas.height
  ) {
    return maskCanvas;
  }

  const sourceCanvas = document.createElement('canvas');
  sourceCanvas.width = sourceSize.width;
  sourceCanvas.height = sourceSize.height;
  const sourceCtx = sourceCanvas.getContext('2d', { willReadFrequently: true });
  const maskCtx = maskCanvas.getContext('2d', { willReadFrequently: true });

  if (!sourceCtx || !maskCtx) {
    return maskCanvas;
  }

  try {
    sourceCtx.drawImage(sourceImage, 0, 0, sourceSize.width, sourceSize.height);
    const sourceImageData = sourceCtx.getImageData(
      0,
      0,
      sourceSize.width,
      sourceSize.height
    );
    const maskImageData = maskCtx.getImageData(
      0,
      0,
      maskCanvas.width,
      maskCanvas.height
    );
    const sourceData = sourceImageData.data;
    const maskData = maskImageData.data;
    const pixelCount = sourceSize.width * sourceSize.height;
    const originalMask = new Uint8Array(pixelCount);
    const refinedMask = new Uint8Array(pixelCount);

    for (let index = 0; index < pixelCount; index += 1) {
      originalMask[index] = maskData[index * 4];
    }

    refinedMask.set(originalMask);

    const width = sourceSize.width;
    const height = sourceSize.height;
    const getLuma = (pixelIndex) => {
      const dataIndex = pixelIndex * 4;
      return (
        sourceData[dataIndex] * 0.299 +
        sourceData[dataIndex + 1] * 0.587 +
        sourceData[dataIndex + 2] * 0.114
      );
    };

    for (let y = 1; y < height - 1; y += 1) {
      const rowOffset = y * width;

      for (let x = 1; x < width - 1; x += 1) {
        const pixelIndex = rowOffset + x;
        const alpha = originalMask[pixelIndex];

        if (alpha <= 4 || alpha >= 251) {
          continue;
        }

        const sourceIndex = pixelIndex * 4;
        const centerRed = sourceData[sourceIndex];
        const centerGreen = sourceData[sourceIndex + 1];
        const centerBlue = sourceData[sourceIndex + 2];
        let weightedAlpha = 0;
        let weightTotal = 0;

        for (let offsetY = -1; offsetY <= 1; offsetY += 1) {
          for (let offsetX = -1; offsetX <= 1; offsetX += 1) {
            const neighborIndex = pixelIndex + offsetY * width + offsetX;
            const neighborSourceIndex = neighborIndex * 4;
            const redDiff = centerRed - sourceData[neighborSourceIndex];
            const greenDiff = centerGreen - sourceData[neighborSourceIndex + 1];
            const blueDiff = centerBlue - sourceData[neighborSourceIndex + 2];
            const colorDistance =
              redDiff * redDiff + greenDiff * greenDiff + blueDiff * blueDiff;
            const spatialWeight = offsetX === 0 && offsetY === 0 ? 1 : 0.72;
            const colorWeight = 1 / (1 + colorDistance / 1800);
            const weight = spatialWeight * colorWeight;
            weightedAlpha += originalMask[neighborIndex] * weight;
            weightTotal += weight;
          }
        }

        const bilateralAlpha = weightTotal > 0 ? weightedAlpha / weightTotal : alpha;
        const lumaHorizontal =
          Math.abs(getLuma(pixelIndex - 1) - getLuma(pixelIndex + 1));
        const lumaVertical =
          Math.abs(getLuma(pixelIndex - width) - getLuma(pixelIndex + width));
        const edgeStrength = clampNumber(
          (Math.max(lumaHorizontal, lumaVertical) - 10) / 72,
          0,
          1,
          0
        );
        const sharpenedAlpha =
          128 + (bilateralAlpha - 128) * (1 + edgeStrength * 0.9);
        refinedMask[pixelIndex] = clampColorChannel(
          alpha * 0.28 + sharpenedAlpha * 0.72
        );
      }
    }

    for (let index = 0; index < pixelCount; index += 1) {
      const value = refinedMask[index];
      const dataIndex = index * 4;
      maskData[dataIndex] = value;
      maskData[dataIndex + 1] = value;
      maskData[dataIndex + 2] = value;
      maskData[dataIndex + 3] = 255;
    }

    maskCtx.putImageData(maskImageData, 0, 0);
  } catch {
    return maskCanvas;
  }

  return maskCanvas;
}

function createPhotoEditorSubjectMaskDataUrlFromTensor(
  outputTensor,
  geometry,
  sourceImage = null
) {
  const inputSize = geometry.inputSize;
  const outputData = outputTensor?.data;

  if (!outputData || outputData.length === 0) {
    throw new Error('AIマスクの出力が空でした');
  }

  const maskCanvas = document.createElement('canvas');
  maskCanvas.width = inputSize;
  maskCanvas.height = inputSize;
  const maskCtx = maskCanvas.getContext('2d', { willReadFrequently: true });

  if (!maskCtx) {
    throw new Error('AIマスクを作成できませんでした');
  }

  const imageData = maskCtx.createImageData(inputSize, inputSize);
  const data = imageData.data;
  let minValue = Infinity;
  let maxValue = -Infinity;

  for (let index = 0; index < outputData.length; index += 1) {
    const value = Number(outputData[index]);
    if (Number.isFinite(value)) {
      minValue = Math.min(minValue, value);
      maxValue = Math.max(maxValue, value);
    }
  }

  const range = maxValue > minValue ? maxValue - minValue : 1;
  const pixelCount = inputSize * inputSize;

  for (let index = 0; index < pixelCount; index += 1) {
    const rawValue = Number(outputData[index] || 0);
    const normalizedValue =
      maxValue > 1 || minValue < 0
        ? (rawValue - minValue) / range
        : rawValue;
    const value = clampColorChannel(normalizedValue * 255);
    const pixelIndex = index * 4;
    data[pixelIndex] = value;
    data[pixelIndex + 1] = value;
    data[pixelIndex + 2] = value;
    data[pixelIndex + 3] = 255;
  }

  maskCtx.putImageData(imageData, 0, 0);
  cleanupPhotoEditorSubjectMaskCanvas(maskCanvas);

  const cropCanvas = document.createElement('canvas');
  cropCanvas.width = geometry.drawWidth;
  cropCanvas.height = geometry.drawHeight;
  const cropCtx = cropCanvas.getContext('2d');

  if (!cropCtx) {
    throw new Error('AIマスクを切り出せませんでした');
  }

  cropCtx.drawImage(
    maskCanvas,
    geometry.offsetX,
    geometry.offsetY,
    geometry.drawWidth,
    geometry.drawHeight,
    0,
    0,
    geometry.drawWidth,
    geometry.drawHeight
  );

  const outputCanvas = document.createElement('canvas');
  outputCanvas.width = geometry.sourceWidth;
  outputCanvas.height = geometry.sourceHeight;
  const outputCtx = outputCanvas.getContext('2d');

  if (!outputCtx) {
    throw new Error('AIマスクを元画像サイズに戻せませんでした');
  }

  outputCtx.imageSmoothingEnabled = true;
  outputCtx.imageSmoothingQuality = 'high';
  outputCtx.drawImage(
    cropCanvas,
    0,
    0,
    geometry.drawWidth,
    geometry.drawHeight,
    0,
    0,
    geometry.sourceWidth,
    geometry.sourceHeight
  );

  refinePhotoEditorSubjectMaskWithSourceImage(outputCanvas, sourceImage);

  return outputCanvas.toDataURL('image/png');
}

async function generatePhotoEditorSubjectMaskWithModel(modelId, source) {
  if (!photoEditorState?.sourceImage) {
    return;
  }

  let model = getPhotoEditorSubjectModel(modelId);

  if (!model) {
    await refreshPhotoEditorAiSubjectModelStatus();
    model = getPhotoEditorSubjectModel(modelId);
  }

  const label = model?.label || 'AI';

  if (!model?.ready) {
    setPhotoEditorStatus(model?.message || `${label}モデルを利用できません`);
    syncPhotoEditorSubjectMaskControls();
    return;
  }

  const startedAt = performance.now();
  updatePhotoEditorSubjectMask(
    {
      enabled: false,
      status: 'loading',
      source,
      modelId: model.id,
      maskDataUrl: '',
      errorMessage: '',
    },
    { interactive: false }
  );
  setPhotoEditorStatus(`${label}で被写体を選択中...`);

  try {
    let maskDataUrl = '';

    if (model.pipeline === 'withoutbg-snap' || model.pipeline === 'withoutbg-focus') {
      maskDataUrl = await createPhotoEditorWithoutBgSubjectMaskDataUrl(
        model,
        photoEditorState.sourceImage
      );
    } else {
      const geometry = createPhotoEditorSubjectModelInput(
        photoEditorState.sourceImage,
        model.inputSize || 320,
        { keepAspectRatio: model.keepAspectRatio !== false }
      );
      const outputTensor = await runPhotoEditorSubjectModel(model, geometry);
      setPhotoEditorStatus('高解像度マスクを整えています...');
      maskDataUrl = createPhotoEditorSubjectMaskDataUrlFromTensor(
        outputTensor,
        geometry,
        photoEditorState.sourceImage
      );
    }
    const timestamp = new Date().toISOString();
    const imageSize = getPhotoEditorImageSize(photoEditorState.sourceImage);

    beginPhotoEditorHistoryMutation();
    photoEditorState.subjectMask = normalizePhotoEditorSubjectMaskState({
      enabled: true,
      status: 'ready',
      source,
      modelId: model.id,
      maskDataUrl,
      width: imageSize.width,
      height: imageSize.height,
      showOverlay: true,
      createdAt: timestamp,
      updatedAt: timestamp,
    });
    syncPhotoEditorSubjectMaskControls();
    schedulePhotoEditorRender();
    commitPhotoEditorHistoryMutation();
    setPhotoEditorStatus(
      `${label}で被写体を選択しました (${Math.round(performance.now() - startedAt)}ms)`
    );
  } catch (error) {
    beginPhotoEditorHistoryMutation();
    photoEditorState.subjectMask = normalizePhotoEditorSubjectMaskState({
      enabled: false,
      status: 'failed',
      source,
      modelId: model.id,
      errorMessage: `${label}の実行に失敗しました: ${error.message}`,
    });
    syncPhotoEditorSubjectMaskControls();
    schedulePhotoEditorRender();
    commitPhotoEditorHistoryMutation();
    setPhotoEditorStatus(`${label}の実行に失敗しました: ${error.message}`);
  }
}

async function generatePhotoEditorLightweightSubjectMask() {
  await generatePhotoEditorSubjectMaskWithModel('u2netp', 'lightweight');
}

async function generatePhotoEditorAutoSubjectMask() {
  const model = getPhotoEditorAutoSubjectModel();
  await generatePhotoEditorSubjectMaskWithModel(
    model?.id || 'u2netp',
    model?.tier === 'standard' ? 'standard' : 'lightweight'
  );
}

async function generatePhotoEditorHighQualitySubjectMask() {
  await generatePhotoEditorSubjectMaskWithModel('withoutbg-focus', 'high-quality');
}

function loadPhotoEditorSubjectMaskFile(file) {
  if (!file || !photoEditorState) {
    return;
  }

  if (!/^image\/(?:png|jpeg|webp)$/i.test(file.type || '')) {
    setPhotoEditorStatus('PNG / JPEG / WebP のマスク画像を選択してください');
    return;
  }

  const reader = new FileReader();

  reader.onload = () => {
    const dataUrl = String(reader.result || '');

    if (!dataUrl.startsWith('data:image/')) {
      setPhotoEditorStatus('マスク画像を読み込めませんでした');
      return;
    }

    const image = new Image();
    image.onload = () => {
      if (!photoEditorState) {
        return;
      }

      const timestamp = new Date().toISOString();
      beginPhotoEditorHistoryMutation();
      photoEditorState.subjectMask = normalizePhotoEditorSubjectMaskState({
        enabled: true,
        status: 'ready',
        source: 'imported',
        maskDataUrl: dataUrl,
        width: Number(image.naturalWidth || image.width) || 0,
        height: Number(image.naturalHeight || image.height) || 0,
        feather: 0,
        expand: 0,
        opacity: 0.55,
        showOverlay: true,
        createdAt: timestamp,
        updatedAt: timestamp,
      });
      syncPhotoEditorSubjectMaskControls();
      schedulePhotoEditorRender();
      commitPhotoEditorHistoryMutation();
    };
    image.onerror = () => {
      setPhotoEditorStatus('マスク画像を読み込めませんでした');
    };
    image.src = dataUrl;
  };
  reader.onerror = () => {
    setPhotoEditorStatus('マスク画像を読み込めませんでした');
  };
  reader.readAsDataURL(file);
}

function syncPhotoEditorAutoEnhanceControls() {
  if (!photoEditorState) {
    return;
  }

  photoEditorState.autoEnhance = normalizePhotoEditorAutoEnhanceState(
    photoEditorState.autoEnhance
  );
  const strength = photoEditorState.autoEnhance.strength;

  if (photoEditorAutoStrengthInput) {
    photoEditorAutoStrengthInput.value = String(strength);
  }

  if (photoEditorAutoStrengthValue) {
    photoEditorAutoStrengthValue.textContent = String(strength);
  }
}

function syncPhotoEditorCropControls() {
  if (!photoEditorState) {
    return;
  }

  photoEditorState.crop = normalizePhotoEditorCropState(photoEditorState.crop);
  const { crop } = photoEditorState;

  if (photoEditorCropZoomInput) {
    photoEditorCropZoomInput.value = String(crop.zoom);
  }

  if (photoEditorCropZoomValue) {
    photoEditorCropZoomValue.textContent = String(crop.zoom);
  }

  if (photoEditorCropXInput) {
    photoEditorCropXInput.value = String(crop.offsetX);
  }

  if (photoEditorCropXValue) {
    photoEditorCropXValue.textContent = String(crop.offsetX);
  }

  if (photoEditorCropYInput) {
    photoEditorCropYInput.value = String(crop.offsetY);
  }

  if (photoEditorCropYValue) {
    photoEditorCropYValue.textContent = String(crop.offsetY);
  }

  if (photoEditorCropTiltInput) {
    photoEditorCropTiltInput.value = String(crop.tilt);
  }

  if (photoEditorCropTiltValue) {
    photoEditorCropTiltValue.textContent = `${crop.tilt}°`;
  }

  if (photoEditorCropFlipXButton) {
    photoEditorCropFlipXButton.setAttribute(
      'aria-pressed',
      crop.flipX ? 'true' : 'false'
    );
  }

  if (photoEditorCropFlipYButton) {
    photoEditorCropFlipYButton.setAttribute(
      'aria-pressed',
      crop.flipY ? 'true' : 'false'
    );
  }

  for (const button of photoEditorCropPresetList?.querySelectorAll(
    '[data-photo-editor-crop]'
  ) || []) {
    const isActive = button.dataset.photoEditorCrop === crop.preset;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  }
}

function syncPhotoEditorBlurControls() {
  if (!photoEditorState) {
    return;
  }

  const blur = normalizePhotoEditorBlurState(photoEditorState.blur);
  photoEditorState.blur = blur;

  for (const button of photoEditorBlurModeButtons) {
    const isActive = button.dataset.photoEditorBlurMode === blur.mode;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  }

  if (photoEditorBlurAmountInput) {
    photoEditorBlurAmountInput.value = String(blur.amount);
  }

  if (photoEditorBlurAmountValue) {
    photoEditorBlurAmountValue.textContent = String(blur.amount);
  }

  if (photoEditorBlurConfirmButton) {
    photoEditorBlurConfirmButton.hidden = blur.mode !== 'radial';
  }

  if (photoEditorBlurConfirmLabel) {
    photoEditorBlurConfirmLabel.textContent = blur.isConfirmed
      ? '範囲を編集'
      : '確定';
  }

  const blurConfirmIcon = photoEditorBlurConfirmButton?.querySelector(
    '.material-symbols-outlined'
  );

  if (blurConfirmIcon) {
    blurConfirmIcon.textContent = blur.isConfirmed ? 'edit' : 'check';
  }

  photoEditorCanvas?.classList.toggle(
    'is-blur-tool-active',
    blur.mode === 'radial' &&
      !blur.isConfirmed &&
      isPhotoEditorAccordionOpen('blur') &&
      photoEditorState.maskTool === 'none' &&
      !photoEditorState.draftMask
  );
}

function getPhotoEditorActiveCurvePoints() {
  const curve = normalizePhotoEditorCurveState(photoEditorState?.curve);
  return curve.points[curve.mode][curve.channel];
}

function safelySetPointerCapture(element, pointerId) {
  if (!element || pointerId === undefined || pointerId === null) {
    return;
  }

  try {
    element.setPointerCapture?.(pointerId);
  } catch {
    // Synthetic input and some pointer devices can report no active pointer.
  }
}

function getPhotoEditorCurveCanvasMetrics() {
  const bounds = photoEditorCurveCanvas?.getBoundingClientRect();
  const cssWidth = Math.max(1, Math.round(bounds?.width || 280));
  const cssHeight = Math.max(1, Math.round(bounds?.height || 160));
  const deviceScale = Math.max(1, Math.min(2, window.devicePixelRatio || 1));

  return {
    cssWidth,
    cssHeight,
    width: Math.round(cssWidth * deviceScale),
    height: Math.round(cssHeight * deviceScale),
    scale: deviceScale,
    padding: Math.round(14 * deviceScale),
  };
}

function createPhotoEditorHistogramBins() {
  return Array.from({ length: PHOTO_EDITOR_CURVE_HISTOGRAM_BINS }, () => 0);
}

function getPhotoEditorCurveHistogramKey(sourceRect, crop, imageSize) {
  return JSON.stringify({
    imageWidth: imageSize.width,
    imageHeight: imageSize.height,
    sourceRect: {
      x: Math.round(sourceRect.x),
      y: Math.round(sourceRect.y),
      width: Math.round(sourceRect.width),
      height: Math.round(sourceRect.height),
    },
    crop: {
      preset: crop.preset,
      zoom: Math.round(crop.zoom * 10) / 10,
      offsetX: Math.round(crop.offsetX * 10) / 10,
      offsetY: Math.round(crop.offsetY * 10) / 10,
      rotation: crop.rotation,
      flipX: crop.flipX,
      flipY: crop.flipY,
      tilt: Math.round(crop.tilt * 10) / 10,
    },
  });
}

function normalizePhotoEditorHistogramBins(histogram) {
  const maxValue = Math.max(1, ...histogram);
  return histogram.map((value) => value / maxValue);
}

function getPhotoEditorCurveHistogram() {
  if (!photoEditorState?.sourceImage) {
    return null;
  }

  const crop = normalizePhotoEditorCropState(photoEditorState.crop);
  const sourceRect = getPhotoEditorSourceRect(photoEditorState.sourceImage);
  const imageSize = getPhotoEditorImageSize(photoEditorState.sourceImage);
  const histogramKey = getPhotoEditorCurveHistogramKey(sourceRect, crop, imageSize);

  if (
    photoEditorState.curveHistogram &&
    photoEditorState.curveHistogramKey === histogramKey
  ) {
    return photoEditorState.curveHistogram;
  }

  const sampleSize = getPhotoEditorOutputSize(sourceRect, 160, crop);
  const sampleCanvas = document.createElement('canvas');
  sampleCanvas.width = sampleSize.width;
  sampleCanvas.height = sampleSize.height;
  const sampleCtx = sampleCanvas.getContext('2d', { willReadFrequently: true });

  if (!sampleCtx) {
    return null;
  }

  sampleCtx.imageSmoothingEnabled = true;
  sampleCtx.imageSmoothingQuality = 'high';
  drawPhotoEditorCroppedSourceToCanvas(
    sampleCtx,
    photoEditorState.sourceImage,
    sourceRect,
    sampleSize,
    crop
  );

  let imageData;

  try {
    imageData = sampleCtx.getImageData(0, 0, sampleSize.width, sampleSize.height);
  } catch {
    const previewCtx = photoEditorCanvas?.getContext('2d', {
      willReadFrequently: true,
    });

    if (!previewCtx || !photoEditorCanvas.width || !photoEditorCanvas.height) {
      return null;
    }

    try {
      imageData = previewCtx.getImageData(
        0,
        0,
        photoEditorCanvas.width,
        photoEditorCanvas.height
      );
    } catch {
      return null;
    }
  }

  const bins = {
    rgb: {
      master: createPhotoEditorHistogramBins(),
      r: createPhotoEditorHistogramBins(),
      g: createPhotoEditorHistogramBins(),
      b: createPhotoEditorHistogramBins(),
    },
    hsv: {
      master: createPhotoEditorHistogramBins(),
      h: createPhotoEditorHistogramBins(),
      s: createPhotoEditorHistogramBins(),
      v: createPhotoEditorHistogramBins(),
    },
  };
  const data = imageData.data;
  const getBinIndex = (value) =>
    clampNumber(
      Math.floor(clampNumber(value, 0, 1, 0) * (PHOTO_EDITOR_CURVE_HISTOGRAM_BINS - 1)),
      0,
      PHOTO_EDITOR_CURVE_HISTOGRAM_BINS - 1,
      0
    );

  for (let index = 0; index < data.length; index += 4) {
    const red = data[index];
    const green = data[index + 1];
    const blue = data[index + 2];
    const luminance = (0.2126 * red + 0.7152 * green + 0.0722 * blue) / 255;
    const hsv = convertRgbToHsv(red, green, blue);

    bins.rgb.master[getBinIndex(luminance)] += 1;
    bins.rgb.r[getBinIndex(red / 255)] += 1;
    bins.rgb.g[getBinIndex(green / 255)] += 1;
    bins.rgb.b[getBinIndex(blue / 255)] += 1;
    bins.hsv.master[getBinIndex(hsv.v)] += 1;
    bins.hsv.h[getBinIndex(hsv.h)] += 1;
    bins.hsv.s[getBinIndex(hsv.s)] += 1;
    bins.hsv.v[getBinIndex(hsv.v)] += 1;
  }

  const histogram = {
    rgb: {
      master: normalizePhotoEditorHistogramBins(bins.rgb.master),
      r: normalizePhotoEditorHistogramBins(bins.rgb.r),
      g: normalizePhotoEditorHistogramBins(bins.rgb.g),
      b: normalizePhotoEditorHistogramBins(bins.rgb.b),
    },
    hsv: {
      master: normalizePhotoEditorHistogramBins(bins.hsv.master),
      h: normalizePhotoEditorHistogramBins(bins.hsv.h),
      s: normalizePhotoEditorHistogramBins(bins.hsv.s),
      v: normalizePhotoEditorHistogramBins(bins.hsv.v),
    },
  };

  photoEditorState.curveHistogram = histogram;
  photoEditorState.curveHistogramKey = histogramKey;
  return histogram;
}

function getPhotoEditorCurveHistogramColor(curve) {
  if (curve.mode === 'rgb') {
    return {
      master: 'rgba(148, 163, 184, 0.3)',
      r: 'rgba(248, 113, 113, 0.34)',
      g: 'rgba(74, 222, 128, 0.32)',
      b: 'rgba(96, 165, 250, 0.34)',
    }[curve.channel] || 'rgba(148, 163, 184, 0.3)';
  }

  return {
    master: 'rgba(148, 163, 184, 0.3)',
    h: 'rgba(244, 114, 182, 0.32)',
    s: 'rgba(251, 191, 36, 0.3)',
    v: 'rgba(125, 211, 252, 0.32)',
  }[curve.channel] || 'rgba(148, 163, 184, 0.3)';
}

function drawPhotoEditorCurveHistogram(ctx, curve, graphRect) {
  const histogram = getPhotoEditorCurveHistogram();
  const channelHistogram = histogram?.[curve.mode]?.[curve.channel];

  if (!Array.isArray(channelHistogram) || channelHistogram.length === 0) {
    return;
  }

  const { left, top, width, height } = graphRect;
  const barWidth = width / channelHistogram.length;

  ctx.save();
  ctx.fillStyle = getPhotoEditorCurveHistogramColor(curve);
  channelHistogram.forEach((value, index) => {
    const barHeight = Math.max(1, height * Math.pow(value, 0.72));
    ctx.fillRect(
      left + index * barWidth,
      top + height - barHeight,
      Math.max(1, barWidth + 0.5),
      barHeight
    );
  });
  ctx.restore();
}

function drawPhotoEditorCurveCanvas() {
  if (!photoEditorCurveCanvas || !photoEditorState) {
    return;
  }

  const ctx = photoEditorCurveCanvas.getContext('2d');

  if (!ctx) {
    return;
  }

  const metrics = getPhotoEditorCurveCanvasMetrics();
  const { width, height, padding } = metrics;

  if (
    photoEditorCurveCanvas.width !== width ||
    photoEditorCurveCanvas.height !== height
  ) {
    photoEditorCurveCanvas.width = width;
    photoEditorCurveCanvas.height = height;
  }

  const curve = normalizePhotoEditorCurveState(photoEditorState.curve);
  const points = curve.points[curve.mode][curve.channel];
  const left = padding;
  const top = padding;
  const graphWidth = Math.max(1, width - padding * 2);
  const graphHeight = Math.max(1, height - padding * 2);
  const maxPointIndex = Math.max(1, points.length - 1);
  const pointToCanvas = (point, index) => ({
    x: left + graphWidth * (index / maxPointIndex),
    y: top + graphHeight * (1 - point),
  });
  const canvasPoints = points.map(pointToCanvas);

  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = getComputedStyle(photoEditorCurveCanvas).getPropertyValue(
    '--surface-card'
  ) || '#111827';
  ctx.fillRect(0, 0, width, height);

  drawPhotoEditorCurveHistogram(ctx, curve, {
    left,
    top,
    width: graphWidth,
    height: graphHeight,
  });

  ctx.strokeStyle = 'rgba(148, 163, 184, 0.2)';
  ctx.lineWidth = Math.max(1, Math.round(metrics.scale));

  for (let index = 0; index <= 4; index += 1) {
    const x = left + graphWidth * (index / 4);
    const y = top + graphHeight * (index / 4);
    ctx.beginPath();
    ctx.moveTo(x, top);
    ctx.lineTo(x, top + graphHeight);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(left, y);
    ctx.lineTo(left + graphWidth, y);
    ctx.stroke();
  }

  ctx.strokeStyle = 'rgba(79, 140, 255, 0.95)';
  ctx.lineWidth = Math.max(2, Math.round(2 * metrics.scale));
  ctx.beginPath();
  canvasPoints.forEach((point, index) => {
    if (index === 0) {
      ctx.moveTo(point.x, point.y);
    } else {
      ctx.lineTo(point.x, point.y);
    }
  });
  ctx.stroke();

  for (const point of canvasPoints) {
    ctx.fillStyle = 'rgba(79, 140, 255, 0.95)';
    ctx.strokeStyle = 'rgba(255, 255, 255, 0.95)';
    ctx.lineWidth = Math.max(2, Math.round(metrics.scale));
    ctx.beginPath();
    ctx.arc(point.x, point.y, Math.max(5, 5 * metrics.scale), 0, Math.PI * 2);
    ctx.fill();
    ctx.stroke();
  }
}

function getPhotoEditorCurveCanvasPoint(event) {
  if (!photoEditorCurveCanvas) {
    return null;
  }

  const bounds = photoEditorCurveCanvas.getBoundingClientRect();

  if (!bounds.width || !bounds.height) {
    return null;
  }

  return {
    x: clampNumber((event.clientX - bounds.left) / bounds.width, 0, 1, 0),
    y: clampNumber((event.clientY - bounds.top) / bounds.height, 0, 1, 0),
  };
}

function updatePhotoEditorCurvePoint(pointIndex, pointerPoint) {
  const maxPointIndex = PHOTO_EDITOR_CURVE_DEFAULT_POINTS.length - 1;

  if (!photoEditorState || pointIndex < 0 || pointIndex > maxPointIndex) {
    return;
  }

  const curve = normalizePhotoEditorCurveState(photoEditorState.curve);
  const points = clonePhotoEditorCurvePoints(
    curve.points[curve.mode][curve.channel]
  );
  const nextValue = clampNumber(1 - pointerPoint.y, 0, 1, points[pointIndex]);

  if (Math.abs(points[pointIndex] - nextValue) < 0.002) {
    return;
  }

  points[pointIndex] = nextValue;
  curve.points[curve.mode][curve.channel] = points;
  photoEditorState.curve = curve;
  drawPhotoEditorCurveCanvas();
  schedulePhotoEditorRender({
    debounceMs: PHOTO_EDITOR_CURVE_PREVIEW_DEBOUNCE_MS,
    interactive: true,
  });
}

function beginPhotoEditorCurveDrag(event) {
  if (!photoEditorState) {
    return;
  }

  const point = getPhotoEditorCurveCanvasPoint(event);

  if (!point) {
    return;
  }

  event.preventDefault();
  safelySetPointerCapture(photoEditorCurveCanvas, event.pointerId);
  beginPhotoEditorHistoryMutation();
  photoEditorState.dragMode = 'curve';
  photoEditorState.dragCurvePointIndex = clampNumber(
    Math.round(point.x * (PHOTO_EDITOR_CURVE_DEFAULT_POINTS.length - 1)),
    0,
    PHOTO_EDITOR_CURVE_DEFAULT_POINTS.length - 1,
    2
  );
  updatePhotoEditorCurvePoint(photoEditorState.dragCurvePointIndex, point);
}

function updatePhotoEditorCurveDrag(event) {
  if (
    !photoEditorState ||
    photoEditorState.dragMode !== 'curve' ||
    photoEditorState.dragCurvePointIndex === null
  ) {
    return;
  }

  const point = getPhotoEditorCurveCanvasPoint(event);

  if (!point) {
    return;
  }

  event.preventDefault();
  updatePhotoEditorCurvePoint(photoEditorState.dragCurvePointIndex, point);
}

function finishPhotoEditorCurveDrag() {
  if (!photoEditorState || photoEditorState.dragMode !== 'curve') {
    return;
  }

  photoEditorState.dragMode = null;
  photoEditorState.dragCurvePointIndex = null;
  syncPhotoEditorCurveControls();
  finishPhotoEditorInteractivePreview();
  commitPhotoEditorHistoryMutation();
}

function syncPhotoEditorCurveControls() {
  if (!photoEditorState) {
    return;
  }

  const curve = normalizePhotoEditorCurveState(photoEditorState.curve);
  photoEditorState.curve = curve;

  for (const button of photoEditorCurveModeButtons) {
    const isActive = button.dataset.photoEditorCurveMode === curve.mode;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  }

  const channelButtons = Array.from(
    photoEditorCurveChannelList?.querySelectorAll(
      '[data-photo-editor-curve-channel]'
    ) || []
  );
  const channels = PHOTO_EDITOR_CURVE_CHANNELS[curve.mode];

  channelButtons.forEach((button, index) => {
    const channel = channels[index];

    if (!channel) {
      button.hidden = true;
      return;
    }

    const isActive = channel.key === curve.channel;
    button.hidden = false;
    button.dataset.photoEditorCurveChannel = channel.key;
    button.textContent = channel.label;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  });

  drawPhotoEditorCurveCanvas();
}

function syncPhotoEditorMaskToolUi() {
  if (!photoEditorState) {
    return;
  }

  const pendingMask = photoEditorState.draftMask;
  const activeMaskType = pendingMask?.type || photoEditorState.maskTool;
  const isMaskDragging = String(photoEditorState.dragMode || '').startsWith('mask');
  const hasPendingMask =
    Boolean(pendingMask) && !isMaskDragging;

  for (const button of photoEditorMaskToolButtons) {
    const isActive =
      button.dataset.photoEditorMaskTool === activeMaskType;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  }

  photoEditorCanvas?.classList.toggle(
    'is-mask-tool-active',
    photoEditorState.maskTool !== 'none'
  );
  photoEditorCanvas?.classList.toggle(
    'is-blur-tool-active',
    photoEditorState.blur?.mode === 'radial' &&
      !photoEditorState.blur?.isConfirmed &&
      isPhotoEditorAccordionOpen('blur') &&
      photoEditorState.maskTool === 'none' &&
      !pendingMask
  );
  photoEditorCanvas?.classList.toggle(
    'is-mask-draft-active',
    Boolean(pendingMask)
  );

  if (photoEditorMaskOptions) {
    photoEditorMaskOptions.hidden =
      photoEditorState.maskTool === 'none' && !pendingMask;
  }

  for (const button of photoEditorMaskShapeButtons) {
    const isActive =
      button.dataset.photoEditorMaskShape === photoEditorState.maskShape;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  }

  const hasMaskStrength = ['blur', 'mosaic', 'fill'].includes(activeMaskType);

  if (photoEditorMaskBlurStrengthRow) {
    photoEditorMaskBlurStrengthRow.hidden = !hasMaskStrength;
  }

  if (photoEditorMaskStrengthLabel) {
    photoEditorMaskStrengthLabel.textContent =
      PHOTO_EDITOR_MASK_STRENGTH_LABELS[activeMaskType] || '濃さ';
  }

  if (photoEditorMaskBlurStrengthInput) {
    photoEditorMaskBlurStrengthInput.value = String(photoEditorState.blurStrength);
  }

  if (photoEditorMaskBlurStrengthValue) {
    photoEditorMaskBlurStrengthValue.textContent = String(
      photoEditorState.blurStrength
    );
  }

  if (photoEditorFillColorRow) {
    photoEditorFillColorRow.hidden = activeMaskType !== 'fill';
  }

  if (photoEditorMaskConfirmButton) {
    photoEditorMaskConfirmButton.disabled = !hasPendingMask;
  }

  const hasMasksOrPending = photoEditorState.masks.length > 0 || hasPendingMask;

  if (photoEditorMaskUndoButton) {
    photoEditorMaskUndoButton.disabled = !hasMasksOrPending;
  }

  if (photoEditorMaskClearButton) {
    photoEditorMaskClearButton.disabled = !hasMasksOrPending;
  }
}

function renderPhotoEditorTextFontOptions() {
  if (!photoEditorTextFontSelect) {
    return;
  }

  const selectedKey =
    getPhotoEditorActiveTextOverlay()?.fontKey ||
    getPhotoEditorTextFontOption('system').key;
  const currentValue = photoEditorTextFontSelect.value || selectedKey;
  photoEditorTextFontSelect.innerHTML = '';
  const recentKeys = photoEditorTextRecentFonts.filter(
    (fontKey) => getPhotoEditorTextFontOption(fontKey).key === fontKey
  );
  const appendedKeys = new Set();

  if (recentKeys.length > 0) {
    const recentGroup = document.createElement('optgroup');
    recentGroup.label = '最近使用';

    for (const fontKey of recentKeys) {
      const font = getPhotoEditorTextFontOption(fontKey);
      const option = createPhotoEditorTextFontOption(font);
      recentGroup.appendChild(option);
      appendedKeys.add(font.key);
    }

    photoEditorTextFontSelect.appendChild(recentGroup);
  }

  const fontGroup = document.createElement('optgroup');
  fontGroup.label = 'フォント';
  const sortedFonts = [...PHOTO_EDITOR_TEXT_FONT_OPTIONS].sort((left, right) =>
    left.label.localeCompare(right.label, 'en', { sensitivity: 'base' })
  );

  for (const font of sortedFonts) {
    if (appendedKeys.has(font.key)) {
      continue;
    }

    const option = createPhotoEditorTextFontOption(font);
    fontGroup.appendChild(option);
  }

  photoEditorTextFontSelect.appendChild(fontGroup);
  photoEditorTextFontSelect.value = getPhotoEditorTextFontOption(currentValue).key;
  syncPhotoEditorTextFontSelectPreview(photoEditorTextFontSelect.value);
}

function formatPhotoEditorTextWeightLabel(weight) {
  const weightLabelMap = {
    400: '標準',
    500: '中太',
    600: 'セミボールド',
    700: '太字',
    800: '特太',
    900: '極太',
  };

  return weightLabelMap[weight] || `${weight}`;
}

function renderPhotoEditorTextWeightOptions(fontKey, selectedWeight) {
  if (!photoEditorTextWeightSelect) {
    return;
  }

  const weights = getPhotoEditorTextFontWeights(fontKey);
  const normalizedWeight = getClosestPhotoEditorTextWeight(
    fontKey,
    selectedWeight
  );
  photoEditorTextWeightSelect.innerHTML = '';

  for (const weight of weights) {
    const option = document.createElement('option');
    option.value = weight;
    option.textContent = formatPhotoEditorTextWeightLabel(weight);
    photoEditorTextWeightSelect.appendChild(option);
  }

  photoEditorTextWeightSelect.value = normalizedWeight;
}

function getPhotoEditorTextListLabel(textOverlay, index) {
  const text = String(textOverlay?.text || '').trim();
  return text
    ? text.slice(0, 24)
    : `${translateUiText('テキスト')} ${index + 1}`;
}

function renderPhotoEditorTextList(textOverlays, activeTextId) {
  if (!photoEditorTextList) {
    return;
  }

  photoEditorTextList.innerHTML = '';

  textOverlays.forEach((textOverlay, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'photo-editor-text-item';
    button.dataset.photoEditorTextId = textOverlay.id;
    button.textContent = getPhotoEditorTextListLabel(textOverlay, index);
    button.classList.toggle('is-active', textOverlay.id === activeTextId);
    button.setAttribute(
      'aria-pressed',
      textOverlay.id === activeTextId ? 'true' : 'false'
    );
    button.addEventListener('click', () => {
      selectPhotoEditorTextOverlay(textOverlay.id);
    });
    photoEditorTextList.appendChild(button);
  });
}

function setPhotoEditorTextControlsDisabled(isDisabled) {
  [
    photoEditorTextContentInput,
    photoEditorTextFontSelect,
    photoEditorTextSizeInput,
    photoEditorTextColorInput,
    photoEditorTextWeightSelect,
    photoEditorTextStrokeTypeSelect,
    photoEditorTextStrokeWidthInput,
    photoEditorTextStrokeColorInput,
    photoEditorTextFillTransparentInput,
    photoEditorTextMaskModeSelect,
    photoEditorTextLetterSpacingInput,
  ]
    .filter(Boolean)
    .forEach((control) => {
      control.disabled = Boolean(isDisabled);
    });
}

function syncPhotoEditorTextControls() {
  if (!photoEditorState) {
    return;
  }

  const collection = getPhotoEditorTextCollectionFromState();
  setPhotoEditorTextCollection(collection.textOverlays, collection.activeTextId);
  const textOverlay = getPhotoEditorActiveTextOverlay();
  const displayState = textOverlay || getDefaultPhotoEditorTextState();

  renderPhotoEditorTextList(collection.textOverlays, collection.activeTextId);

  if (
    photoEditorTextFontSelect &&
    !Array.from(photoEditorTextFontSelect.options).some(
      (option) => option.value === displayState.fontKey
    )
  ) {
    renderPhotoEditorTextFontOptions();
  }

  if (
    photoEditorTextContentInput &&
    document.activeElement !== photoEditorTextContentInput
  ) {
    photoEditorTextContentInput.value = displayState.text;
  }

  if (photoEditorTextFontSelect) {
    photoEditorTextFontSelect.value = displayState.fontKey;
    syncPhotoEditorTextFontSelectPreview(displayState.fontKey);
  }

  if (photoEditorTextSizeInput) {
    photoEditorTextSizeInput.value = String(displayState.size);
  }

  if (photoEditorTextSizeValue) {
    photoEditorTextSizeValue.textContent = String(Math.round(displayState.size));
  }

  if (photoEditorTextColorInput) {
    photoEditorTextColorInput.value = displayState.color;
  }

  renderPhotoEditorTextWeightOptions(displayState.fontKey, displayState.weight);

  if (photoEditorTextStrokeTypeSelect) {
    photoEditorTextStrokeTypeSelect.value = displayState.strokeType;
  }

  if (photoEditorTextStrokeWidthInput) {
    photoEditorTextStrokeWidthInput.value = String(displayState.strokeWidth);
  }

  if (photoEditorTextStrokeWidthValue) {
    photoEditorTextStrokeWidthValue.textContent = String(
      Math.round(displayState.strokeWidth)
    );
  }

  if (photoEditorTextStrokeColorInput) {
    photoEditorTextStrokeColorInput.value = displayState.strokeColor;
  }

  if (photoEditorTextFillTransparentInput) {
    photoEditorTextFillTransparentInput.classList.toggle(
      'is-active',
      Boolean(displayState.fillTransparent)
    );
    photoEditorTextFillTransparentInput.setAttribute(
      'aria-pressed',
      String(Boolean(displayState.fillTransparent))
    );
  }

  if (photoEditorTextMaskModeSelect) {
    photoEditorTextMaskModeSelect.value = displayState.maskMode;
  }

  if (photoEditorTextLetterSpacingInput) {
    photoEditorTextLetterSpacingInput.value = String(displayState.letterSpacing);
  }

  if (photoEditorTextLetterSpacingValue) {
    photoEditorTextLetterSpacingValue.textContent = String(
      Math.round(displayState.letterSpacing)
    );
  }

  if (photoEditorTextDeleteButton) {
    photoEditorTextDeleteButton.disabled = !textOverlay;
  }

  if (photoEditorTextControls) {
    photoEditorTextControls.hidden = !textOverlay;
  }

  setPhotoEditorTextControlsDisabled(!textOverlay);

  photoEditorCanvas?.classList.toggle(
    'is-text-tool-active',
    Boolean(textOverlay?.enabled) &&
      Boolean(textOverlay?.text.trim()) &&
      isPhotoEditorAccordionOpen('text') &&
      photoEditorState.maskTool === 'none' &&
      !photoEditorState.draftMask
  );
}

function ensurePhotoEditorTextOverlay() {
  if (!photoEditorState) {
    return null;
  }

  const activeText = getPhotoEditorActiveTextOverlay();

  if (activeText) {
    return activeText;
  }

  const nextText = normalizePhotoEditorTextState(
    getDefaultPhotoEditorTextState({
      enabled: true,
      text: '',
    })
  );
  setPhotoEditorTextCollection([nextText], nextText.id);
  return nextText;
}

function selectPhotoEditorTextOverlay(textId) {
  if (!photoEditorState) {
    return;
  }

  const collection = getPhotoEditorTextCollectionFromState();

  if (!collection.textOverlays.some((textOverlay) => textOverlay.id === textId)) {
    return;
  }

  photoEditorState.activeTextId = textId;
  photoEditorState.activeImageOverlayId = '';
  syncPhotoEditorTextControls();
  syncPhotoEditorImageOverlayControls();
  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender();
  }
}

function clearPhotoEditorTextSelection() {
  if (!photoEditorState?.activeTextId) {
    return false;
  }

  photoEditorState.activeTextId = '';
  photoEditorState.textOverlay = getDefaultPhotoEditorTextState();
  syncPhotoEditorTextControls();
  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender();
  }
  return true;
}

function addPhotoEditorTextOverlay(overrides = {}) {
  if (!photoEditorState) {
    return null;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  const collection = getPhotoEditorTextCollectionFromState();
  const nextText = normalizePhotoEditorTextState(
    getDefaultPhotoEditorTextState({
      enabled: true,
      text: translateUiText('テキスト'),
      y: clampNumber(0.5 + collection.textOverlays.length * 0.06, 0.12, 0.88, 0.5),
      ...overrides,
    })
  );

  beginPhotoEditorHistoryMutation();
  photoEditorState.activeImageOverlayId = '';
  setPhotoEditorTextCollection(
    [...collection.textOverlays, nextText],
    nextText.id
  );
  syncPhotoEditorTextControls();
  syncPhotoEditorImageOverlayControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
  return nextText;
}

function deletePhotoEditorActiveTextOverlay() {
  if (!photoEditorState) {
    return;
  }

  const collection = getPhotoEditorTextCollectionFromState();
  const activeIndex = collection.textOverlays.findIndex(
    (textOverlay) => textOverlay.id === collection.activeTextId
  );

  if (activeIndex < 0) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  const nextOverlays = collection.textOverlays.filter(
    (textOverlay) => textOverlay.id !== collection.activeTextId
  );
  const nextActiveTextId =
    nextOverlays[Math.min(activeIndex, nextOverlays.length - 1)]?.id || '';
  setPhotoEditorTextCollection(nextOverlays, nextActiveTextId);
  syncPhotoEditorTextControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function updatePhotoEditorTextOverlay(nextText = {}, { interactive = true } = {}) {
  if (!photoEditorState) {
    return;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  const activeText = ensurePhotoEditorTextOverlay();

  if (!activeText) {
    return;
  }

  const collection = getPhotoEditorTextCollectionFromState();
  const nextState = normalizePhotoEditorTextState({
    ...activeText,
    ...nextText,
    enabled:
      nextText.enabled !== undefined
        ? nextText.enabled
        : activeText.enabled || String(nextText.text ?? activeText.text).trim(),
    weight:
      nextText.fontKey !== undefined && nextText.weight === undefined
        ? getClosestPhotoEditorTextWeight(nextText.fontKey, activeText.weight)
        : nextText.weight ?? activeText.weight,
  });

  beginPhotoEditorHistoryMutation();
  photoEditorState.activeImageOverlayId = '';
  setPhotoEditorTextCollection(
    collection.textOverlays.map((textOverlay) =>
      textOverlay.id === activeText.id ? nextState : textOverlay
    ),
    nextState.id
  );

  if (
    nextText.fontKey !== undefined ||
    nextText.weight !== undefined ||
    nextText.size !== undefined
  ) {
    rememberPhotoEditorTextFont(nextState.fontKey);
    loadPhotoEditorTextFont(nextState).then(() => {
      if (photoEditorState?.activeTextId === nextState.id) {
        schedulePhotoEditorRender();
      }
    });
  }

  syncPhotoEditorTextControls();
  syncPhotoEditorImageOverlayControls();
  schedulePhotoEditorRender({
    debounceMs: interactive ? PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS : 0,
    interactive,
  });
  schedulePhotoEditorHistoryCommit();
}

function resetPhotoEditorTextOverlay() {
  if (!photoEditorState) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  setPhotoEditorTextCollection([], '');
  photoEditorState.dragInitialText = null;
  photoEditorState.snapGuide = null;
  syncPhotoEditorTextControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function getPhotoEditorImageOverlayLabel(overlay, index) {
  const fileName = String(overlay?.fileName || '').trim();
  return fileName || `画像 ${index + 1}`;
}

function renderPhotoEditorImageOverlayLibrary() {
  if (!photoEditorImageOverlayLibrary) {
    return;
  }

  photoEditorImageOverlayLibrary.innerHTML = '';

  for (const asset of photoEditorOverlayAssets) {
    const item = document.createElement('div');
    item.className = 'photo-editor-image-overlay-library-item';

    const thumbnail = document.createElement('img');
    thumbnail.className = 'photo-editor-image-overlay-library-thumb';
    thumbnail.src = asset.fileUrl;
    thumbnail.alt = asset.fileName;
    thumbnail.loading = 'lazy';
    thumbnail.decoding = 'async';
    item.appendChild(thumbnail);

    const addButton = document.createElement('button');
    addButton.type = 'button';
    addButton.className = 'photo-editor-image-overlay-library-add';
    addButton.textContent = asset.fileName;
    addButton.title = asset.fileName;
    addButton.addEventListener('click', () => {
      addPhotoEditorImageOverlayFromAsset(asset);
    });
    item.appendChild(addButton);

    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'photo-editor-image-overlay-library-delete';
    deleteButton.setAttribute('aria-label', `${asset.fileName}を管理素材から削除`);
    deleteButton.title = '管理素材から削除';
    deleteButton.innerHTML =
      '<span class="material-symbols-outlined">delete</span>';
    deleteButton.addEventListener('click', async (event) => {
      event.stopPropagation();
      await deletePhotoEditorManagedOverlayAsset(asset);
    });
    item.appendChild(deleteButton);

    photoEditorImageOverlayLibrary.appendChild(item);
  }
}

function renderPhotoEditorImageOverlayList(imageOverlays, activeImageOverlayId) {
  if (!photoEditorImageOverlayList) {
    return;
  }

  photoEditorImageOverlayList.innerHTML = '';

  imageOverlays
    .map((overlay, index) => ({ overlay, index }))
    .reverse()
    .forEach(({ overlay, index }) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'photo-editor-image-overlay-item';
      button.dataset.photoEditorImageOverlayId = overlay.id;
      button.textContent = getPhotoEditorImageOverlayLabel(overlay, index);
      button.title = overlay.fileName || '';
      button.classList.toggle('is-active', overlay.id === activeImageOverlayId);
      button.setAttribute(
        'aria-pressed',
        overlay.id === activeImageOverlayId ? 'true' : 'false'
      );
      button.addEventListener('click', () => {
        selectPhotoEditorImageOverlay(overlay.id);
      });
      photoEditorImageOverlayList.appendChild(button);
    });
}

function syncPhotoEditorImageOverlayControls() {
  if (!photoEditorState) {
    return;
  }

  const collection = getPhotoEditorImageOverlayCollectionFromState();
  setPhotoEditorImageOverlayCollection(
    collection.imageOverlays,
    collection.activeImageOverlayId
  );
  const activeOverlay = getPhotoEditorActiveImageOverlay();
  const activeIndex = collection.imageOverlays.findIndex(
    (overlay) => overlay.id === collection.activeImageOverlayId
  );

  renderPhotoEditorImageOverlayList(
    collection.imageOverlays,
    collection.activeImageOverlayId
  );
  renderPhotoEditorImageOverlayLibrary();

  if (photoEditorImageOverlayControls) {
    photoEditorImageOverlayControls.hidden = !activeOverlay;
  }

  if (photoEditorImageOverlayOpacityInput) {
    photoEditorImageOverlayOpacityInput.disabled = !activeOverlay;
    photoEditorImageOverlayOpacityInput.value = String(
      Math.round((activeOverlay?.opacity ?? 1) * 100)
    );
  }

  if (photoEditorImageOverlayOpacityValue) {
    photoEditorImageOverlayOpacityValue.textContent =
      `${Math.round((activeOverlay?.opacity ?? 1) * 100)}%`;
  }

  if (photoEditorImageOverlayBlendModeSelect) {
    photoEditorImageOverlayBlendModeSelect.disabled = !activeOverlay;
    photoEditorImageOverlayBlendModeSelect.value =
      activeOverlay?.blendMode || 'source-over';
  }

  if (photoEditorImageOverlayMaskModeSelect) {
    photoEditorImageOverlayMaskModeSelect.disabled = !activeOverlay;
    photoEditorImageOverlayMaskModeSelect.value =
      activeOverlay?.maskMode || 'normal';
  }

  if (photoEditorImageOverlayForwardButton) {
    photoEditorImageOverlayForwardButton.disabled =
      !activeOverlay || activeIndex >= collection.imageOverlays.length - 1;
  }

  if (photoEditorImageOverlayBackwardButton) {
    photoEditorImageOverlayBackwardButton.disabled =
      !activeOverlay || activeIndex <= 0;
  }

  if (photoEditorImageOverlayDeleteButton) {
    photoEditorImageOverlayDeleteButton.disabled = !activeOverlay;
  }

  if (photoEditorImageOverlayResetButton) {
    photoEditorImageOverlayResetButton.disabled =
      collection.imageOverlays.length === 0;
  }

  photoEditorCanvas?.classList.toggle(
    'is-image-overlay-tool-active',
    Boolean(activeOverlay) &&
      isPhotoEditorAccordionOpen('imageOverlay') &&
      photoEditorState.maskTool === 'none' &&
      !photoEditorState.draftMask
  );
}

async function refreshPhotoEditorOverlayAssets({ silent = false } = {}) {
  if (!window.electronAPI.getPhotoEditorOverlayAssets) {
    return;
  }

  try {
    const result = await window.electronAPI.getPhotoEditorOverlayAssets();

    if (!result?.ok) {
      throw new Error(result?.message || '画像オーバーレイ素材を読み込めませんでした');
    }

    photoEditorOverlayAssets = normalizePhotoEditorOverlayAssets(result.assets);
    renderPhotoEditorImageOverlayLibrary();
  } catch (error) {
    if (!silent) {
      setPhotoEditorStatus(error.message);
    }
  }
}

function addPhotoEditorImageOverlayFromAsset(asset) {
  if (!photoEditorState) {
    return null;
  }

  const normalizedAsset = normalizePhotoEditorOverlayAsset(asset);

  if (!normalizedAsset) {
    setPhotoEditorStatus('画像オーバーレイ素材を読み込めませんでした');
    return null;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  const collection = getPhotoEditorImageOverlayCollectionFromState();

  if (collection.imageOverlays.length >= PHOTO_EDITOR_IMAGE_OVERLAY_LIMIT) {
    setPhotoEditorStatus(`画像オーバーレイは${PHOTO_EDITOR_IMAGE_OVERLAY_LIMIT}件までです`);
    return null;
  }

  const offset = Math.min(collection.imageOverlays.length * 0.045, 0.22);
  const nextOverlay = normalizePhotoEditorImageOverlayState(
    getDefaultPhotoEditorImageOverlayState(normalizedAsset, {
      x: clampNumber(0.5 + offset, 0.12, 0.88, 0.5),
      y: clampNumber(0.5 + offset, 0.12, 0.88, 0.5),
    })
  );

  beginPhotoEditorHistoryMutation();
  photoEditorState.activeTextId = '';
  setPhotoEditorImageOverlayCollection(
    [...collection.imageOverlays, nextOverlay],
    nextOverlay.id
  );
  syncPhotoEditorTextControls();
  getPhotoEditorOverlayImageCacheEntry(nextOverlay);
  syncPhotoEditorImageOverlayControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
  return nextOverlay;
}

async function selectPhotoEditorOverlayImages() {
  if (!window.electronAPI.selectPhotoEditorOverlayImages) {
    setPhotoEditorStatus('画像オーバーレイ素材の追加機能を利用できません');
    return;
  }

  try {
    const result = await window.electronAPI.selectPhotoEditorOverlayImages();

    if (result?.assets) {
      photoEditorOverlayAssets = normalizePhotoEditorOverlayAssets(result.assets);
      renderPhotoEditorImageOverlayLibrary();
    }

    if (result?.canceled) {
      return;
    }

    if (!result?.ok && result?.message) {
      setPhotoEditorStatus(result.message);
    }

    const importedAssets = normalizePhotoEditorOverlayAssets(
      result?.importedAssets
    );

    for (const asset of importedAssets) {
      addPhotoEditorImageOverlayFromAsset(asset);
    }

    if (importedAssets.length > 0) {
      setPhotoEditorStatus(`画像オーバーレイを追加しました: ${importedAssets.length}件`);
    }
  } catch (error) {
    setPhotoEditorStatus(`画像オーバーレイの追加に失敗しました: ${error.message}`);
  }
}

function selectPhotoEditorImageOverlay(overlayId) {
  if (!photoEditorState) {
    return;
  }

  const collection = getPhotoEditorImageOverlayCollectionFromState();

  if (!collection.imageOverlays.some((overlay) => overlay.id === overlayId)) {
    return;
  }

  photoEditorState.activeImageOverlayId = overlayId;
  photoEditorState.activeTextId = '';
  syncPhotoEditorTextControls();
  syncPhotoEditorImageOverlayControls();
  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender();
  }
}

function clearPhotoEditorImageOverlaySelection() {
  if (!photoEditorState?.activeImageOverlayId) {
    return false;
  }

  photoEditorState.activeImageOverlayId = '';
  syncPhotoEditorImageOverlayControls();
  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender();
  }
  return true;
}

function updatePhotoEditorActiveImageOverlay(nextValues = {}, { interactive = true } = {}) {
  if (!photoEditorState) {
    return;
  }

  const collection = getPhotoEditorImageOverlayCollectionFromState();
  const activeOverlay = getPhotoEditorActiveImageOverlay();

  if (!activeOverlay) {
    return;
  }

  const nextOverlay = normalizePhotoEditorImageOverlayState({
    ...activeOverlay,
    ...nextValues,
  });

  beginPhotoEditorHistoryMutation();
  setPhotoEditorImageOverlayCollection(
    collection.imageOverlays.map((overlay) =>
      overlay.id === activeOverlay.id ? nextOverlay : overlay
    ),
    nextOverlay.id
  );
  syncPhotoEditorImageOverlayControls();
  schedulePhotoEditorRender({
    debounceMs: interactive ? PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS : 0,
    interactive,
  });
  schedulePhotoEditorHistoryCommit();
}

function movePhotoEditorActiveImageOverlay(delta) {
  if (!photoEditorState) {
    return;
  }

  const collection = getPhotoEditorImageOverlayCollectionFromState();
  const currentIndex = collection.imageOverlays.findIndex(
    (overlay) => overlay.id === collection.activeImageOverlayId
  );

  if (currentIndex < 0) {
    return;
  }

  const nextIndex = clampNumber(
    currentIndex + delta,
    0,
    collection.imageOverlays.length - 1,
    currentIndex
  );

  if (nextIndex === currentIndex) {
    return;
  }

  const nextOverlays = [...collection.imageOverlays];
  const [movedOverlay] = nextOverlays.splice(currentIndex, 1);
  nextOverlays.splice(nextIndex, 0, movedOverlay);

  beginPhotoEditorHistoryMutation();
  setPhotoEditorImageOverlayCollection(nextOverlays, movedOverlay.id);
  syncPhotoEditorImageOverlayControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function deletePhotoEditorActiveImageOverlay() {
  if (!photoEditorState) {
    return;
  }

  const collection = getPhotoEditorImageOverlayCollectionFromState();
  const activeIndex = collection.imageOverlays.findIndex(
    (overlay) => overlay.id === collection.activeImageOverlayId
  );

  if (activeIndex < 0) {
    return;
  }

  const nextOverlays = collection.imageOverlays.filter(
    (overlay) => overlay.id !== collection.activeImageOverlayId
  );
  const nextActiveId =
    nextOverlays[Math.min(activeIndex, nextOverlays.length - 1)]?.id || '';

  beginPhotoEditorHistoryMutation();
  setPhotoEditorImageOverlayCollection(nextOverlays, nextActiveId);
  syncPhotoEditorImageOverlayControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function resetPhotoEditorImageOverlays() {
  if (!photoEditorState) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  setPhotoEditorImageOverlayCollection([], '');
  photoEditorState.dragInitialImageOverlay = null;
  photoEditorState.snapGuide = null;
  syncPhotoEditorImageOverlayControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

async function deletePhotoEditorManagedOverlayAsset(asset) {
  const normalizedAsset = normalizePhotoEditorOverlayAsset(asset);

  if (!normalizedAsset || !window.electronAPI.deletePhotoEditorOverlayAsset) {
    return;
  }

  const confirmed = await openConfirmModal({
    title: '管理素材を削除',
    message: `${normalizedAsset.fileName} を管理素材から削除します。現在の編集で使っている同じ画像レイヤーも外します。`,
    confirmText: '削除する',
  });

  if (!confirmed) {
    return;
  }

  try {
    const result = await window.electronAPI.deletePhotoEditorOverlayAsset(
      normalizedAsset.id
    );

    if (!result?.ok) {
      throw new Error(result?.message || '管理素材を削除できませんでした');
    }

    const collection = getPhotoEditorImageOverlayCollectionFromState();
    const nextOverlays = collection.imageOverlays.filter(
      (overlay) => overlay.assetId !== normalizedAsset.id
    );

    if (nextOverlays.length !== collection.imageOverlays.length) {
      setPhotoEditorImageOverlayCollection(nextOverlays, '');
    }

    photoEditorOverlayAssets = normalizePhotoEditorOverlayAssets(result.assets);
    photoEditorOverlayImageCache.delete(normalizedAsset.fileUrl);
    syncPhotoEditorImageOverlayControls();
    schedulePhotoEditorRender();
    showToast('管理素材を削除しました');
  } catch (error) {
    setPhotoEditorStatus(`管理素材の削除に失敗しました: ${error.message}`);
  }
}

function syncPhotoEditorExportControls() {
  if (!photoEditorState) {
    return;
  }

  const settings = getPhotoEditorEffectiveExportSettings(
    photoEditorState.exportSettings
  );
  photoEditorState.exportSettings = settings;
  const formatMeta = PHOTO_EDITOR_EXPORT_FORMATS[settings.format];

  if (photoEditorExportFormatSelect) {
    photoEditorExportFormatSelect.value = settings.format;
  }

  if (photoEditorExportMaxEdgeSelect) {
    photoEditorExportMaxEdgeSelect.value = String(settings.maxEdge);
  }

  if (photoEditorExportQualityInput) {
    photoEditorExportQualityInput.value = String(settings.quality);
    photoEditorExportQualityInput.disabled = !formatMeta.supportsQuality;
  }

  if (photoEditorExportQualityValue) {
    photoEditorExportQualityValue.textContent = formatMeta.supportsQuality
      ? String(settings.quality)
      : '-';
  }

  if (photoEditorExportQualityRow) {
    photoEditorExportQualityRow.hidden = !formatMeta.supportsQuality;
  }
}

function syncPhotoEditorCompareControl() {
  const isComparing = Boolean(photoEditorState?.showOriginalPreview);

  if (!photoEditorCompareButton) {
    return;
  }

  photoEditorCompareButton.classList.toggle('is-active', isComparing);
  photoEditorCompareButton.setAttribute('aria-pressed', isComparing ? 'true' : 'false');
  photoEditorCompareButton.title = isComparing
    ? '編集後の表示に戻す'
    : '編集前と比較';
  const label = photoEditorCompareButton.querySelector('span:last-child');

  if (label) {
    label.textContent = isComparing ? '編集中' : '比較';
  }
}

function syncPhotoEditorOverlayControls() {
  if (photoEditorRuleGridButton) {
    const isGridVisible = Boolean(photoEditorState?.showRuleOfThirdsGrid);
    photoEditorRuleGridButton.classList.toggle('is-active', isGridVisible);
    photoEditorRuleGridButton.setAttribute(
      'aria-pressed',
      isGridVisible ? 'true' : 'false'
    );
  }

  if (photoEditorRulerButton) {
    const isRulerVisible = Boolean(photoEditorState?.showRulers);
    photoEditorRulerButton.classList.toggle('is-active', isRulerVisible);
    photoEditorRulerButton.setAttribute(
      'aria-pressed',
      isRulerVisible ? 'true' : 'false'
    );
  }
}

function syncPhotoEditorUi() {
  syncPhotoEditorAdjustmentControls();
  syncPhotoEditorAutoEnhanceControls();
  syncPhotoEditorCropControls();
  syncPhotoEditorTextControls();
  syncPhotoEditorImageOverlayControls();
  syncPhotoEditorSubjectMaskControls();
  syncPhotoEditorExportControls();
  syncPhotoEditorBlurControls();
  syncPhotoEditorCurveControls();
  syncPhotoEditorMaskToolUi();
  syncPhotoEditorCompareControl();
  syncPhotoEditorOverlayControls();
  syncPhotoEditorHistoryControls();

  if (photoEditorFillColorInput && photoEditorState) {
    photoEditorFillColorInput.value = photoEditorState.fillColor;
  }

  if (photoEditorSaveButton && photoEditorState) {
    photoEditorSaveButton.disabled = photoEditorState.isSaving;
    photoEditorSaveButton.textContent = photoEditorState.isSaving
      ? '保存中...'
      : '別名で保存';
  }

  if (photoEditorSubjectTransparentSaveButton && photoEditorState) {
    photoEditorSubjectTransparentSaveButton.disabled =
      photoEditorState.isSaving || !hasPhotoEditorReadySubjectMask();
  }
}

function getPhotoEditorCropPreset(key) {
  return (
    PHOTO_EDITOR_CROP_PRESETS.find((preset) => preset.key === key) ||
    PHOTO_EDITOR_CROP_PRESETS[0]
  );
}

function isPhotoEditorTransparentPaddingCrop(crop = photoEditorState?.crop) {
  return Boolean(getPhotoEditorCropPreset(crop?.preset).transparentPadding);
}

function getPhotoEditorCropRatio(image) {
  const preset = getPhotoEditorCropPreset(photoEditorState?.crop?.preset);

  if (preset.transparentPadding) {
    const width = Number(image?.naturalWidth || image?.width);
    const height = Number(image?.naturalHeight || image?.height);
    return width > 0 && height > 0 ? width / height : 1;
  }

  if (Number.isFinite(preset.ratio) && preset.ratio > 0) {
    return preset.ratio;
  }

  const width = Number(image?.naturalWidth || image?.width);
  const height = Number(image?.naturalHeight || image?.height);

  return width > 0 && height > 0 ? width / height : 1;
}

function getPhotoEditorSourceRect(image) {
  const imageWidth = Number(image?.naturalWidth || image?.width) || 0;
  const imageHeight = Number(image?.naturalHeight || image?.height) || 0;

  if (imageWidth <= 0 || imageHeight <= 0) {
    return {
      x: 0,
      y: 0,
      width: 1,
      height: 1,
    };
  }

  const crop = photoEditorState?.crop || getDefaultPhotoEditorCropState();

  if (isPhotoEditorTransparentPaddingCrop(crop)) {
    return {
      x: 0,
      y: 0,
      width: imageWidth,
      height: imageHeight,
    };
  }

  const targetRatio = getPhotoEditorCropRatio(image);
  const imageRatio = imageWidth / imageHeight;
  let baseWidth = imageWidth;
  let baseHeight = imageHeight;

  if (imageRatio > targetRatio) {
    baseWidth = imageHeight * targetRatio;
  } else {
    baseHeight = imageWidth / targetRatio;
  }

  const zoom = getPhotoEditorCropZoomScale(crop);
  const cropWidth = Math.max(1, baseWidth / zoom);
  const cropHeight = Math.max(1, baseHeight / zoom);
  const maxOffsetX = Math.max(0, (imageWidth - cropWidth) / 2);
  const maxOffsetY = Math.max(0, (imageHeight - cropHeight) / 2);
  const offsetX = maxOffsetX * clampNumber(crop.offsetX, -100, 100, 0) / 100;
  const offsetY = maxOffsetY * clampNumber(crop.offsetY, -100, 100, 0) / 100;
  const centerX = imageWidth / 2 + offsetX;
  const centerY = imageHeight / 2 + offsetY;

  return {
    x: clampNumber(centerX - cropWidth / 2, 0, imageWidth - cropWidth, 0),
    y: clampNumber(centerY - cropHeight / 2, 0, imageHeight - cropHeight, 0),
    width: cropWidth,
    height: cropHeight,
  };
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

function getPhotoEditorFixedOutputSize(crop = photoEditorState?.crop) {
  const preset = getPhotoEditorCropPreset(crop?.preset);
  const width = Number(preset.exportWidth);
  const height = Number(preset.exportHeight);

  if (
    Number.isFinite(width) &&
    Number.isFinite(height) &&
    width > 0 &&
    height > 0
  ) {
    return {
      width: Math.max(1, Math.round(width)),
      height: Math.max(1, Math.round(height)),
    };
  }

  return null;
}

function getPhotoEditorOutputSize(
  sourceRect,
  maxEdge = null,
  crop = null,
  fixedOutputSize = null
) {
  const activeCrop = crop || photoEditorState?.crop;
  const isSideways = isPhotoEditorCropRotationSideways(
    activeCrop
  );
  const sourceWidth = Math.max(1, Math.round(sourceRect.width));
  const sourceHeight = Math.max(1, Math.round(sourceRect.height));
  const baseFullWidth = isSideways ? sourceHeight : sourceWidth;
  const baseFullHeight = isSideways ? sourceWidth : sourceHeight;
  const shouldUseTransparentPadding =
    isPhotoEditorTransparentPaddingCrop(activeCrop);
  const transparentFullEdge = Math.max(baseFullWidth, baseFullHeight);
  const fullWidth = shouldUseTransparentPadding
    ? transparentFullEdge
    : baseFullWidth;
  const fullHeight = shouldUseTransparentPadding
    ? transparentFullEdge
    : baseFullHeight;
  const fixedWidth = Number(fixedOutputSize?.width);
  const fixedHeight = Number(fixedOutputSize?.height);

  if (
    Number.isFinite(fixedWidth) &&
    Number.isFinite(fixedHeight) &&
    fixedWidth > 0 &&
    fixedHeight > 0
  ) {
    return {
      width: Math.max(1, Math.round(fixedWidth)),
      height: Math.max(1, Math.round(fixedHeight)),
      fullWidth,
      fullHeight,
      scale: 1,
    };
  }

  const numericMaxEdge = Number(maxEdge);

  if (!Number.isFinite(numericMaxEdge) || numericMaxEdge <= 0) {
    return {
      width: fullWidth,
      height: fullHeight,
      fullWidth,
      fullHeight,
      scale: 1,
    };
  }

  const scale = Math.min(1, numericMaxEdge / Math.max(fullWidth, fullHeight));

  return {
    width: Math.max(1, Math.round(fullWidth * scale)),
    height: Math.max(1, Math.round(fullHeight * scale)),
    fullWidth,
    fullHeight,
    scale,
  };
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

function drawPhotoEditorCroppedSourceToCanvas(
  ctx,
  image,
  sourceRect,
  outputSize,
  crop = {}
) {
  const geometry = getPhotoEditorCropRenderGeometry(
    sourceRect,
    outputSize.width,
    outputSize.height,
    crop
  );
  const drawWidth = geometry.drawWidth * geometry.coverScale;
  const drawHeight = geometry.drawHeight * geometry.coverScale;

  ctx.save();
  ctx.translate(
    outputSize.width / 2 + geometry.translateX,
    outputSize.height / 2 + geometry.translateY
  );
  ctx.rotate(geometry.rotationRadians);
  ctx.scale(geometry.flipX ? -1 : 1, geometry.flipY ? -1 : 1);
  ctx.drawImage(
    image,
    sourceRect.x,
    sourceRect.y,
    sourceRect.width,
    sourceRect.height,
    -drawWidth / 2,
    -drawHeight / 2,
    drawWidth,
    drawHeight
  );
  ctx.restore();
}

function syncPhotoEditorPreviewCanvasDisplaySize(outputSize) {
  if (!photoEditorCanvas || !photoEditorCanvasWrap || !outputSize) {
    return;
  }

  const fullWidth = Math.max(1, Number(outputSize.fullWidth) || outputSize.width || 1);
  const fullHeight = Math.max(1, Number(outputSize.fullHeight) || outputSize.height || 1);
  const bounds = photoEditorCanvasWrap.getBoundingClientRect();
  photoEditorCanvas.style.aspectRatio = `${fullWidth} / ${fullHeight}`;

  if (!bounds.width || !bounds.height) {
    photoEditorCanvas.style.width = '';
    photoEditorCanvas.style.height = '';
    return;
  }

  const wrapStyle = getComputedStyle(photoEditorCanvasWrap);
  const paddingX =
    (parseFloat(wrapStyle.paddingLeft) || 0) +
    (parseFloat(wrapStyle.paddingRight) || 0);
  const paddingY =
    (parseFloat(wrapStyle.paddingTop) || 0) +
    (parseFloat(wrapStyle.paddingBottom) || 0);
  const availableWidth = Math.max(1, bounds.width - paddingX);
  const availableHeight = Math.max(1, bounds.height - paddingY);
  const displayScale = Math.min(
    availableWidth / fullWidth,
    availableHeight / fullHeight
  );

  photoEditorCanvas.style.width = `${Math.max(1, Math.round(fullWidth * displayScale))}px`;
  photoEditorCanvas.style.height = `${Math.max(1, Math.round(fullHeight * displayScale))}px`;
}

function getPhotoEditorPreviewWorkCanvas() {
  if (!photoEditorPreviewWorkCanvas) {
    photoEditorPreviewWorkCanvas = document.createElement('canvas');
  }

  return photoEditorPreviewWorkCanvas;
}

function getPhotoEditorPreviewOverlayCanvas() {
  if (!photoEditorPreviewOverlayCanvas) {
    photoEditorPreviewOverlayCanvas = document.createElement('canvas');
  }

  return photoEditorPreviewOverlayCanvas;
}

function copyPhotoEditorCanvasFrame(sourceCanvas, targetCanvas, outputSize) {
  if (!sourceCanvas || !targetCanvas || !outputSize) {
    return null;
  }

  if (targetCanvas.width !== outputSize.width) {
    targetCanvas.width = outputSize.width;
  }

  if (targetCanvas.height !== outputSize.height) {
    targetCanvas.height = outputSize.height;
  }

  const ctx = targetCanvas.getContext('2d', { willReadFrequently: true });

  if (!ctx) {
    return null;
  }

  ctx.clearRect(0, 0, outputSize.width, outputSize.height);
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(sourceCanvas, 0, 0, outputSize.width, outputSize.height);
  return ctx;
}

function commitPhotoEditorPreviewFrame(sourceCanvas, outputSize) {
  if (!photoEditorCanvas || !sourceCanvas || !outputSize) {
    return null;
  }

  return copyPhotoEditorCanvasFrame(sourceCanvas, photoEditorCanvas, outputSize);
}

function storePhotoEditorPreviewOverlayBase(sourceCanvas, outputSize, sourceRect) {
  const overlayCanvas = getPhotoEditorPreviewOverlayCanvas();
  const ctx = copyPhotoEditorCanvasFrame(sourceCanvas, overlayCanvas, outputSize);

  if (!ctx) {
    photoEditorPreviewOverlayMeta = null;
    return;
  }

  photoEditorPreviewOverlayMeta = {
    outputSize: { ...outputSize },
    sourceRect: { ...sourceRect },
  };
}

function drawPhotoEditorRuleOfThirdsGrid(ctx, width, height) {
  if (!ctx || width <= 0 || height <= 0) {
    return;
  }

  ctx.save();
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.68)';
  ctx.lineWidth = Math.max(1, Math.min(width, height) / 360);
  ctx.setLineDash([Math.max(5, width / 96), Math.max(5, width / 96)]);
  ctx.shadowColor = 'rgba(15, 23, 42, 0.55)';
  ctx.shadowBlur = Math.max(2, ctx.lineWidth * 2);

  for (const ratio of [1 / 3, 2 / 3]) {
    const x = width * ratio;
    const y = height * ratio;
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, height);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(0, y);
    ctx.lineTo(width, y);
    ctx.stroke();
  }

  ctx.restore();
}

function getPhotoEditorRulerEdgeSize(width, height) {
  return Math.max(18, Math.min(width, height) * 0.035);
}

function drawPhotoEditorRulers(ctx, width, height) {
  if (!ctx || width <= 0 || height <= 0) {
    return;
  }

  const edge = getPhotoEditorRulerEdgeSize(width, height);
  const majorStep = Math.max(40, Math.min(width, height) / 6);
  const minorStep = majorStep / 4;

  ctx.save();
  ctx.fillStyle = 'rgba(15, 23, 42, 0.42)';
  ctx.fillRect(0, 0, width, edge);
  ctx.fillRect(0, 0, edge, height);
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.74)';
  ctx.lineWidth = Math.max(1, Math.min(width, height) / 420);
  ctx.shadowColor = 'rgba(15, 23, 42, 0.5)';
  ctx.shadowBlur = 2;

  for (let x = 0; x <= width; x += minorStep) {
    const isMajor = Math.abs((x / majorStep) - Math.round(x / majorStep)) < 0.01;
    const length = isMajor ? edge * 0.82 : edge * 0.48;
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, length);
    ctx.stroke();
  }

  for (let y = 0; y <= height; y += minorStep) {
    const isMajor = Math.abs((y / majorStep) - Math.round(y / majorStep)) < 0.01;
    const length = isMajor ? edge * 0.82 : edge * 0.48;
    ctx.beginPath();
    ctx.moveTo(0, y);
    ctx.lineTo(length, y);
    ctx.stroke();
  }

  ctx.restore();
}

function drawPhotoEditorRulerGuides(ctx, width, height) {
  if (!ctx || !photoEditorState?.showRulers || width <= 0 || height <= 0) {
    return;
  }

  const guides = normalizePhotoEditorRulerGuides(photoEditorState.rulerGuides);
  const draftGuide = photoEditorState.draftRulerGuide;
  const verticalGuides = [...guides.x];
  const horizontalGuides = [...guides.y];

  if (draftGuide?.axis === 'x') {
    verticalGuides.push(clampNumber(draftGuide.position, 0, 1, 0.5));
  } else if (draftGuide?.axis === 'y') {
    horizontalGuides.push(clampNumber(draftGuide.position, 0, 1, 0.5));
  }

  if (verticalGuides.length === 0 && horizontalGuides.length === 0) {
    return;
  }

  const edge = Math.max(width, height);
  const lineWidth = Math.max(1.5, edge / 1100);

  ctx.save();
  ctx.strokeStyle = 'rgba(96, 165, 250, 0.82)';
  ctx.lineWidth = lineWidth;
  ctx.setLineDash([Math.max(8, lineWidth * 4), Math.max(6, lineWidth * 3)]);
  ctx.shadowColor = 'rgba(15, 23, 42, 0.58)';
  ctx.shadowBlur = Math.max(2, lineWidth * 2);

  for (const guideX of verticalGuides) {
    const x = clampNumber(guideX, 0, 1, 0.5) * width;
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, height);
    ctx.stroke();
  }

  for (const guideY of horizontalGuides) {
    const y = clampNumber(guideY, 0, 1, 0.5) * height;
    ctx.beginPath();
    ctx.moveTo(0, y);
    ctx.lineTo(width, y);
    ctx.stroke();
  }

  ctx.restore();
}

function drawPhotoEditorPreviewOverlays(ctx, outputSize, sourceRect) {
  if (!ctx || !photoEditorState || photoEditorState.showOriginalPreview) {
    return;
  }

  if (photoEditorState.showRuleOfThirdsGrid) {
    drawPhotoEditorRuleOfThirdsGrid(ctx, outputSize.width, outputSize.height);
  }

  if (photoEditorState.showRulers) {
    drawPhotoEditorRulers(ctx, outputSize.width, outputSize.height);
    drawPhotoEditorRulerGuides(ctx, outputSize.width, outputSize.height);
  }

  drawPhotoEditorSnapGuides(ctx, outputSize.width, outputSize.height);
  drawPhotoEditorDraftMask(
    ctx,
    outputSize.width,
    outputSize.height,
    sourceRect,
    photoEditorState.sourceImage
  );
  drawPhotoEditorRadialBlurControl(
    ctx,
    outputSize.width,
    outputSize.height
  );
  drawPhotoEditorImageOverlayControls(
    ctx,
    outputSize.width,
    outputSize.height
  );
  drawPhotoEditorTextControls(ctx, outputSize.width, outputSize.height);
}

function drawPhotoEditorOverlayGroup(
  ctx,
  outputSize,
  textOverlays,
  imageOverlays,
  maskMode,
  { textScale = 1, sourceRect = null, includeImageOverlays = true } = {}
) {
  if (includeImageOverlays) {
    applyPhotoEditorImageOverlayToCanvas(
      ctx,
      outputSize.width,
      outputSize.height,
      imageOverlays,
      { maskMode, outputSize, sourceRect }
    );
  }
  applyPhotoEditorTextOverlayToCanvas(
    ctx,
    outputSize.width,
    outputSize.height,
    textOverlays,
    { textScale, maskMode }
  );
}

function applyPhotoEditorMaskedOverlaysToCanvas(
  ctx,
  outputSize,
  sourceRect,
  { textScale = 1 } = {}
) {
  if (!ctx || !outputSize || !photoEditorState) {
    return;
  }

  const textOverlays = getPhotoEditorTextCollectionFromState().textOverlays;
  const imageOverlays = getPhotoEditorImageOverlayCollectionFromState().imageOverlays;

  if (!hasPhotoEditorReadySubjectMask()) {
    applyPhotoEditorImageOverlayToCanvas(
      ctx,
      outputSize.width,
      outputSize.height,
      imageOverlays
    );
    applyPhotoEditorTextOverlayToCanvas(
      ctx,
      outputSize.width,
      outputSize.height,
      textOverlays,
      { textScale }
    );
    return;
  }

  const baseBeforeOverlays = document.createElement('canvas');
  baseBeforeOverlays.width = outputSize.width;
  baseBeforeOverlays.height = outputSize.height;
  const baseCtx = baseBeforeOverlays.getContext('2d');

  if (baseCtx) {
    baseCtx.drawImage(ctx.canvas, 0, 0);
  }

  applyPhotoEditorImageOverlayToCanvas(
    ctx,
    outputSize.width,
    outputSize.height,
    imageOverlays,
    { maskMode: 'background-only', outputSize, sourceRect }
  );
  withPhotoEditorSubjectMaskLayer(
    ctx,
    outputSize,
    sourceRect,
    'background-only',
    (layerCtx) => {
      drawPhotoEditorOverlayGroup(
        layerCtx,
        outputSize,
        textOverlays,
        imageOverlays,
        'background-only',
        { textScale, includeImageOverlays: false }
      );
    }
  );
  if (baseCtx) {
    drawPhotoEditorCanvasWithSubjectMask(
      ctx,
      baseBeforeOverlays,
      outputSize,
      sourceRect,
      'subject-only'
    );
  }

  applyPhotoEditorImageOverlayToCanvas(
    ctx,
    outputSize.width,
    outputSize.height,
    imageOverlays,
    { maskMode: 'subject-only', outputSize, sourceRect }
  );
  withPhotoEditorSubjectMaskLayer(
    ctx,
    outputSize,
    sourceRect,
    'subject-only',
    (layerCtx) => {
      drawPhotoEditorOverlayGroup(
        layerCtx,
        outputSize,
        textOverlays,
        imageOverlays,
        'subject-only',
        { textScale, includeImageOverlays: false }
      );
    }
  );
  drawPhotoEditorOverlayGroup(
    ctx,
    outputSize,
    textOverlays,
    imageOverlays,
    'normal',
    { textScale }
  );
}

function drawPhotoEditorPreviewTextAndOverlays(ctx, outputSize, sourceRect) {
  if (!ctx || !outputSize || !photoEditorState) {
    return;
  }

  applyPhotoEditorMaskedOverlaysToCanvas(ctx, outputSize, sourceRect, {
    textScale: 1,
  });
  drawPhotoEditorSubjectMaskPreview(ctx, outputSize, sourceRect);
  drawPhotoEditorPreviewOverlays(ctx, outputSize, sourceRect);
}

function paintPhotoEditorPreviewOverlayOnly() {
  if (
    !photoEditorState ||
    !photoEditorPreviewOverlayCanvas ||
    !photoEditorPreviewOverlayMeta ||
    photoEditorState.showOriginalPreview
  ) {
    return false;
  }

  const { outputSize, sourceRect } = photoEditorPreviewOverlayMeta;
  const ctx = commitPhotoEditorPreviewFrame(photoEditorPreviewOverlayCanvas, outputSize);

  if (!ctx) {
    return false;
  }

  drawPhotoEditorPreviewTextAndOverlays(ctx, outputSize, sourceRect);
  return true;
}

function getDeterministicNoise(x, y) {
  let seed = ((x + 1) * 374761393 + (y + 1) * 668265263) >>> 0;
  seed = (seed ^ (seed >>> 13)) >>> 0;
  seed = Math.imul(seed, 1274126177) >>> 0;
  return (seed / 4294967295) * 2 - 1;
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
      isPhotoEditorCurvePointsIdentity(curve.points[mode][channel.key])
    )
  );
}

function createPhotoEditorCurveLookups(curveState) {
  const curve = normalizePhotoEditorCurveState(curveState);
  const rgbIdentity = {};
  const hsvIdentity = {};

  for (const channel of PHOTO_EDITOR_CURVE_CHANNELS.rgb) {
    rgbIdentity[channel.key] = isPhotoEditorCurvePointsIdentity(
      curve.points.rgb[channel.key]
    );
  }

  for (const channel of PHOTO_EDITOR_CURVE_CHANNELS.hsv) {
    hsvIdentity[channel.key] = isPhotoEditorCurvePointsIdentity(
      curve.points.hsv[channel.key]
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

function applyPhotoEditorAdjustmentsToCanvas(
  ctx,
  width,
  height,
  values,
  curveState = photoEditorState?.curve
) {
  const renderParams = getPhotoEditorAdjustmentRenderParams(values, curveState);

  if (renderParams.hasPixelWork) {
    let imageData;

    try {
      imageData = ctx.getImageData(0, 0, width, height);
    } catch {
      return;
    }

    applyPhotoEditorPixelAdjustmentsToImageData(
      imageData,
      width,
      renderParams
    );
    ctx.putImageData(imageData, 0, 0);
  }

  applyPhotoEditorDenoiseToCanvas(ctx, width, height, values.denoise);
  applyPhotoEditorClarityToCanvas(ctx, width, height, values.clarity);
  applyPhotoEditorTextureToCanvas(ctx, width, height, values.texture);
  applyPhotoEditorSharpnessToCanvas(ctx, width, height, values.sharpness);
  applyPhotoEditorVignette(ctx, width, height, values.vignette);
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

  const tempCanvas = document.createElement('canvas');
  tempCanvas.width = width;
  tempCanvas.height = height;
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

  const tempCanvas = document.createElement('canvas');
  tempCanvas.width = width;
  tempCanvas.height = height;
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

  const tempCanvas = document.createElement('canvas');
  tempCanvas.width = width;
  tempCanvas.height = height;
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

  const tempCanvas = document.createElement('canvas');
  tempCanvas.width = width;
  tempCanvas.height = height;
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

  const isWhiteVignette = vignette > 0;
  const strength = Math.abs(vignette) / 100;
  const edgeColor = isWhiteVignette ? '255, 255, 255' : '0, 0, 0';
  const radius = Math.hypot(width, height) * 0.54;
  const gradient = ctx.createRadialGradient(
    width / 2,
    height / 2,
    Math.min(width, height) * 0.28,
    width / 2,
    height / 2,
    radius
  );
  gradient.addColorStop(0, 'rgba(0, 0, 0, 0)');
  gradient.addColorStop(1, `rgba(${edgeColor}, ${0.56 * strength})`);

  ctx.save();
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, width, height);
  ctx.restore();
}

function getPhotoEditorBlurGeometry(width, height, blurState = null) {
  const blur = normalizePhotoEditorBlurState(blurState);
  const minEdge = Math.max(1, Math.min(width, height));
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
    minEdge,
  };
}

function drawPhotoEditorBlurPaddingSource(targetCtx, sourceCanvas, padding, width, height) {
  targetCtx.drawImage(sourceCanvas, padding, padding);
  targetCtx.drawImage(sourceCanvas, 0, 0, width, 1, padding, 0, width, padding);
  targetCtx.drawImage(
    sourceCanvas,
    0,
    height - 1,
    width,
    1,
    padding,
    padding + height,
    width,
    padding
  );
  targetCtx.drawImage(sourceCanvas, 0, 0, 1, height, 0, padding, padding, height);
  targetCtx.drawImage(
    sourceCanvas,
    width - 1,
    0,
    1,
    height,
    padding + width,
    padding,
    padding,
    height
  );
  targetCtx.drawImage(sourceCanvas, 0, 0, 1, 1, 0, 0, padding, padding);
  targetCtx.drawImage(
    sourceCanvas,
    width - 1,
    0,
    1,
    1,
    padding + width,
    0,
    padding,
    padding
  );
  targetCtx.drawImage(
    sourceCanvas,
    0,
    height - 1,
    1,
    1,
    0,
    padding + height,
    padding,
    padding
  );
  targetCtx.drawImage(
    sourceCanvas,
    width - 1,
    height - 1,
    1,
    1,
    padding + width,
    padding + height,
    padding,
    padding
  );
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
  const paddedCanvas = document.createElement('canvas');
  paddedCanvas.width = paddedWidth;
  paddedCanvas.height = paddedHeight;
  const paddedCtx = paddedCanvas.getContext('2d');

  if (!paddedCtx) {
    return null;
  }

  drawPhotoEditorBlurPaddingSource(paddedCtx, ctx.canvas, padding, width, height);

  const paddedBlurCanvas = document.createElement('canvas');
  paddedBlurCanvas.width = paddedWidth;
  paddedBlurCanvas.height = paddedHeight;
  const paddedBlurCtx = paddedBlurCanvas.getContext('2d');

  if (!paddedBlurCtx) {
    return null;
  }

  paddedBlurCtx.filter = `blur(${blurRadius}px)`;
  paddedBlurCtx.drawImage(paddedCanvas, 0, 0);
  paddedBlurCtx.filter = 'none';

  const tempCanvas = document.createElement('canvas');
  tempCanvas.width = width;
  tempCanvas.height = height;
  const tempCtx = tempCanvas.getContext('2d');

  if (!tempCtx) {
    return null;
  }

  tempCtx.drawImage(
    paddedBlurCanvas,
    padding,
    padding,
    width,
    height,
    0,
    0,
    width,
    height
  );

  return {
    canvas: tempCanvas,
    ctx: tempCtx,
    radius: blurRadius,
  };
}

function applyPhotoEditorBlurEffectToCanvas(
  ctx,
  width,
  height,
  blurState = null,
  sourceRect = null,
  { isInteractivePreview = false } = {}
) {
  const blur = normalizePhotoEditorBlurState(blurState);
  const blurred = createPhotoEditorBlurredCanvas(
    ctx,
    width,
    height,
    blur.amount,
    { isInteractivePreview }
  );

  if (!blurred) {
    return;
  }

  if (blur.mode === 'full') {
    ctx.drawImage(blurred.canvas, 0, 0);
    return;
  }

  if (blur.mode === 'background') {
    if (!sourceRect || !hasPhotoEditorReadySubjectMask()) {
      return;
    }

    drawPhotoEditorCanvasWithSubjectMask(
      ctx,
      blurred.canvas,
      { width, height },
      sourceRect,
      'background-only'
    );
    return;
  }

  const { centerX, centerY, radius, outerRadius } = getPhotoEditorBlurGeometry(
    width,
    height,
    blur
  );
  const maskCanvas = document.createElement('canvas');
  maskCanvas.width = width;
  maskCanvas.height = height;
  const maskCtx = maskCanvas.getContext('2d');

  if (!maskCtx) {
    return;
  }

  const gradient = maskCtx.createRadialGradient(
    centerX,
    centerY,
    radius,
    centerX,
    centerY,
    outerRadius
  );
  gradient.addColorStop(0, 'rgba(0, 0, 0, 0)');
  gradient.addColorStop(1, 'rgba(0, 0, 0, 1)');

  maskCtx.fillStyle = gradient;
  maskCtx.fillRect(0, 0, width, height);
  blurred.ctx.globalCompositeOperation = 'destination-in';
  blurred.ctx.drawImage(maskCanvas, 0, 0);
  blurred.ctx.globalCompositeOperation = 'source-over';
  ctx.drawImage(blurred.canvas, 0, 0);
}

function getPhotoEditorSubjectMaskImageCacheEntry(subjectMask) {
  const normalizedMask = normalizePhotoEditorSubjectMaskState(subjectMask);
  const cacheKey = normalizedMask.maskDataUrl;

  if (!cacheKey) {
    return null;
  }

  const cachedEntry = photoEditorSubjectMaskImageCache.get(cacheKey);

  if (cachedEntry) {
    return cachedEntry;
  }

  const image = new Image();
  let resolveLoad = null;
  const entry = {
    image,
    loaded: false,
    failed: false,
    loadPromise: new Promise((resolve) => {
      resolveLoad = resolve;
    }),
  };

  image.onload = () => {
    entry.loaded = true;
    entry.failed = false;
    resolveLoad?.(entry);
    if (photoEditorState) {
      schedulePhotoEditorRender();
    }
  };
  image.onerror = () => {
    entry.failed = true;
    resolveLoad?.(entry);
  };
  image.decoding = 'async';
  image.src = cacheKey;
  photoEditorSubjectMaskImageCache.set(cacheKey, entry);

  return entry;
}

function normalizePhotoEditorMaskCanvasAlpha(maskCanvas, { invert = false } = {}) {
  const ctx = maskCanvas?.getContext('2d', { willReadFrequently: true });

  if (!ctx || maskCanvas.width <= 0 || maskCanvas.height <= 0) {
    return maskCanvas;
  }

  let imageData;

  try {
    imageData = ctx.getImageData(0, 0, maskCanvas.width, maskCanvas.height);
  } catch {
    return maskCanvas;
  }

  const data = imageData.data;

  for (let index = 0; index < data.length; index += 4) {
    const alpha = data[index + 3];
    const luminance = 0.2126 * data[index] + 0.7152 * data[index + 1] + 0.0722 * data[index + 2];
    const maskValue = alpha < 255 ? alpha : luminance;
    const nextAlpha = invert ? 255 - maskValue : maskValue;
    data[index] = 255;
    data[index + 1] = 255;
    data[index + 2] = 255;
    data[index + 3] = clampColorChannel(nextAlpha);
  }

  ctx.putImageData(imageData, 0, 0);
  return maskCanvas;
}

function applyPhotoEditorSubjectMaskExpansion(maskCanvas, expand) {
  const radius = Math.round(Math.abs(expand));

  if (!maskCanvas || radius <= 0) {
    return maskCanvas;
  }

  const tempCanvas = document.createElement('canvas');
  tempCanvas.width = maskCanvas.width;
  tempCanvas.height = maskCanvas.height;
  const tempCtx = tempCanvas.getContext('2d');

  if (!tempCtx) {
    return maskCanvas;
  }

  if (expand > 0) {
    for (let offsetX = -radius; offsetX <= radius; offsetX += radius) {
      for (let offsetY = -radius; offsetY <= radius; offsetY += radius) {
        if (Math.hypot(offsetX, offsetY) <= radius * 1.42) {
          tempCtx.drawImage(maskCanvas, offsetX, offsetY);
        }
      }
    }
    tempCtx.drawImage(maskCanvas, 0, 0);
  } else {
    tempCtx.filter = `blur(${radius}px)`;
    tempCtx.drawImage(maskCanvas, 0, 0);
    tempCtx.filter = 'none';

    try {
      const imageData = tempCtx.getImageData(0, 0, tempCanvas.width, tempCanvas.height);
      const data = imageData.data;
      const threshold = Math.min(252, 130 + radius * 3.2);

      for (let index = 0; index < data.length; index += 4) {
        data[index + 3] = data[index + 3] >= threshold ? 255 : 0;
      }

      tempCtx.putImageData(imageData, 0, 0);
    } catch {
      return maskCanvas;
    }
  }

  return tempCanvas;
}

function applyPhotoEditorSubjectMaskFeather(maskCanvas, feather) {
  const radius = Math.round(feather);

  if (!maskCanvas || radius <= 0) {
    return maskCanvas;
  }

  const tempCanvas = document.createElement('canvas');
  tempCanvas.width = maskCanvas.width;
  tempCanvas.height = maskCanvas.height;
  const tempCtx = tempCanvas.getContext('2d');

  if (!tempCtx) {
    return maskCanvas;
  }

  tempCtx.filter = `blur(${radius}px)`;
  tempCtx.drawImage(maskCanvas, 0, 0);
  tempCtx.filter = 'none';
  return tempCanvas;
}

function createPhotoEditorSubjectMaskOutputCanvas(outputSize, sourceRect) {
  const subjectMask = normalizePhotoEditorSubjectMaskState(
    photoEditorState?.subjectMask
  );

  if (!subjectMask.enabled || !subjectMask.maskDataUrl || !outputSize || !sourceRect) {
    return null;
  }

  const cacheEntry = getPhotoEditorSubjectMaskImageCacheEntry(subjectMask);

  if (!cacheEntry?.loaded || cacheEntry.failed) {
    return null;
  }

  const maskCanvas = document.createElement('canvas');
  maskCanvas.width = outputSize.width;
  maskCanvas.height = outputSize.height;
  const maskCtx = maskCanvas.getContext('2d', { willReadFrequently: true });

  if (!maskCtx) {
    return null;
  }

  maskCtx.clearRect(0, 0, maskCanvas.width, maskCanvas.height);
  drawPhotoEditorCroppedSourceToCanvas(
    maskCtx,
    cacheEntry.image,
    sourceRect,
    outputSize,
    photoEditorState?.crop || getDefaultPhotoEditorCropState()
  );
  normalizePhotoEditorMaskCanvasAlpha(maskCanvas, { invert: subjectMask.invert });

  const expandedMask = applyPhotoEditorSubjectMaskExpansion(
    maskCanvas,
    subjectMask.expand
  );
  return applyPhotoEditorSubjectMaskFeather(expandedMask, subjectMask.feather);
}

function drawPhotoEditorCanvasWithSubjectMask(
  ctx,
  sourceCanvas,
  outputSize,
  sourceRect,
  maskMode = 'subject-only'
) {
  const normalizedMode = normalizePhotoEditorInternalMaskMode(maskMode);

  if (!sourceCanvas || normalizedMode === 'normal') {
    ctx.drawImage(sourceCanvas, 0, 0);
    return;
  }

  const maskCanvas = createPhotoEditorSubjectMaskOutputCanvas(outputSize, sourceRect);

  if (!maskCanvas) {
    return;
  }

  const layerCanvas = document.createElement('canvas');
  layerCanvas.width = outputSize.width;
  layerCanvas.height = outputSize.height;
  const layerCtx = layerCanvas.getContext('2d');

  if (!layerCtx) {
    ctx.drawImage(sourceCanvas, 0, 0);
    return;
  }

  layerCtx.drawImage(sourceCanvas, 0, 0);
  layerCtx.globalCompositeOperation =
    normalizedMode === 'background-only' ? 'destination-out' : 'destination-in';
  layerCtx.drawImage(maskCanvas, 0, 0);
  layerCtx.globalCompositeOperation = 'source-over';
  ctx.drawImage(layerCanvas, 0, 0);
}

function withPhotoEditorSubjectMaskLayer(
  ctx,
  outputSize,
  sourceRect,
  maskMode,
  draw
) {
  const normalizedMode = normalizePhotoEditorInternalMaskMode(maskMode);

  if (normalizedMode === 'normal' || !hasPhotoEditorReadySubjectMask()) {
    draw(ctx);
    return;
  }

  const layerCanvas = document.createElement('canvas');
  layerCanvas.width = outputSize.width;
  layerCanvas.height = outputSize.height;
  const layerCtx = layerCanvas.getContext('2d');

  if (!layerCtx) {
    draw(ctx);
    return;
  }

  draw(layerCtx);
  drawPhotoEditorCanvasWithSubjectMask(
    ctx,
    layerCanvas,
    outputSize,
    sourceRect,
    normalizedMode === 'background-only' ? 'background-only' : 'subject-only'
  );
}

function applyPhotoEditorTargetedAdjustmentsToCanvas(
  ctx,
  outputSize,
  sourceRect,
  { isInteractivePreview = false } = {}
) {
  const target = normalizePhotoEditorAdjustmentTarget(
    photoEditorState.adjustmentTarget
  );

  if (target === 'whole' || !hasPhotoEditorReadySubjectMask()) {
    applyPhotoEditorAdjustmentsToCanvas(
      ctx,
      outputSize.width,
      outputSize.height,
      photoEditorState.values,
      photoEditorState.curve
    );
    return;
  }

  const adjustedCanvas = document.createElement('canvas');
  adjustedCanvas.width = outputSize.width;
  adjustedCanvas.height = outputSize.height;
  const adjustedCtx = adjustedCanvas.getContext('2d', { willReadFrequently: true });

  if (!adjustedCtx) {
    return;
  }

  adjustedCtx.drawImage(ctx.canvas, 0, 0);
  applyPhotoEditorAdjustmentsToCanvas(
    adjustedCtx,
    outputSize.width,
    outputSize.height,
    photoEditorState.values,
    photoEditorState.curve
  );
  drawPhotoEditorCanvasWithSubjectMask(
    ctx,
    adjustedCanvas,
    outputSize,
    sourceRect,
    target === 'background' ? 'background-only' : 'subject-only'
  );

  if (isInteractivePreview) {
    ctx.canvas.dataset.subjectAdjustmentPreview = '1';
  }
}

function drawPhotoEditorSubjectMaskPreview(ctx, outputSize, sourceRect) {
  const subjectMask = normalizePhotoEditorSubjectMaskState(
    photoEditorState?.subjectMask
  );

  if (!subjectMask.enabled || !subjectMask.showOverlay) {
    return;
  }

  const maskCanvas = createPhotoEditorSubjectMaskOutputCanvas(outputSize, sourceRect);

  if (!maskCanvas) {
    return;
  }

  const overlayCanvas = document.createElement('canvas');
  overlayCanvas.width = outputSize.width;
  overlayCanvas.height = outputSize.height;
  const overlayCtx = overlayCanvas.getContext('2d');

  if (!overlayCtx) {
    return;
  }

  overlayCtx.fillStyle = 'rgba(79, 140, 255, 1)';
  overlayCtx.fillRect(0, 0, overlayCanvas.width, overlayCanvas.height);
  overlayCtx.globalCompositeOperation = 'destination-in';
  overlayCtx.drawImage(maskCanvas, 0, 0);
  overlayCtx.globalCompositeOperation = 'source-over';

  ctx.save();
  ctx.globalAlpha = subjectMask.opacity;
  ctx.drawImage(overlayCanvas, 0, 0);
  ctx.restore();
}

function drawPhotoEditorRadialBlurControl(ctx, width, height) {
  if (
    photoEditorState?.blur?.mode !== 'radial' ||
    photoEditorState.blur.isConfirmed ||
    !isPhotoEditorAccordionOpen('blur') ||
    photoEditorState.maskTool !== 'none' ||
    photoEditorState.draftMask
  ) {
    return;
  }

  const { centerX, centerY, radius, outerRadius } = getPhotoEditorBlurGeometry(
    width,
    height,
    photoEditorState.blur
  );
  const lineWidth = Math.max(2, Math.round(Math.max(width, height) / 900));
  const handleRadius = Math.max(7, lineWidth * 3);

  ctx.save();
  ctx.fillStyle = 'rgba(79, 140, 255, 0.1)';
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.9)';
  ctx.lineWidth = lineWidth;
  ctx.setLineDash([10, 7]);
  ctx.beginPath();
  ctx.arc(centerX, centerY, outerRadius, 0, Math.PI * 2);
  ctx.fill();
  ctx.stroke();
  ctx.setLineDash([5, 6]);
  ctx.beginPath();
  ctx.arc(centerX, centerY, radius, 0, Math.PI * 2);
  ctx.stroke();
  ctx.setLineDash([]);
  ctx.fillStyle = 'rgba(79, 140, 255, 0.92)';
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.95)';
  ctx.lineWidth = Math.max(2, Math.round(lineWidth * 0.8));
  ctx.beginPath();
  ctx.arc(centerX, centerY, handleRadius, 0, Math.PI * 2);
  ctx.fill();
  ctx.stroke();
  ctx.beginPath();
  ctx.arc(centerX + radius, centerY, handleRadius, 0, Math.PI * 2);
  ctx.fill();
  ctx.stroke();
  ctx.beginPath();
  ctx.arc(centerX + outerRadius, centerY, handleRadius, 0, Math.PI * 2);
  ctx.fill();
  ctx.stroke();
  ctx.restore();
}

function getPhotoEditorImageSize(image) {
  return {
    width: Number(image?.naturalWidth || image?.width) || 0,
    height: Number(image?.naturalHeight || image?.height) || 0,
  };
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

function mapPhotoEditorMaskPointToCanvas(
  point,
  mask,
  width,
  height,
  sourceRect = null,
  sourceImage = null
) {
  const allowOutside = canPhotoEditorMaskExtendOutside(mask);
  const bounds = getPhotoEditorMaskCoordinateBounds(allowOutside);
  const x = clampNumber(point?.x, bounds.min, bounds.max, 0);
  const y = clampNumber(point?.y, bounds.min, bounds.max, 0);
  const imageSize = getPhotoEditorImageSize(sourceImage);
  const canUseSourceSpace =
    mask?.space === 'source' &&
    sourceRect &&
    sourceRect.width > 0 &&
    sourceRect.height > 0 &&
    imageSize.width > 0 &&
    imageSize.height > 0;

  if (!canUseSourceSpace) {
    return {
      x: x * width,
      y: y * height,
    };
  }

  const geometry = getPhotoEditorCropRenderGeometry(
    sourceRect,
    width,
    height,
    photoEditorState?.crop || getDefaultPhotoEditorCropState()
  );
  const localX =
    ((x * imageSize.width - sourceRect.x) / sourceRect.width - 0.5) *
    geometry.drawWidth *
    geometry.coverScale;
  const localY =
    ((y * imageSize.height - sourceRect.y) / sourceRect.height - 0.5) *
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

function getPhotoEditorMaskCanvasPoints(
  mask,
  width,
  height,
  sourceRect = null,
  sourceImage = null
) {
  return (Array.isArray(mask?.points) ? mask.points : [])
    .filter((point) => Number.isFinite(point?.x) && Number.isFinite(point?.y))
    .map((point) =>
      mapPhotoEditorMaskPointToCanvas(
        point,
        mask,
        width,
        height,
        sourceRect,
        sourceImage
      )
    );
}

function getRawCanvasRectFromMask(
  mask,
  width,
  height,
  sourceRect = null,
  sourceImage = null
) {
  if (mask?.shape === 'freehand') {
    const points = getPhotoEditorMaskCanvasPoints(
      mask,
      width,
      height,
      sourceRect,
      sourceImage
    );

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
      sourceImage
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

function getVisibleCanvasRectFromMask(
  mask,
  width,
  height,
  sourceRect = null,
  sourceImage = null
) {
  const maskRect = getRawCanvasRectFromMask(
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );

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

function isPhotoEditorPointInsideMask(point, mask, width, height, sourceRect, sourceImage) {
  if (!point || !mask || !photoEditorCanvas) {
    return false;
  }

  const pointX = clampNumber(point.x, 0, 1, 0) * width;
  const pointY = clampNumber(point.y, 0, 1, 0) * height;
  const ctx = photoEditorCanvas.getContext('2d');

  if (!ctx) {
    const maskRect = getRawCanvasRectFromMask(
      mask,
      width,
      height,
      sourceRect,
      sourceImage
    );

    return Boolean(
      maskRect &&
        pointX >= maskRect.x &&
        pointX <= maskRect.x + maskRect.width &&
        pointY >= maskRect.y &&
        pointY <= maskRect.y + maskRect.height
    );
  }

  ctx.save();
  const pathRect = addPhotoEditorMaskPath(
    ctx,
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );
  const isInside = Boolean(pathRect && ctx.isPointInPath(pointX, pointY));
  ctx.restore();
  return isInside;
}

function addPhotoEditorMaskPath(
  ctx,
  mask,
  width,
  height,
  sourceRect = null,
  sourceImage = null
) {
  const maskRect = getRawCanvasRectFromMask(
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );

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
    const points = getPhotoEditorMaskCanvasPoints(
      mask,
      width,
      height,
      sourceRect,
      sourceImage
    );

    points.forEach((point, index) => {
      const { x, y } = point;

      if (index === 0) {
        ctx.moveTo(x, y);
      } else {
        ctx.lineTo(x, y);
      }
    });
    ctx.closePath();
    return maskRect;
  }

  ctx.rect(maskRect.x, maskRect.y, maskRect.width, maskRect.height);
  return maskRect;
}

function withPhotoEditorMaskClip(
  ctx,
  mask,
  width,
  height,
  sourceRect,
  sourceImage,
  draw
) {
  const maskRect = getVisibleCanvasRectFromMask(
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );

  if (!maskRect) {
    return;
  }

  ctx.save();
  const pathRect = addPhotoEditorMaskPath(
    ctx,
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );

  if (!pathRect) {
    ctx.restore();
    return;
  }

  ctx.clip();
  draw(maskRect);
  ctx.restore();
}

function applyPhotoEditorFillMask(
  ctx,
  mask,
  width,
  height,
  sourceRect,
  sourceImage
) {
  const strength = clampNumber(mask?.strength, 0, 100, 45);

  if (strength <= 0) {
    return;
  }

  withPhotoEditorMaskClip(ctx, mask, width, height, sourceRect, sourceImage, (maskRect) => {
    ctx.fillStyle = mask.color || '#111827';
    ctx.globalAlpha = strength / 100;
    ctx.fillRect(maskRect.x, maskRect.y, maskRect.width, maskRect.height);
    ctx.globalAlpha = 1;
  });
}

function applyPhotoEditorMosaicMask(
  ctx,
  mask,
  width,
  height,
  sourceRect,
  sourceImage
) {
  const strength = clampNumber(mask?.strength, 0, 100, 45);

  if (strength <= 0) {
    return;
  }

  const maskRect = getVisibleCanvasRectFromMask(
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );

  if (!maskRect) {
    return;
  }

  const tempCanvas = document.createElement('canvas');
  const mosaicDivisor = Math.max(8, 30 - strength * 0.22);
  const cellSize = Math.max(
    3,
    Math.round(Math.min(maskRect.width, maskRect.height) / mosaicDivisor)
  );
  tempCanvas.width = Math.max(1, Math.ceil(maskRect.width / cellSize));
  tempCanvas.height = Math.max(1, Math.ceil(maskRect.height / cellSize));
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
  const pathRect = addPhotoEditorMaskPath(
    ctx,
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );

  if (!pathRect) {
    ctx.restore();
    return;
  }

  ctx.clip();
  ctx.imageSmoothingEnabled = false;
  ctx.drawImage(
    tempCanvas,
    0,
    0,
    tempCanvas.width,
    tempCanvas.height,
    maskRect.x,
    maskRect.y,
    maskRect.width,
    maskRect.height
  );
  ctx.restore();
}

function applyPhotoEditorBlurMask(
  ctx,
  mask,
  width,
  height,
  sourceRect,
  sourceImage,
  { isInteractivePreview = false } = {}
) {
  const strength = clampNumber(mask?.strength, 0, 100, 45);

  if (strength <= 0) {
    return;
  }

  const maskRect = getVisibleCanvasRectFromMask(
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );

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
  const tempCanvas = document.createElement('canvas');
  tempCanvas.width = sourceWidth;
  tempCanvas.height = sourceHeight;
  const tempCtx = tempCanvas.getContext('2d');

  if (!tempCtx) {
    return;
  }

  tempCtx.drawImage(
    ctx.canvas,
    sourceX,
    sourceY,
    sourceWidth,
    sourceHeight,
    0,
    0,
    sourceWidth,
    sourceHeight
  );

  ctx.save();
  const pathRect = addPhotoEditorMaskPath(
    ctx,
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );

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
  sourceRect = null,
  sourceImage = null,
  { isInteractivePreview = false } = {}
) {
  for (const mask of Array.isArray(masks) ? masks : []) {
    if (mask.type === 'blur') {
      applyPhotoEditorBlurMask(ctx, mask, width, height, sourceRect, sourceImage, {
        isInteractivePreview,
      });
      continue;
    }

    if (mask.type === 'mosaic') {
      applyPhotoEditorMosaicMask(ctx, mask, width, height, sourceRect, sourceImage);
      continue;
    }

    applyPhotoEditorFillMask(ctx, mask, width, height, sourceRect, sourceImage);
  }
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
  const numericLetterSpacing = Number(textState.letterSpacing);
  const letterSpacing = Number.isFinite(numericLetterSpacing)
    ? numericLetterSpacing
    : 0;
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

function getPhotoEditorTextRenderScale(textScale = 1) {
  const numericScale = Number(textScale);
  return Number.isFinite(numericScale) && numericScale > 0 ? numericScale : 1;
}

function getPhotoEditorTextCanvasRenderScale(width, height, textScale = 1) {
  const canvasEdge = Math.max(Number(width) || 0, Number(height) || 0);
  const baseScale = getPhotoEditorTextRenderScale(textScale);

  if (canvasEdge <= 0) {
    return baseScale;
  }

  return clampNumber(
    (canvasEdge / PHOTO_EDITOR_TEXT_REFERENCE_EDGE) * baseScale,
    0.05,
    32,
    baseScale
  );
}

function getPhotoEditorTextRenderState(textOverlay, { textScale = 1 } = {}) {
  const textState = normalizePhotoEditorTextState(textOverlay);
  const scale = getPhotoEditorTextRenderScale(textScale);

  if (Math.abs(scale - 1) < 0.0001) {
    return textState;
  }

  return {
    ...textState,
    size: clampNumber(textState.size * scale, 1, 4096, textState.size),
    strokeWidth: clampNumber(
      textState.strokeWidth * scale,
      0,
      1024,
      textState.strokeWidth
    ),
    letterSpacing: clampNumber(
      textState.letterSpacing * scale,
      -1024,
      4096,
      textState.letterSpacing
    ),
  };
}

function getPhotoEditorTextMetricsForState(ctx, textState) {
  const lines = textState.enabled ? getPhotoEditorTextLines(textState.text) : [];

  if (lines.length === 0) {
    return null;
  }

  const size = clampNumber(textState.size, 1, 4096, 64);
  const lineHeight = size * 1.18;
  const width = Math.max(
    1,
    ...lines.map((line) =>
      measurePhotoEditorTextLineWidth(ctx, line, textState.letterSpacing)
    )
  );
  const height = Math.max(size, lineHeight * lines.length);

  return {
    lines,
    width,
    height,
    size,
    lineHeight,
    startY: -(lineHeight * (lines.length - 1)) / 2,
  };
}

function getPhotoEditorTextMetrics(ctx, textOverlay, options = {}) {
  return getPhotoEditorTextMetricsForState(
    ctx,
    getPhotoEditorTextRenderState(textOverlay, options)
  );
}

function drawPhotoEditorSingleTextOverlay(
  ctx,
  width,
  height,
  textOverlay,
  { textScale = 1 } = {}
) {
  const textState = getPhotoEditorTextRenderState(textOverlay, {
    textScale: getPhotoEditorTextCanvasRenderScale(width, height, textScale),
  });

  if (!textState.enabled || !textState.text.trim()) {
    return;
  }

  const x = clampNumber(textState.x, 0, 1, 0.5) * width;
  const y = clampNumber(textState.y, 0, 1, 0.5) * height;
  const size = clampNumber(textState.size, 1, 4096, 64);
  const strokeWidth = clampNumber(textState.strokeWidth, 0, 1024, 4);

  ctx.save();
  ctx.font = `${textState.weight} ${size}px ${textState.fontFamily}`;
  const metrics = getPhotoEditorTextMetricsForState(ctx, textState);

  if (!metrics) {
    ctx.restore();
    return;
  }

  ctx.translate(x, y);
  ctx.rotate(textState.rotation * Math.PI / 180);
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

  for (let lineIndex = 0; lineIndex < metrics.lines.length; lineIndex += 1) {
    drawPhotoEditorTextLine(
      ctx,
      metrics.lines[lineIndex],
      0,
      metrics.startY + metrics.lineHeight * lineIndex,
      textState
    );
  }

  ctx.restore();
}

function applyPhotoEditorTextOverlayToCanvas(
  ctx,
  width,
  height,
  textOverlays,
  { textScale = 1, maskMode = null } = {}
) {
  const overlays = Array.isArray(textOverlays)
    ? textOverlays
    : textOverlays
      ? [textOverlays]
      : [];
  const targetMaskMode = maskMode
    ? normalizePhotoEditorInternalMaskMode(maskMode)
    : null;

  for (const textOverlay of normalizePhotoEditorTextOverlays(overlays)) {
    if (
      targetMaskMode &&
      normalizePhotoEditorMaskMode(textOverlay.maskMode) !== targetMaskMode
    ) {
      continue;
    }

    drawPhotoEditorSingleTextOverlay(ctx, width, height, textOverlay, {
      textScale,
    });
  }
}

function drawPhotoEditorSingleImageOverlay(ctx, width, height, overlay) {
  const imageOverlay = normalizePhotoEditorImageOverlayState(overlay);
  const cacheEntry = getPhotoEditorOverlayImageCacheEntry(imageOverlay);

  if (!cacheEntry?.loaded || cacheEntry.failed) {
    return false;
  }

  const image = cacheEntry.image;
  const centerX = imageOverlay.x * width;
  const centerY = imageOverlay.y * height;
  const drawWidth = Math.max(
    1,
    Math.abs(imageOverlay.width) * width
  );
  const drawHeight = Math.max(
    1,
    Math.abs(imageOverlay.height) * height
  );

  ctx.save();
  ctx.globalAlpha = clampNumber(imageOverlay.opacity, 0, 1, 1);
  ctx.globalCompositeOperation = normalizePhotoEditorImageOverlayBlendMode(
    imageOverlay.blendMode
  );
  ctx.drawImage(
    image,
    centerX - drawWidth / 2,
    centerY - drawHeight / 2,
    drawWidth,
    drawHeight
  );
  ctx.restore();
  return true;
}

function drawPhotoEditorSingleMaskedImageOverlay(
  ctx,
  width,
  height,
  overlay,
  outputSize,
  sourceRect,
  maskMode
) {
  const imageOverlay = normalizePhotoEditorImageOverlayState(overlay);
  const cacheEntry = getPhotoEditorOverlayImageCacheEntry(imageOverlay);

  if (!cacheEntry?.loaded || cacheEntry.failed || !outputSize || !sourceRect) {
    return false;
  }

  const maskCanvas = createPhotoEditorSubjectMaskOutputCanvas(outputSize, sourceRect);

  if (!maskCanvas) {
    return false;
  }

  const image = cacheEntry.image;
  const centerX = imageOverlay.x * width;
  const centerY = imageOverlay.y * height;
  const drawWidth = Math.max(1, Math.abs(imageOverlay.width) * width);
  const drawHeight = Math.max(1, Math.abs(imageOverlay.height) * height);
  const layerCanvas = document.createElement('canvas');
  layerCanvas.width = width;
  layerCanvas.height = height;
  const layerCtx = layerCanvas.getContext('2d');

  if (!layerCtx) {
    return false;
  }

  layerCtx.save();
  layerCtx.globalAlpha = clampNumber(imageOverlay.opacity, 0, 1, 1);
  layerCtx.drawImage(
    image,
    centerX - drawWidth / 2,
    centerY - drawHeight / 2,
    drawWidth,
    drawHeight
  );
  layerCtx.restore();
  layerCtx.globalCompositeOperation =
    normalizePhotoEditorInternalMaskMode(maskMode) === 'background-only'
      ? 'destination-out'
      : 'destination-in';
  layerCtx.drawImage(maskCanvas, 0, 0);
  layerCtx.globalCompositeOperation = 'source-over';

  ctx.save();
  ctx.globalCompositeOperation = normalizePhotoEditorImageOverlayBlendMode(
    imageOverlay.blendMode
  );
  ctx.drawImage(layerCanvas, 0, 0);
  ctx.restore();
  return true;
}

function applyPhotoEditorImageOverlayToCanvas(
  ctx,
  width,
  height,
  imageOverlays,
  { maskMode = null, outputSize = null, sourceRect = null } = {}
) {
  const overlays = normalizePhotoEditorImageOverlays(imageOverlays);
  const targetMaskMode = maskMode
    ? normalizePhotoEditorInternalMaskMode(maskMode)
    : null;

  for (const overlay of overlays) {
    if (
      targetMaskMode &&
      normalizePhotoEditorInternalMaskMode(overlay.maskMode) !== targetMaskMode
    ) {
      continue;
    }

    if (
      targetMaskMode &&
      targetMaskMode !== 'normal' &&
      outputSize &&
      sourceRect &&
      hasPhotoEditorReadySubjectMask()
    ) {
      drawPhotoEditorSingleMaskedImageOverlay(
        ctx,
        width,
        height,
        overlay,
        outputSize,
        sourceRect,
        targetMaskMode
      );
      continue;
    }

    drawPhotoEditorSingleImageOverlay(ctx, width, height, overlay);
  }
}

function getPhotoEditorImageOverlayCanvasMetrics(width, height, overlay) {
  const imageOverlay = normalizePhotoEditorImageOverlayState(overlay);

  if (!imageOverlay.fileUrl) {
    return null;
  }

  return {
    overlay: imageOverlay,
    centerX: imageOverlay.x * width,
    centerY: imageOverlay.y * height,
    width: Math.max(1, imageOverlay.width * width),
    height: Math.max(1, imageOverlay.height * height),
  };
}

function getPhotoEditorImageOverlayHandles(metrics) {
  if (!metrics) {
    return null;
  }

  return {
    resize: {
      x: metrics.centerX + metrics.width / 2,
      y: metrics.centerY + metrics.height / 2,
    },
  };
}

function drawPhotoEditorImageOverlayControls(ctx, width, height) {
  if (
    !photoEditorState ||
    !isPhotoEditorAccordionOpen('imageOverlay') ||
    photoEditorState.maskTool !== 'none' ||
    photoEditorState.draftMask
  ) {
    return;
  }

  const activeOverlay = getPhotoEditorActiveImageOverlay();
  const metrics = getPhotoEditorImageOverlayCanvasMetrics(
    width,
    height,
    activeOverlay
  );

  if (!metrics) {
    return;
  }

  const handles = getPhotoEditorImageOverlayHandles(metrics);
  const handleSize = Math.max(
    12,
    Math.min(24, Math.min(metrics.width, metrics.height) * 0.18)
  );

  ctx.save();
  ctx.fillStyle = 'rgba(79, 140, 255, 0.08)';
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.94)';
  ctx.lineWidth = Math.max(2, Math.round(Math.max(width, height) / 950));
  ctx.setLineDash([8, 6]);
  ctx.fillRect(
    metrics.centerX - metrics.width / 2,
    metrics.centerY - metrics.height / 2,
    metrics.width,
    metrics.height
  );
  ctx.strokeRect(
    metrics.centerX - metrics.width / 2,
    metrics.centerY - metrics.height / 2,
    metrics.width,
    metrics.height
  );
  ctx.setLineDash([]);

  if (handles?.resize) {
    ctx.shadowColor = 'rgba(0, 0, 0, 0.42)';
    ctx.shadowBlur = Math.max(4, handleSize * 0.45);
    ctx.shadowOffsetY = Math.max(1, handleSize * 0.16);
    ctx.fillStyle = 'rgba(255, 255, 255, 0.96)';
    ctx.strokeStyle = 'rgba(79, 140, 255, 0.96)';
    ctx.beginPath();
    ctx.rect(
      handles.resize.x - handleSize / 2,
      handles.resize.y - handleSize / 2,
      handleSize,
      handleSize
    );
    ctx.fill();
    ctx.stroke();
  }

  ctx.restore();
}

function getPhotoEditorTextCanvasMetrics(ctx, width, height, textOverlay) {
  const textState = getPhotoEditorTextRenderState(textOverlay, {
    textScale: getPhotoEditorTextCanvasRenderScale(width, height),
  });

  if (!textState.enabled || !textState.text.trim()) {
    return null;
  }

  ctx.save();
  ctx.font = `${textState.weight} ${Math.round(textState.size)}px ${textState.fontFamily}`;
  const metrics = getPhotoEditorTextMetricsForState(ctx, textState);
  ctx.restore();

  if (!metrics) {
    return null;
  }

  const padding = Math.max(10, metrics.size * 0.12);

  return {
    textState,
    centerX: clampNumber(textState.x, 0, 1, 0.5) * width,
    centerY: clampNumber(textState.y, 0, 1, 0.5) * height,
    width: metrics.width + padding * 2,
    height: metrics.height + padding * 2,
    rotationRadians: textState.rotation * Math.PI / 180,
    handleOffset: Math.max(30, metrics.size * 0.52),
  };
}

function rotatePhotoEditorCanvasPoint(point, center, rotationRadians) {
  const cos = Math.cos(rotationRadians);
  const sin = Math.sin(rotationRadians);
  const deltaX = point.x - center.x;
  const deltaY = point.y - center.y;

  return {
    x: center.x + deltaX * cos - deltaY * sin,
    y: center.y + deltaX * sin + deltaY * cos,
  };
}

function getPhotoEditorTextHandles(metrics) {
  if (!metrics) {
    return null;
  }

  const center = {
    x: metrics.centerX,
    y: metrics.centerY,
  };
  const top = rotatePhotoEditorCanvasPoint(
    {
      x: metrics.centerX,
      y: metrics.centerY - metrics.height / 2,
    },
    center,
    metrics.rotationRadians
  );
  const rotate = rotatePhotoEditorCanvasPoint(
    {
      x: metrics.centerX,
      y: metrics.centerY - metrics.height / 2 - metrics.handleOffset,
    },
    center,
    metrics.rotationRadians
  );

  return {
    top,
    rotate,
  };
}

function getPhotoEditorSnapTargets(axis, { includeCenter = true } = {}) {
  const targets = includeCenter ? [0.5] : [];
  const guides = normalizePhotoEditorRulerGuides(photoEditorState?.rulerGuides);

  if (photoEditorState?.showRuleOfThirdsGrid && (axis === 'x' || axis === 'y')) {
    targets.push(1 / 3, 2 / 3);
  }

  if (photoEditorState?.showRulers && (axis === 'x' || axis === 'y')) {
    targets.push(...guides[axis]);
  }

  return targets
    .map((value) => clampNumber(value, 0, 1, 0.5))
    .sort((left, right) => left - right)
    .filter((value, index, values) => index === 0 || Math.abs(value - values[index - 1]) > 0.001);
}

function getPhotoEditorAxisSnap(value, axis, { includeCenter = true } = {}) {
  const numericValue = Number(value);
  const originalValue = Number.isFinite(numericValue) ? numericValue : 0.5;
  const targets = getPhotoEditorSnapTargets(axis, { includeCenter });
  let bestTarget = null;
  let bestDistance = Infinity;

  for (const target of targets) {
    const distance = Math.abs(originalValue - target);

    if (distance < bestDistance) {
      bestTarget = target;
      bestDistance = distance;
    }
  }

  if (
    bestTarget !== null &&
    bestDistance <= PHOTO_EDITOR_TEXT_CENTER_SNAP_THRESHOLD
  ) {
    return {
      value: bestTarget,
      originalValue,
      target: bestTarget,
      distance: bestDistance,
      snapped: true,
    };
  }

  return {
    value: originalValue,
    originalValue,
    target: null,
    distance: bestDistance,
    snapped: false,
  };
}

function getPhotoEditorSnapGuideFromAxisSnaps(snapX, snapY) {
  const guide = {};

  if (snapX?.snapped) {
    guide.x = snapX.target;
  }

  if (snapY?.snapped) {
    guide.y = snapY.target;
  }

  return guide.x !== undefined || guide.y !== undefined ? guide : null;
}

function setPhotoEditorSnapGuideFromAxisSnaps(snapX, snapY) {
  if (!photoEditorState) {
    return;
  }

  photoEditorState.snapGuide = getPhotoEditorSnapGuideFromAxisSnaps(
    snapX,
    snapY
  );
}

function getPhotoEditorTextSnapCandidates(textOverlay) {
  const textState = normalizePhotoEditorTextState(textOverlay);
  const width = Math.max(1, Number(photoEditorCanvas?.width) || 1);
  const height = Math.max(1, Number(photoEditorCanvas?.height) || 1);
  const ctx = photoEditorCanvas?.getContext('2d');

  if (!ctx) {
    return {
      x: [textState.x],
      y: [textState.y],
    };
  }

  const metrics = getPhotoEditorTextCanvasMetrics(
    ctx,
    width,
    height,
    textState
  );

  if (!metrics) {
    return {
      x: [textState.x],
      y: [textState.y],
    };
  }

  const center = {
    x: metrics.centerX,
    y: metrics.centerY,
  };
  const halfWidth = metrics.width / 2;
  const halfHeight = metrics.height / 2;
  const corners = [
    { x: -halfWidth, y: -halfHeight },
    { x: halfWidth, y: -halfHeight },
    { x: halfWidth, y: halfHeight },
    { x: -halfWidth, y: halfHeight },
  ].map((corner) =>
    rotatePhotoEditorCanvasPoint(
      {
        x: center.x + corner.x,
        y: center.y + corner.y,
      },
      center,
      metrics.rotationRadians
    )
  );
  const xs = [center.x, ...corners.map((corner) => corner.x)];
  const ys = [center.y, ...corners.map((corner) => corner.y)];

  return {
    x: [
      center.x / width,
      Math.min(...xs) / width,
      Math.max(...xs) / width,
    ],
    y: [
      center.y / height,
      Math.min(...ys) / height,
      Math.max(...ys) / height,
    ],
  };
}

function snapPhotoEditorTextToGuides(textOverlay) {
  const textState = normalizePhotoEditorTextState(textOverlay);
  const candidates = getPhotoEditorTextSnapCandidates(textState);
  const snapX = getPhotoEditorBestAxisSnap(candidates.x, 'x');
  const snapY = getPhotoEditorBestAxisSnap(candidates.y, 'y');

  setPhotoEditorSnapGuideFromAxisSnaps(snapX, snapY);

  if (!snapX && !snapY) {
    return textState;
  }

  return normalizePhotoEditorTextState({
    ...textState,
    x: textState.x + (snapX ? snapX.value - snapX.originalValue : 0),
    y: textState.y + (snapY ? snapY.value - snapY.originalValue : 0),
  });
}

function drawPhotoEditorSnapGuides(ctx, width, height) {
  const guide = photoEditorState?.snapGuide;

  if (guide?.x === undefined && guide?.y === undefined) {
    return;
  }

  const edge = Math.max(width, height);
  const lineWidth = Math.max(2, Math.round(edge / 900));

  ctx.save();
  ctx.strokeStyle = 'rgba(125, 211, 252, 0.88)';
  ctx.lineWidth = lineWidth;
  ctx.setLineDash([Math.max(10, lineWidth * 5), Math.max(8, lineWidth * 4)]);

  if (guide.x !== undefined) {
    const x = clampNumber(guide.x, 0, 1, 0.5) * width;
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, height);
    ctx.stroke();
  }

  if (guide.y !== undefined) {
    const y = clampNumber(guide.y, 0, 1, 0.5) * height;
    ctx.beginPath();
    ctx.moveTo(0, y);
    ctx.lineTo(width, y);
    ctx.stroke();
  }

  ctx.restore();
}

function drawPhotoEditorTextControls(ctx, width, height) {
  if (
    !photoEditorState ||
    !isPhotoEditorAccordionOpen('text') ||
    photoEditorState.maskTool !== 'none' ||
    photoEditorState.draftMask
  ) {
    return;
  }

  const activeText = getPhotoEditorActiveTextOverlay();
  const metrics = getPhotoEditorTextCanvasMetrics(ctx, width, height, activeText);

  if (!metrics) {
    return;
  }

  const handles = getPhotoEditorTextHandles(metrics);
  const handleRadius = Math.max(8, Math.min(15, metrics.height * 0.18));

  ctx.save();
  ctx.translate(metrics.centerX, metrics.centerY);
  ctx.rotate(metrics.rotationRadians);
  ctx.fillStyle = 'rgba(79, 140, 255, 0.08)';
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.92)';
  ctx.lineWidth = Math.max(2, Math.round(Math.max(width, height) / 950));
  ctx.setLineDash([8, 6]);
  ctx.fillRect(-metrics.width / 2, -metrics.height / 2, metrics.width, metrics.height);
  ctx.strokeRect(-metrics.width / 2, -metrics.height / 2, metrics.width, metrics.height);
  ctx.restore();

  if (handles) {
    ctx.save();
    ctx.strokeStyle = 'rgba(255, 255, 255, 0.9)';
    ctx.fillStyle = 'rgba(245, 158, 11, 0.96)';
    ctx.lineWidth = Math.max(2, Math.round(Math.max(width, height) / 900));
    ctx.beginPath();
    ctx.moveTo(handles.top.x, handles.top.y);
    ctx.lineTo(handles.rotate.x, handles.rotate.y);
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(handles.rotate.x, handles.rotate.y, handleRadius, 0, Math.PI * 2);
    ctx.fill();
    ctx.stroke();
    ctx.restore();
  }
}

function drawPhotoEditorDraftMask(ctx, width, height, sourceRect = null, sourceImage = null) {
  const draftMask = photoEditorState?.draftMask;

  if (!draftMask) {
    return;
  }

  ctx.save();
  ctx.fillStyle = 'rgba(79, 140, 255, 0.16)';
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.88)';
  ctx.lineWidth = Math.max(2, Math.round(Math.max(width, height) / 800));
  ctx.setLineDash([8, 6]);
  const pathRect = addPhotoEditorMaskPath(
    ctx,
    draftMask,
    width,
    height,
    sourceRect,
    sourceImage
  );

  if (!pathRect) {
    ctx.restore();
    return;
  }

  ctx.fill();
  ctx.stroke();
  ctx.setLineDash([]);
  const handles = getPhotoEditorDraftMaskHandles(width, height, sourceRect, sourceImage);
  const handleRadius = Math.max(9, Math.round(Math.max(width, height) / 118));
  const resizeHandleSize = handleRadius * 1.55;
  ctx.fillStyle = 'rgba(79, 140, 255, 0.94)';
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.95)';
  ctx.lineWidth = Math.max(2, Math.round(Math.max(width, height) / 900));

  if (handles) {
    ctx.beginPath();
    ctx.moveTo(handles.top.x, handles.top.y);
    ctx.lineTo(handles.rotate.x, handles.rotate.y);
    ctx.stroke();

    ctx.shadowColor = 'rgba(0, 0, 0, 0.42)';
    ctx.shadowBlur = Math.max(4, handleRadius * 0.55);
    ctx.shadowOffsetY = Math.max(1, handleRadius * 0.18);

    ctx.beginPath();
    ctx.arc(handles.move.x, handles.move.y, handleRadius, 0, Math.PI * 2);
    ctx.fill();
    ctx.stroke();

    ctx.save();
    ctx.translate(handles.resize.x, handles.resize.y);
    ctx.rotate(getPhotoEditorMaskRotation(draftMask) * Math.PI / 180);
    ctx.fillStyle = 'rgba(255, 255, 255, 0.95)';
    ctx.strokeStyle = 'rgba(79, 140, 255, 0.96)';
    ctx.beginPath();
    ctx.rect(
      -resizeHandleSize / 2,
      -resizeHandleSize / 2,
      resizeHandleSize,
      resizeHandleSize
    );
    ctx.fill();
    ctx.stroke();
    ctx.restore();

    ctx.fillStyle = 'rgba(245, 158, 11, 0.96)';
    ctx.strokeStyle = 'rgba(255, 255, 255, 0.96)';
    ctx.beginPath();
    ctx.arc(handles.rotate.x, handles.rotate.y, handleRadius, 0, Math.PI * 2);
    ctx.fill();
    ctx.stroke();
    ctx.shadowBlur = 0;
    ctx.shadowOffsetY = 0;
  }

  ctx.restore();
}

function clonePhotoEditorPlainValue(value) {
  if (value === null || value === undefined) {
    return value;
  }

  return JSON.parse(JSON.stringify(value));
}

function getPhotoEditorRenderPayload({
  outputSize,
  sourceRect,
  includeDraft = false,
  responseType = 'bitmap',
  exportSettings = null,
  isInteractivePreview = false,
} = {}) {
  if (!photoEditorState?.sourceImage || !outputSize || !sourceRect) {
    return null;
  }

  const sourceImageSize = getPhotoEditorImageSize(photoEditorState.sourceImage);

  return {
    responseType,
    width: outputSize.width,
    height: outputSize.height,
    values: clonePhotoEditValues(photoEditorState.values),
    curve: normalizePhotoEditorCurveState(photoEditorState.curve),
    crop: normalizePhotoEditorCropState(photoEditorState.crop),
    blur: normalizePhotoEditorBlurState(photoEditorState.blur),
    textOverlays: clonePhotoEditorPlainValue(
      getPhotoEditorTextCollectionFromState().textOverlays
    ),
    masks: clonePhotoEditorPlainValue(photoEditorState.masks || []),
    draftMask:
      includeDraft && photoEditorState.draftMask
        ? clonePhotoEditorPlainValue(photoEditorState.draftMask)
        : null,
    includeDraft,
    sourceRect: {
      x: sourceRect.x,
      y: sourceRect.y,
      width: sourceRect.width,
      height: sourceRect.height,
    },
    sourceImageSize,
    mimeType: exportSettings?.mimeType,
    quality: exportSettings?.quality,
    isInteractivePreview: Boolean(isInteractivePreview),
  };
}

function handlePhotoEditorWorkerMessage(event) {
  const message = event.data || {};

  if (message.type !== 'render-result') {
    return;
  }

  const request = photoEditorWorkerRequests.get(message.requestId);

  if (!request) {
    message.bitmap?.close?.();
    return;
  }

  clearTimeout(request.timer);
  photoEditorWorkerRequests.delete(message.requestId);

  if (!message.ok) {
    request.reject(new Error(message.error || '画像処理Workerで失敗しました'));
    return;
  }

  request.resolve(message);
}

function handlePhotoEditorWorkerError(error) {
  photoEditorWorkerUnavailable = true;

  for (const request of photoEditorWorkerRequests.values()) {
    clearTimeout(request.timer);
    request.reject(error instanceof Error ? error : new Error('画像処理Workerで失敗しました'));
  }

  photoEditorWorkerRequests.clear();
  photoEditorWorker?.terminate();
  photoEditorWorker = null;
}

function cancelPhotoEditorPreviewWorkerRequests(reason = '画像処理プレビューを更新しました') {
  let canceledPreviewCount = 0;
  let hasNonPreviewRequest = false;
  const cancelError = new Error(reason);

  for (const [requestId, request] of photoEditorWorkerRequests.entries()) {
    if (request.purpose === 'preview') {
      clearTimeout(request.timer);
      photoEditorWorkerRequests.delete(requestId);
      request.reject(cancelError);
      canceledPreviewCount += 1;
    } else {
      hasNonPreviewRequest = true;
    }
  }

  if (canceledPreviewCount > 0 && !hasNonPreviewRequest) {
    photoEditorWorker?.terminate();
    photoEditorWorker = null;
  }
}

function getPhotoEditorWorker() {
  if (
    photoEditorWorkerUnavailable ||
    typeof Worker !== 'function' ||
    typeof createImageBitmap !== 'function'
  ) {
    return null;
  }

  if (photoEditorWorker) {
    return photoEditorWorker;
  }

  try {
    const workerUrl = new URL('photo-editor-worker.js', window.location.href);
    photoEditorWorker = new Worker(workerUrl.href);
    photoEditorWorker.addEventListener('message', handlePhotoEditorWorkerMessage);
    photoEditorWorker.addEventListener('error', handlePhotoEditorWorkerError);
    return photoEditorWorker;
  } catch (error) {
    photoEditorWorkerUnavailable = true;
    console.warn('画像処理Workerを起動できませんでした', error);
    return null;
  }
}

function requestPhotoEditorWorkerRender(
  payload,
  transferList = [],
  { purpose = 'preview', cancelPendingPreview = false } = {}
) {
  if (cancelPendingPreview) {
    cancelPhotoEditorPreviewWorkerRequests();
  }

  const worker = getPhotoEditorWorker();

  if (!worker) {
    return Promise.reject(new Error('画像処理Workerを利用できません'));
  }

  const requestId = ++photoEditorWorkerRequestId;

  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      photoEditorWorkerRequests.delete(requestId);
      reject(new Error('画像処理Workerがタイムアウトしました'));
    }, 45000);

    photoEditorWorkerRequests.set(requestId, { resolve, reject, timer, purpose });

    try {
      worker.postMessage(
        {
          type: 'render',
          requestId,
          payload,
        },
        transferList
      );
    } catch (error) {
      clearTimeout(timer);
      photoEditorWorkerRequests.delete(requestId);
      reject(error);
    }
  });
}

async function applyPhotoEditorEffectsWithWorker(
  targetCanvas,
  renderBase,
  {
    includeDraft = false,
    responseType = 'bitmap',
    exportSettings = null,
    purpose = 'preview',
    cancelPendingPreview = false,
    isInteractivePreview = false,
  } = {}
) {
  if (!targetCanvas || !renderBase) {
    return null;
  }

  const payload = getPhotoEditorRenderPayload({
    outputSize: renderBase.outputSize,
    sourceRect: renderBase.sourceRect,
    includeDraft,
    responseType,
    exportSettings,
    isInteractivePreview,
  });

  if (!payload) {
    return null;
  }

  if (!getPhotoEditorWorker()) {
    return Promise.reject(new Error('画像処理Workerを利用できません'));
  }

  let sourceBitmap = null;

  try {
    sourceBitmap = await createImageBitmap(targetCanvas);
    payload.sourceBitmap = sourceBitmap;
    return await requestPhotoEditorWorkerRender(payload, [sourceBitmap], {
      purpose,
      cancelPendingPreview,
    });
  } catch (error) {
    try {
      sourceBitmap?.close?.();
    } catch {
      // Ignore cleanup failures so the real render error is preserved.
    }
    throw error;
  }
}

function drawPhotoEditorBaseToCanvas(
  targetCanvas,
  { maxEdge = null, fixedOutputSize = null } = {}
) {
  if (!photoEditorState?.sourceImage || !targetCanvas) {
    return null;
  }

  const crop = normalizePhotoEditorCropState(photoEditorState.crop);
  photoEditorState.crop = crop;
  const sourceRect = getPhotoEditorSourceRect(photoEditorState.sourceImage);
  const outputSize = getPhotoEditorOutputSize(
    sourceRect,
    maxEdge,
    crop,
    fixedOutputSize
  );
  const ctx = targetCanvas.getContext('2d', { willReadFrequently: true });

  if (!ctx) {
    return null;
  }

  targetCanvas.width = outputSize.width;
  targetCanvas.height = outputSize.height;
  ctx.clearRect(0, 0, outputSize.width, outputSize.height);
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  drawPhotoEditorCroppedSourceToCanvas(
    ctx,
    photoEditorState.sourceImage,
    sourceRect,
    outputSize,
    crop
  );

  return {
    ctx,
    outputSize,
    sourceRect,
    crop,
  };
}

function applyPhotoEditorEffectsToCanvas(
  ctx,
  outputSize,
  sourceRect,
  {
    includeDraft = false,
    drawOverlays = true,
    isInteractivePreview = false,
    textScale = 1,
    drawTextOverlay = true,
    drawImageOverlay = true,
    drawSubjectMaskPreview = true,
  } = {}
) {
  applyPhotoEditorTargetedAdjustmentsToCanvas(ctx, outputSize, sourceRect, {
    isInteractivePreview,
  });
  applyPhotoEditorBlurEffectToCanvas(
    ctx,
    outputSize.width,
    outputSize.height,
    photoEditorState.blur,
    sourceRect,
    { isInteractivePreview }
  );
  applyPhotoEditorMasksToCanvas(
    ctx,
    photoEditorState.masks,
    outputSize.width,
    outputSize.height,
    sourceRect,
    photoEditorState.sourceImage,
    { isInteractivePreview }
  );

  if (includeDraft) {
    applyPhotoEditorMasksToCanvas(
      ctx,
      photoEditorState.draftMask ? [photoEditorState.draftMask] : [],
      outputSize.width,
      outputSize.height,
      sourceRect,
      photoEditorState.sourceImage,
      { isInteractivePreview }
    );
  }

  if (drawImageOverlay || drawTextOverlay) {
    if (hasPhotoEditorReadySubjectMask()) {
      applyPhotoEditorMaskedOverlaysToCanvas(ctx, outputSize, sourceRect, {
        textScale,
      });
    } else {
      if (drawImageOverlay) {
        applyPhotoEditorImageOverlayToCanvas(
          ctx,
          outputSize.width,
          outputSize.height,
          photoEditorState.imageOverlays
        );
      }

      if (drawTextOverlay) {
        applyPhotoEditorTextOverlayToCanvas(
          ctx,
          outputSize.width,
          outputSize.height,
          getPhotoEditorTextCollectionFromState().textOverlays,
          { textScale }
        );
      }
    }
  }

  if (drawOverlays && drawSubjectMaskPreview) {
    drawPhotoEditorSubjectMaskPreview(ctx, outputSize, sourceRect);
  }

  if (includeDraft && drawOverlays) {
    drawPhotoEditorDraftMask(
      ctx,
      outputSize.width,
      outputSize.height,
      sourceRect,
      photoEditorState.sourceImage
    );
  }

  if (includeDraft && drawOverlays) {
    drawPhotoEditorRadialBlurControl(
      ctx,
      outputSize.width,
      outputSize.height
    );
  }
}

function renderPhotoEditorToCanvas(
  targetCanvas,
  {
    maxEdge = null,
    includeDraft = false,
    showOriginal = false,
    textScale = 1,
  } = {}
) {
  const renderBase = drawPhotoEditorBaseToCanvas(targetCanvas, { maxEdge });

  if (!renderBase) {
    return null;
  }

  const { ctx, outputSize, sourceRect } = renderBase;

  if (!showOriginal) {
    applyPhotoEditorEffectsToCanvas(ctx, outputSize, sourceRect, {
      includeDraft,
      textScale,
    });
  }

  return outputSize;
}

function analyzePhotoEditorPreviewClipping(canvas) {
  if (!canvas || canvas.width <= 0 || canvas.height <= 0) {
    return null;
  }

  const sampleEdge = 180;
  const scale = Math.min(1, sampleEdge / Math.max(canvas.width, canvas.height));
  const sampleWidth = Math.max(1, Math.round(canvas.width * scale));
  const sampleHeight = Math.max(1, Math.round(canvas.height * scale));
  const sampleCanvas = document.createElement('canvas');
  sampleCanvas.width = sampleWidth;
  sampleCanvas.height = sampleHeight;
  const sampleCtx = sampleCanvas.getContext('2d', { willReadFrequently: true });

  if (!sampleCtx) {
    return null;
  }

  sampleCtx.drawImage(canvas, 0, 0, sampleWidth, sampleHeight);

  let imageData;

  try {
    imageData = sampleCtx.getImageData(0, 0, sampleWidth, sampleHeight);
  } catch {
    return null;
  }

  const data = imageData.data;
  let shadowPixels = 0;
  let highlightPixels = 0;
  const totalPixels = Math.max(1, sampleWidth * sampleHeight);

  for (let index = 0; index < data.length; index += 4) {
    const luminance = 0.2126 * data[index] + 0.7152 * data[index + 1] + 0.0722 * data[index + 2];

    if (luminance <= 8) {
      shadowPixels += 1;
    } else if (luminance >= 248) {
      highlightPixels += 1;
    }
  }

  return {
    shadows: shadowPixels / totalPixels,
    highlights: highlightPixels / totalPixels,
  };
}

function formatPhotoEditorPercent(value) {
  return `${(clampNumber(value, 0, 1, 0) * 100).toFixed(value >= 0.01 ? 1 : 2)}%`;
}

function getPhotoEditorAdjustmentDelta(values, key) {
  const slider = getPhotoEditorSliderMeta(key);

  if (!slider) {
    return 0;
  }

  const baseValues = photoEditorPreviewCommittedValues || PHOTO_EDIT_DEFAULT_VALUES;
  const currentValue = clampNumber(
    values?.[key],
    slider.min,
    slider.max,
    slider.defaultValue
  );
  const committedValue = clampNumber(
    baseValues?.[key],
    slider.min,
    slider.max,
    slider.defaultValue
  );

  return currentValue - committedValue;
}

function getPhotoEditorAdjustmentLivePreviewFilter(values) {
  if (!values) {
    return '';
  }

  const brightnessDelta =
    getPhotoEditorAdjustmentDelta(values, 'brightness') * 0.44 +
    getPhotoEditorAdjustmentDelta(values, 'exposure') * 0.58 +
    getPhotoEditorAdjustmentDelta(values, 'gamma') * 0.28 +
    getPhotoEditorAdjustmentDelta(values, 'highlights') * 0.12 +
    getPhotoEditorAdjustmentDelta(values, 'shadows') * 0.12 +
    getPhotoEditorAdjustmentDelta(values, 'whites') * 0.28 -
    getPhotoEditorAdjustmentDelta(values, 'blacks') * 0.10;
  const contrastDelta =
    getPhotoEditorAdjustmentDelta(values, 'contrast') * 0.64 +
    getPhotoEditorAdjustmentDelta(values, 'clarity') * 0.20 +
    getPhotoEditorAdjustmentDelta(values, 'texture') * 0.12 +
    getPhotoEditorAdjustmentDelta(values, 'sharpness') * 0.06 -
    getPhotoEditorAdjustmentDelta(values, 'denoise') * 0.04 -
    getPhotoEditorAdjustmentDelta(values, 'fade') * 0.36;
  const saturationDelta =
    getPhotoEditorAdjustmentDelta(values, 'saturation') * 0.64 +
    getPhotoEditorAdjustmentDelta(values, 'vibrance') * 0.44;

  if (
    Math.abs(brightnessDelta) < 0.1 &&
    Math.abs(contrastDelta) < 0.1 &&
    Math.abs(saturationDelta) < 0.1
  ) {
    return '';
  }

  const brightness = clampNumber(1 + brightnessDelta / 185, 0.72, 1.38, 1);
  const contrast = clampNumber(1 + contrastDelta / 165, 0.72, 1.42, 1);
  const saturation = clampNumber(1 + saturationDelta / 155, 0.45, 1.70, 1);

  return [
    `brightness(${brightness.toFixed(3)})`,
    `contrast(${contrast.toFixed(3)})`,
    `saturate(${saturation.toFixed(3)})`,
  ].join(' ');
}

function clearPhotoEditorAdjustmentLivePreview() {
  if (!photoEditorCanvas) {
    return;
  }

  photoEditorCanvas.style.filter = '';
  photoEditorCanvas.style.willChange = '';
}

function applyPhotoEditorAdjustmentLivePreview() {
  if (
    !photoEditorCanvas ||
    !photoEditorState ||
    photoEditorState.showOriginalPreview
  ) {
    clearPhotoEditorAdjustmentLivePreview();
    return;
  }

  const filter = getPhotoEditorAdjustmentLivePreviewFilter(photoEditorState.values);
  photoEditorCanvas.style.filter = filter;
  photoEditorCanvas.style.willChange = filter ? 'filter' : '';
}

async function renderPhotoEditorPreview() {
  if (!photoEditorState?.sourceImage || !photoEditorCanvas) {
    return;
  }

  const renderToken = ++photoEditorPreviewRenderToken;
  const workCanvas = getPhotoEditorPreviewWorkCanvas();
  const renderBase = drawPhotoEditorBaseToCanvas(workCanvas, {
    maxEdge: getPhotoEditorPreviewMaxEdge(),
  });

  if (!renderBase) {
    return;
  }

  const { ctx: workCtx, outputSize, sourceRect } = renderBase;
  const shouldRenderEffects = !photoEditorState.showOriginalPreview;
  const isInteractivePreview =
    Boolean(photoEditorState.isInteractivePreview) ||
    Boolean(photoEditorState.dragMode);
  const dragMode = String(photoEditorState.dragMode || '');
  const includeDraftEffect =
    !['mask', 'mask-move', 'mask-resize', 'mask-rotate'].includes(dragMode);

  syncPhotoEditorPreviewCanvasDisplaySize(outputSize);

  if (shouldRenderEffects) {
    let didRenderWithWorker = false;

    try {
      if (hasPhotoEditorReadySubjectMask()) {
        throw new Error('Subject mask composition requires main renderer');
      }

      const workerResult = await applyPhotoEditorEffectsWithWorker(
        workCanvas,
        renderBase,
        {
          includeDraft: includeDraftEffect,
          cancelPendingPreview: true,
          isInteractivePreview,
        }
      );

      if (renderToken !== photoEditorPreviewRenderToken || !photoEditorState) {
        workerResult?.bitmap?.close?.();
        return;
      }

      if (workerResult?.bitmap) {
        workCtx.clearRect(0, 0, outputSize.width, outputSize.height);
        workCtx.drawImage(workerResult.bitmap, 0, 0);
        workerResult.bitmap.close?.();
        didRenderWithWorker = true;
      }
    } catch (error) {
      if (renderToken !== photoEditorPreviewRenderToken || !photoEditorState) {
        return;
      }

      didRenderWithWorker = false;
    }

    if (!didRenderWithWorker) {
      applyPhotoEditorEffectsToCanvas(workCtx, outputSize, sourceRect, {
        includeDraft: includeDraftEffect,
        drawOverlays: false,
        isInteractivePreview,
        drawTextOverlay: false,
        drawImageOverlay: false,
      });
    }
  }

  if (renderToken !== photoEditorPreviewRenderToken || !photoEditorState) {
    return;
  }

  if (shouldRenderEffects) {
    storePhotoEditorPreviewOverlayBase(workCanvas, outputSize, sourceRect);
    drawPhotoEditorPreviewTextAndOverlays(workCtx, outputSize, sourceRect);
  } else {
    photoEditorPreviewOverlayMeta = null;
  }

  clearPhotoEditorAdjustmentLivePreview();
  commitPhotoEditorPreviewFrame(workCanvas, outputSize);
  photoEditorPreviewCommittedValues = shouldRenderEffects
    ? clonePhotoEditValues(photoEditorState.values)
    : null;

  if (!photoEditorState.isInteractivePreview) {
    photoEditorState.lastClipInfo = analyzePhotoEditorPreviewClipping(photoEditorCanvas);
  }

  const pendingMaskText = photoEditorState.draftMask ? ' / 未確定: 1件' : '';
  const imageOverlayCount =
    getPhotoEditorImageOverlayCollectionFromState().imageOverlays.length;
  const imageOverlayText =
    imageOverlayCount > 0 ? ` / 画像: ${imageOverlayCount}件` : '';
  const clipInfo = photoEditorState.lastClipInfo;
  const clipText = clipInfo
    ? ` / 黒つぶれ: ${formatPhotoEditorPercent(clipInfo.shadows)} / 白飛び: ${formatPhotoEditorPercent(clipInfo.highlights)}`
    : '';
  const compareText = photoEditorState.showOriginalPreview ? ' / 比較中' : '';
  const fixedOutputSize = getPhotoEditorFixedOutputSize(photoEditorState.crop);
  const statusOutputWidth = fixedOutputSize?.width || outputSize.fullWidth;
  const statusOutputHeight = fixedOutputSize?.height || outputSize.fullHeight;
  setPhotoEditorStatus(
    `出力: ${statusOutputWidth} x ${statusOutputHeight} / ` +
      `目隠し: ${photoEditorState.masks.length}件${pendingMaskText}${imageOverlayText}${clipText}${compareText}`
  );
  syncPhotoEditorBlurControls();
  syncPhotoEditorMaskToolUi();
}

function markPhotoEditorInteractivePreview(settleMs = PHOTO_EDITOR_INTERACTIVE_PREVIEW_SETTLE_MS) {
  if (!photoEditorState) {
    return;
  }

  photoEditorState.isInteractivePreview = true;

  if (photoEditorPreviewSettleTimer) {
    clearTimeout(photoEditorPreviewSettleTimer);
    photoEditorPreviewSettleTimer = 0;
  }

  photoEditorPreviewSettleTimer = setTimeout(() => {
    photoEditorPreviewSettleTimer = 0;

    if (!photoEditorState || photoEditorState.dragMode) {
      return;
    }

    photoEditorState.isInteractivePreview = false;
    photoEditorState.isCropInteractivePreview = false;
    schedulePhotoEditorRender();
  }, settleMs);
}

function finishPhotoEditorInteractivePreview() {
  if (photoEditorPreviewSettleTimer) {
    clearTimeout(photoEditorPreviewSettleTimer);
    photoEditorPreviewSettleTimer = 0;
  }

  if (photoEditorRenderDebounceTimer) {
    clearTimeout(photoEditorRenderDebounceTimer);
    photoEditorRenderDebounceTimer = 0;
  }

  if (!photoEditorState) {
    return;
  }

  photoEditorState.isInteractivePreview = false;
  photoEditorState.isCropInteractivePreview = false;
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function getPhotoEditorInteractiveRenderDelay() {
  const dragMode = String(photoEditorState?.dragMode || '');

  if (
    dragMode === 'mask' ||
    dragMode.startsWith('blur-') ||
    dragMode === 'curve'
  ) {
    return PHOTO_EDITOR_HEAVY_INTERACTIVE_PREVIEW_DEBOUNCE_MS;
  }

  return 0;
}

function schedulePhotoEditorRender({ debounceMs = 0, interactive = false } = {}) {
  if (interactive) {
    markPhotoEditorInteractivePreview();
  }

  if (photoEditorRenderDebounceTimer) {
    clearTimeout(photoEditorRenderDebounceTimer);
    photoEditorRenderDebounceTimer = 0;
  }

  const renderDelay = debounceMs > 0
    ? debounceMs
    : interactive
      ? getPhotoEditorInteractiveRenderDelay()
      : 0;

  if (renderDelay > 0) {
    photoEditorRenderDebounceTimer = setTimeout(() => {
      photoEditorRenderDebounceTimer = 0;
      schedulePhotoEditorRender();
    }, renderDelay);
    return;
  }

  if (photoEditorRenderFrame) {
    cancelAnimationFrame(photoEditorRenderFrame);
  }

  photoEditorRenderFrame = requestAnimationFrame(() => {
    photoEditorRenderFrame = 0;
    renderPhotoEditorPreview();
  });
}

function resetPhotoEditorAdjustments() {
  if (!photoEditorState) {
    return;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.values = clonePhotoEditValues();
  photoEditorState.adjustmentTarget = 'whole';
  clearPhotoEditorAutoEnhanceState();
  syncPhotoEditorAdjustmentControls();
  syncPhotoEditorAutoEnhanceControls();
  clearPhotoEditorAdjustmentLivePreview();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function resetPhotoEditorCrop() {
  if (!photoEditorState) {
    return;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.crop = getDefaultPhotoEditorCropState();
  syncPhotoEditorCropControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function resetPhotoEditorExportSettings() {
  if (!photoEditorState) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.exportSettings = getDefaultPhotoEditorExportSettings();
  syncPhotoEditorExportControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function resetPhotoEditorBlur() {
  if (!photoEditorState) {
    return;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.blur = getDefaultPhotoEditorBlurState();
  photoEditorState.dragInitialBlur = null;
  syncPhotoEditorBlurControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function resetPhotoEditorCurve() {
  if (!photoEditorState) {
    return;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.curve = getDefaultPhotoEditorCurveState();
  photoEditorState.dragCurvePointIndex = null;
  syncPhotoEditorCurveControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function resetPhotoEditorAll() {
  if (!photoEditorState) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.values = clonePhotoEditValues();
  photoEditorState.adjustmentTarget = 'whole';
  clearPhotoEditorAutoEnhanceState();
  photoEditorState.crop = getDefaultPhotoEditorCropState();
  photoEditorState.exportSettings = getDefaultPhotoEditorExportSettings();
  setPhotoEditorTextCollection([], '');
  setPhotoEditorImageOverlayCollection([], '');
  photoEditorState.subjectMask = getDefaultPhotoEditorSubjectMaskState();
  photoEditorState.rulerGuides = normalizePhotoEditorRulerGuides();
  photoEditorState.draftRulerGuide = null;
  photoEditorState.dragInitialRulerGuide = null;
  photoEditorState.snapGuide = null;
  photoEditorState.blur = getDefaultPhotoEditorBlurState();
  photoEditorState.curve = getDefaultPhotoEditorCurveState();
  photoEditorState.masks = [];
  photoEditorState.maskTool = 'none';
  photoEditorState.maskShape = 'rect';
  photoEditorState.maskStrengths = { ...PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS };
  photoEditorState.blurStrength = PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS.blur;
  photoEditorState.dragMode = null;
  photoEditorState.dragStart = null;
  photoEditorState.dragInitialCrop = null;
  photoEditorState.dragInitialSourceRect = null;
  photoEditorState.dragInitialBlur = null;
  photoEditorState.dragInitialMask = null;
  photoEditorState.dragInitialText = null;
  photoEditorState.dragInitialImageOverlay = null;
  photoEditorState.draftRulerGuide = null;
  photoEditorState.dragInitialRulerGuide = null;
  photoEditorState.snapGuide = null;
  photoEditorState.dragCurvePointIndex = null;
  photoEditorState.draftMask = null;
  photoEditorState.showOriginalPreview = false;
  photoEditorState.isInteractivePreview = false;
  photoEditorState.isCropInteractivePreview = false;
  photoEditorState.lastClipInfo = null;
  photoEditorCanvas?.classList.remove(
    'is-panning',
    'is-mask-draft-active',
    'is-text-tool-active',
    'is-image-overlay-tool-active'
  );
  clearPhotoEditorAdjustmentLivePreview();
  syncPhotoEditorUi();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function updatePhotoEditorExportSettings(nextSettings = {}) {
  if (!photoEditorState) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.exportSettings = getPhotoEditorEffectiveExportSettings({
    ...photoEditorState.exportSettings,
    ...nextSettings,
  });
  syncPhotoEditorExportControls();
  schedulePhotoEditorRender();
  schedulePhotoEditorHistoryCommit();
}

function togglePhotoEditorComparePreview() {
  if (!photoEditorState) {
    return;
  }

  photoEditorState.showOriginalPreview = !photoEditorState.showOriginalPreview;
  clearPhotoEditorAdjustmentLivePreview();
  syncPhotoEditorCompareControl();
  schedulePhotoEditorRender();
}

function updatePhotoEditorBlur(nextBlur = {}) {
  if (!photoEditorState) {
    return;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  beginPhotoEditorHistoryMutation();
  const currentBlur = normalizePhotoEditorBlurState(photoEditorState.blur);
  const shouldReopenRadialControl =
    nextBlur.mode === 'radial' ||
    nextBlur.centerX !== undefined ||
    nextBlur.centerY !== undefined ||
    nextBlur.radius !== undefined ||
    nextBlur.outerRadius !== undefined;

  photoEditorState.blur = normalizePhotoEditorBlurState({
    ...currentBlur,
    ...nextBlur,
    isConfirmed: shouldReopenRadialControl
      ? false
      : nextBlur.isConfirmed ?? currentBlur.isConfirmed,
  });
  syncPhotoEditorBlurControls();
  schedulePhotoEditorRender({
    debounceMs: PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS,
    interactive: true,
  });
  schedulePhotoEditorHistoryCommit();
}

function setPhotoEditorCurveMode(mode) {
  if (!photoEditorState) {
    return;
  }

  const curve = normalizePhotoEditorCurveState(photoEditorState.curve);
  const nextMode = PHOTO_EDITOR_CURVE_MODES.includes(mode) ? mode : curve.mode;
  const validChannelKeys = PHOTO_EDITOR_CURVE_CHANNELS[nextMode].map(
    (channel) => channel.key
  );

  photoEditorState.curve = normalizePhotoEditorCurveState({
    ...curve,
    mode: nextMode,
    channel: validChannelKeys.includes(curve.channel)
      ? curve.channel
      : validChannelKeys[0],
  });
  syncPhotoEditorCurveControls();
}

function setPhotoEditorCurveChannel(channel) {
  if (!photoEditorState) {
    return;
  }

  const curve = normalizePhotoEditorCurveState(photoEditorState.curve);
  const validChannelKeys = PHOTO_EDITOR_CURVE_CHANNELS[curve.mode].map(
    (entry) => entry.key
  );

  if (!validChannelKeys.includes(channel)) {
    return;
  }

  photoEditorState.curve = {
    ...curve,
    channel,
  };
  syncPhotoEditorCurveControls();
}

function confirmPhotoEditorBlur() {
  if (!photoEditorState) {
    return;
  }

  const blur = normalizePhotoEditorBlurState(photoEditorState.blur);

  if (blur.mode !== 'radial') {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.blur = normalizePhotoEditorBlurState({
    ...blur,
    isConfirmed: !blur.isConfirmed,
  });
  syncPhotoEditorBlurControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function setPhotoEditorMaskTool(tool) {
  if (!photoEditorState) {
    return;
  }

  const normalizedTool = ['blur', 'mosaic', 'fill'].includes(tool)
    ? tool
    : 'none';

  const shouldActivateTool =
    normalizedTool !== 'none' && photoEditorState.maskTool !== normalizedTool;

  if (photoEditorState.draftMask) {
    cancelPhotoEditorPendingMask({ deactivateTool: false });
  }

  photoEditorState.maskTool = shouldActivateTool ? normalizedTool : 'none';

  if (shouldActivateTool) {
    photoEditorState.maskShape = 'rect';
    photoEditorState.blurStrength =
      photoEditorState.maskStrengths?.[normalizedTool] ??
      PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS[normalizedTool] ??
      PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS.blur;
  }

  syncPhotoEditorMaskToolUi();
}

function setPhotoEditorMaskShape(shape) {
  if (!photoEditorState) {
    return;
  }

  if (photoEditorState.draftMask) {
    cancelPhotoEditorPendingMask({ deactivateTool: false });
  }

  photoEditorState.maskShape = ['rect', 'ellipse', 'freehand'].includes(shape)
    ? shape
    : 'rect';
  syncPhotoEditorMaskToolUi();
}

function setPhotoEditorBlurStrength(value) {
  if (!photoEditorState) {
    return;
  }

  const activeMaskType = photoEditorState.draftMask?.type || photoEditorState.maskTool;
  const nextStrength = clampNumber(
    value,
    0,
    100,
    PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS[activeMaskType] ??
      PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS.blur
  );

  beginPhotoEditorHistoryMutation();
  photoEditorState.blurStrength = nextStrength;

  if (['blur', 'mosaic', 'fill'].includes(activeMaskType)) {
    photoEditorState.maskStrengths = {
      ...PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS,
      ...(photoEditorState.maskStrengths || {}),
      [activeMaskType]: nextStrength,
    };
  }

  if (['blur', 'mosaic', 'fill'].includes(photoEditorState.draftMask?.type)) {
    updatePhotoEditorDraftMask({ strength: photoEditorState.blurStrength });
  }
  syncPhotoEditorMaskToolUi();
  schedulePhotoEditorHistoryCommit();
}

function getPhotoEditorHistogramPercentile(histogram, percentile) {
  if (!Array.isArray(histogram) || histogram.length === 0) {
    return 0;
  }

  const total = histogram.reduce((sum, count) => sum + count, 0);

  if (total <= 0) {
    return 0;
  }

  const target = total * clampNumber(percentile, 0, 1, 0);
  let cumulative = 0;

  for (let index = 0; index < histogram.length; index += 1) {
    cumulative += histogram[index] || 0;

    if (cumulative >= target) {
      return index / Math.max(1, histogram.length - 1);
    }
  }

  return 1;
}

function analyzePhotoEditorSourceImage() {
  if (!photoEditorState?.sourceImage) {
    return null;
  }

  const crop = normalizePhotoEditorCropState(photoEditorState.crop);
  photoEditorState.crop = crop;
  const sourceRect = getPhotoEditorSourceRect(photoEditorState.sourceImage);
  const sampleSize = getPhotoEditorOutputSize(
    sourceRect,
    PHOTO_EDITOR_AUTO_ENHANCE_ANALYSIS_MAX_EDGE,
    crop
  );
  const sampleCanvas = document.createElement('canvas');
  sampleCanvas.width = sampleSize.width;
  sampleCanvas.height = sampleSize.height;
  const sampleCtx = sampleCanvas.getContext('2d', { willReadFrequently: true });

  if (!sampleCtx) {
    return null;
  }

  sampleCtx.imageSmoothingEnabled = true;
  sampleCtx.imageSmoothingQuality = 'high';
  drawPhotoEditorCroppedSourceToCanvas(
    sampleCtx,
    photoEditorState.sourceImage,
    sourceRect,
    sampleSize,
    crop
  );

  let imageData;

  try {
    imageData = sampleCtx.getImageData(0, 0, sampleSize.width, sampleSize.height);
  } catch {
    return null;
  }

  const data = imageData.data;
  const pixelCount = Math.max(1, data.length / 4);
  let luminanceTotal = 0;
  let luminanceSquareTotal = 0;
  let saturationTotal = 0;
  let saturationSquareTotal = 0;
  let redTotal = 0;
  let greenTotal = 0;
  let blueTotal = 0;
  let neonCount = 0;
  let darkCount = 0;
  let brightCount = 0;
  let nearBlackCount = 0;
  let shadowCount = 0;
  let midtoneCount = 0;
  let highlightCount = 0;
  let nearWhiteCount = 0;
  let channelClipCount = 0;
  let channelNearClipCount = 0;
  let chromaClipCount = 0;
  const luminanceHistogram = Array.from({ length: 256 }, () => 0);

  for (let index = 0; index < data.length; index += 4) {
    const red = data[index] / 255;
    const green = data[index + 1] / 255;
    const blue = data[index + 2] / 255;
    const luminance = 0.2126 * red + 0.7152 * green + 0.0722 * blue;
    const maxChannel = Math.max(red, green, blue);
    const minChannel = Math.min(red, green, blue);
    const saturation = maxChannel === 0 ? 0 : (maxChannel - minChannel) / maxChannel;

    luminanceTotal += luminance;
    luminanceSquareTotal += luminance * luminance;
    saturationTotal += saturation;
    saturationSquareTotal += saturation * saturation;
    redTotal += red;
    greenTotal += green;
    blueTotal += blue;
    luminanceHistogram[
      clampNumber(Math.round(luminance * 255), 0, 255, 0)
    ] += 1;

    if (luminance < 0.08) {
      darkCount += 1;
    }

    if (luminance < 0.03) {
      nearBlackCount += 1;
    }

    if (luminance < 0.25) {
      shadowCount += 1;
    }

    if (luminance >= 0.24 && luminance <= 0.72) {
      midtoneCount += 1;
    }

    if (luminance > 0.75) {
      highlightCount += 1;
    }

    if (luminance > 0.92) {
      brightCount += 1;
    }

    if (luminance > 0.97) {
      nearWhiteCount += 1;
    }

    if (maxChannel > 0.985) {
      channelClipCount += 1;
    }

    if (maxChannel > 0.955) {
      channelNearClipCount += 1;
    }

    if (saturation > 0.62 && maxChannel > 0.94) {
      chromaClipCount += 1;
    }

    if (saturation > 0.55 && luminance > 0.18) {
      neonCount += 1;
    }
  }

  const mean = luminanceTotal / pixelCount;
  const variance = Math.max(0, luminanceSquareTotal / pixelCount - mean * mean);
  const saturationMean = saturationTotal / pixelCount;
  const saturationVariance = Math.max(
    0,
    saturationSquareTotal / pixelCount - saturationMean * saturationMean
  );
  const percentiles = {
    p01: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.01),
    p02: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.02),
    p05: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.05),
    p10: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.10),
    p25: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.25),
    p50: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.50),
    p75: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.75),
    p90: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.90),
    p95: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.95),
    p98: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.98),
    p99: getPhotoEditorHistogramPercentile(luminanceHistogram, 0.99),
  };

  return {
    mean,
    meanLuminance: mean,
    deviation: Math.sqrt(variance),
    sampleWidth: sampleSize.width,
    sampleHeight: sampleSize.height,
    sampleData: new Uint8ClampedArray(data),
    histogram: luminanceHistogram,
    percentiles,
    medianLuminance: percentiles.p50,
    dynamicRange: Math.max(0, percentiles.p95 - percentiles.p05),
    contrastScore: Math.max(0, percentiles.p95 - percentiles.p05),
    darkRatio: darkCount / pixelCount,
    brightRatio: brightCount / pixelCount,
    nearBlackRatio: nearBlackCount / pixelCount,
    shadowRatio: shadowCount / pixelCount,
    midtoneRatio: midtoneCount / pixelCount,
    highlightRatio: highlightCount / pixelCount,
    nearWhiteRatio: nearWhiteCount / pixelCount,
    redMean: redTotal / pixelCount,
    greenMean: greenTotal / pixelCount,
    blueMean: blueTotal / pixelCount,
    channelClipRatio: channelClipCount / pixelCount,
    channelNearClipRatio: channelNearClipCount / pixelCount,
    chromaClipRatio: chromaClipCount / pixelCount,
    neonRatio: neonCount / pixelCount,
    saturation: saturationMean,
    saturationMean,
    saturationStd: Math.sqrt(saturationVariance),
    clippingBlackRatio: nearBlackCount / pixelCount,
    clippingWhiteRatio: nearWhiteCount / pixelCount,
    colorCast: {
      warmthBias: redTotal / pixelCount - blueTotal / pixelCount,
      tintBias:
        (redTotal / pixelCount + blueTotal / pixelCount) / 2 -
        greenTotal / pixelCount,
    },
  };
}

function estimatePhotoEditorSmartAutoResult(analysis, values) {
  if (
    !analysis?.sampleData ||
    !Number.isFinite(analysis.sampleWidth) ||
    !Number.isFinite(analysis.sampleHeight)
  ) {
    return null;
  }

  const width = Math.max(1, Math.round(analysis.sampleWidth));
  const height = Math.max(1, Math.round(analysis.sampleHeight));
  const imageData = new ImageData(new Uint8ClampedArray(analysis.sampleData), width, height);

  applyPhotoEditorPixelAdjustmentsToImageData(
    imageData,
    width,
    getPhotoEditorAdjustmentRenderParams(clonePhotoEditValues(values))
  );

  const data = imageData.data;
  const pixelCount = Math.max(1, data.length / 4);
  let luminanceTotal = 0;
  let nearWhiteCount = 0;
  let nearBlackCount = 0;
  let channelClipCount = 0;
  let channelNearClipCount = 0;

  for (let index = 0; index < data.length; index += 4) {
    const red = data[index] / 255;
    const green = data[index + 1] / 255;
    const blue = data[index + 2] / 255;
    const luminance = 0.2126 * red + 0.7152 * green + 0.0722 * blue;
    const maxChannel = Math.max(red, green, blue);

    luminanceTotal += luminance;

    if (luminance > 0.97) {
      nearWhiteCount += 1;
    }

    if (luminance < 0.03) {
      nearBlackCount += 1;
    }

    if (maxChannel > 0.985) {
      channelClipCount += 1;
    }

    if (maxChannel > 0.955) {
      channelNearClipCount += 1;
    }
  }

  return {
    mean: luminanceTotal / pixelCount,
    nearWhiteRatio: nearWhiteCount / pixelCount,
    nearBlackRatio: nearBlackCount / pixelCount,
    channelClipRatio: channelClipCount / pixelCount,
    channelNearClipRatio: channelNearClipCount / pixelCount,
  };
}

function protectPhotoEditorSmartAutoHighlights(analysis, values) {
  let nextValues = clonePhotoEditValues(values);
  const sourceNearWhite = Number(analysis?.nearWhiteRatio) || 0;
  const sourceChannelClip = Number(analysis?.channelClipRatio) || 0;
  const sourceChannelNearClip = Number(analysis?.channelNearClipRatio) || 0;
  const allowedNearWhite = Math.max(0.004, sourceNearWhite + 0.002);
  const allowedChannelClip = Math.max(0.012, sourceChannelClip + 0.004);
  const allowedChannelNearClip = Math.max(0.040, sourceChannelNearClip + 0.010);

  for (let attempt = 0; attempt < 4; attempt += 1) {
    const result = estimatePhotoEditorSmartAutoResult(analysis, nextValues);

    if (!result) {
      return nextValues;
    }

    const nearWhiteExcess = Math.max(0, result.nearWhiteRatio - allowedNearWhite);
    const channelClipExcess = Math.max(0, result.channelClipRatio - allowedChannelClip);
    const nearClipExcess = Math.max(0, result.channelNearClipRatio - allowedChannelNearClip);
    const outputIsTooHot =
      nearWhiteExcess > 0 ||
      channelClipExcess > 0 ||
      nearClipExcess > 0 ||
      result.mean > 0.66;

    if (!outputIsTooHot) {
      return nextValues;
    }

    const protection = clampNumber(
      nearWhiteExcess * 18 +
        channelClipExcess * 12 +
        nearClipExcess * 5 +
        Math.max(0, result.mean - 0.62) * 3,
      0.35,
      1,
      0.5
    );

    nextValues = clonePhotoEditValues({
      ...nextValues,
      exposure: Math.round(nextValues.exposure - 5 * protection),
      brightness: Math.round(nextValues.brightness - 3 * protection),
      highlights: Math.round(nextValues.highlights - 14 * protection),
      whites: Math.round(nextValues.whites - 18 * protection),
      gamma: Math.round(nextValues.gamma - 8 * protection),
      saturation: Math.round(nextValues.saturation - 3 * protection),
      vibrance: Math.round(nextValues.vibrance - 3 * protection),
    });
  }

  return nextValues;
}

function normalizePhotoEditorAutoStrength(value) {
  const numericValue = Number(value);

  if (!Number.isFinite(numericValue)) {
    return 1;
  }

  return numericValue > 1
    ? clampNumber(numericValue / 100, 0, 1, 1)
    : clampNumber(numericValue, 0, 1, 1);
}

function detectPhotoEditorAutoEnhanceMode(analysis) {
  const median = clampNumber(
    analysis?.medianLuminance ?? analysis?.percentiles?.p50,
    0,
    1,
    analysis?.mean ?? 0.5
  );
  const shadowRatio = Number(analysis?.shadowRatio) || 0;
  const neonRatio = Number(analysis?.neonRatio) || 0;
  const chromaClipRatio = Number(analysis?.chromaClipRatio) || 0;
  const channelNearClipRatio = Number(analysis?.channelNearClipRatio) || 0;
  const saturationMean = Number(analysis?.saturationMean ?? analysis?.saturation) || 0;

  if (
    neonRatio > 0.18 ||
    chromaClipRatio > 0.018 ||
    (channelNearClipRatio > 0.075 && saturationMean > 0.34)
  ) {
    return 'neon';
  }

  if (median < 0.38 || shadowRatio > 0.5) {
    return 'dark';
  }

  return 'standard';
}

function generatePhotoEditorAutoEnhanceParams(analysis) {
  const percentiles = analysis?.percentiles || {};
  const medianLuminance = clampNumber(
    analysis?.medianLuminance ?? percentiles.p50,
    0,
    1,
    analysis?.mean ?? 0.5
  );
  const p05 = clampNumber(percentiles.p05, 0, 1, analysis?.mean ?? 0.5);
  const p95 = clampNumber(percentiles.p95, 0, 1, analysis?.mean ?? 0.5);
  const p98 = clampNumber(percentiles.p98, 0, 1, analysis?.mean ?? 0.5);
  const dynamicRange = clampNumber(
    analysis?.dynamicRange ?? p95 - p05,
    0,
    1,
    Math.max(0, p95 - p05)
  );
  const shadowRatio = Number(analysis?.shadowRatio) || 0;
  const highlightRatio = Number(analysis?.highlightRatio) || 0;
  const clippingBlackRatio =
    Number(analysis?.clippingBlackRatio ?? analysis?.nearBlackRatio) || 0;
  const clippingWhiteRatio =
    Number(analysis?.clippingWhiteRatio ?? analysis?.nearWhiteRatio) || 0;
  const channelNearClipRatio = Number(analysis?.channelNearClipRatio) || 0;
  const saturationMean =
    Number(analysis?.saturationMean ?? analysis?.saturation) || 0;
  const colorCast = analysis?.colorCast || {};
  const mode = detectPhotoEditorAutoEnhanceMode(analysis);
  const hasHotHighlights =
    clippingWhiteRatio > 0.01 ||
    channelNearClipRatio > 0.05 ||
    p98 > 0.88;

  let exposure = 0;
  if (medianLuminance < 0.38) {
    exposure += 0.12;
  } else if (medianLuminance < 0.45) {
    exposure += 0.06;
  }
  if (clippingWhiteRatio > 0.01) {
    exposure -= 0.04;
  }
  if (mode === 'neon') {
    exposure -= 0.04;
  }
  if (mode === 'dark' && hasHotHighlights) {
    exposure = Math.min(exposure, 0.02);
  }
  exposure = clampNumber(exposure, -0.10, 0.18, 0);

  let brilliance = 0.18;
  if (shadowRatio > 0.35) {
    brilliance += 0.18;
  }
  if (dynamicRange < 0.45) {
    brilliance += 0.12;
  }
  if (clippingWhiteRatio > 0.03 || channelNearClipRatio > 0.08) {
    brilliance -= 0.08;
  }
  if (mode === 'dark') {
    brilliance += hasHotHighlights ? 0.02 : 0.05;
  } else if (mode === 'neon') {
    brilliance -= 0.03;
  }
  brilliance = clampNumber(brilliance, 0.05, mode === 'dark' ? 0.40 : 0.45, 0.18);

  let highlights = -0.10;
  if (highlightRatio > 0.25) {
    highlights -= 0.10;
  }
  if (clippingWhiteRatio > 0.01 || channelNearClipRatio > 0.05) {
    highlights -= 0.15;
  }
  if (clippingWhiteRatio > 0.04 || p98 > 0.92) {
    highlights -= 0.10;
  }
  if (mode === 'neon') {
    highlights -= 0.08;
  }
  highlights = clampNumber(highlights, -0.50, 0.05, -0.10);

  let shadows = 0.08;
  if (shadowRatio > 0.30) {
    shadows += 0.15;
  }
  if (shadowRatio > 0.50) {
    shadows += 0.12;
  }
  if (clippingBlackRatio > 0.08) {
    shadows -= 0.05;
  }
  if (mode === 'dark') {
    shadows += hasHotHighlights ? 0.02 : 0.05;
  }
  shadows = clampNumber(
    shadows,
    0,
    mode === 'dark' && hasHotHighlights ? 0.34 : 0.42,
    0.08
  );

  let contrast = 0.04;
  if (dynamicRange < 0.40) {
    contrast += 0.10;
  }
  if (dynamicRange > 0.70 || hasHotHighlights) {
    contrast -= 0.04;
  }
  contrast = clampNumber(contrast, -0.05, 0.18, 0.04);

  let brightness = 0;
  if (medianLuminance < 0.40 && !hasHotHighlights) {
    brightness += 0.05;
  } else if (medianLuminance < 0.40) {
    brightness += 0.02;
  }
  if (medianLuminance > 0.62 || clippingWhiteRatio > 0.03) {
    brightness -= 0.03;
  }
  brightness = clampNumber(brightness, -0.08, 0.10, 0);

  let blackPoint = 0.04;
  if (clippingBlackRatio < 0.01 && dynamicRange < 0.55) {
    blackPoint += 0.08;
  }
  if (shadowRatio > 0.55) {
    blackPoint -= 0.03;
  }
  blackPoint = clampNumber(blackPoint, 0, 0.16, 0.04);

  let saturation = 0.02;
  if (saturationMean < 0.25) {
    saturation += 0.04;
  }
  if (saturationMean > 0.55 || mode === 'neon') {
    saturation -= 0.03;
  }
  saturation = clampNumber(saturation, -0.05, 0.08, 0.02);

  let vibrance = 0.08;
  if (saturationMean < 0.35) {
    vibrance += 0.08;
  }
  if (dynamicRange < 0.45) {
    vibrance += 0.04;
  }
  if (mode === 'neon') {
    vibrance = Math.min(vibrance + 0.02, 0.16);
  }
  vibrance = clampNumber(vibrance, 0, 0.22, 0.08);

  let warmth = 0;
  let tint = 0;
  if (Math.abs(Number(colorCast.warmthBias) || 0) > 0.18) {
    warmth = clampNumber(-colorCast.warmthBias * 0.15, -0.05, 0.05, 0);
  }
  if (Math.abs(Number(colorCast.tintBias) || 0) > 0.18) {
    tint = clampNumber(-colorCast.tintBias * 0.12, -0.04, 0.04, 0);
  }

  let definition = 0.06;
  if (dynamicRange < 0.45) {
    definition += 0.04;
  }
  definition = clampNumber(definition, 0, 0.12, 0.06);

  const sharpness = clampNumber(0.03, 0, 0.08, 0.03);
  const noiseReduction =
    shadowRatio > 0.55 && medianLuminance < 0.35
      ? clampNumber(mode === 'dark' ? 0.06 : 0.05, 0, 0.08, 0.05)
      : 0;

  return {
    mode,
    exposure,
    brilliance,
    highlights,
    shadows,
    contrast,
    brightness,
    blackPoint,
    saturation,
    vibrance,
    warmth,
    tint,
    definition,
    sharpness,
    noiseReduction,
  };
}

function convertPhotoEditorAutoEnhanceParamsToValues(params, analysis) {
  const percentiles = analysis?.percentiles || {};
  const p98 = clampNumber(percentiles.p98, 0, 1, analysis?.mean ?? 0.5);
  const clippingWhiteRatio =
    Number(analysis?.clippingWhiteRatio ?? analysis?.nearWhiteRatio) || 0;
  const channelNearClipRatio = Number(analysis?.channelNearClipRatio) || 0;
  const hotHighlightProtection = clampNumber(
    clippingWhiteRatio * 3.2 +
      channelNearClipRatio * 0.42 +
      Math.max(0, p98 - 0.88) * 1.9 +
      (params.mode === 'neon' ? 0.06 : 0),
    0.04,
    0.50,
    0.08
  );

  return clonePhotoEditValues({
    exposure: Math.round(params.exposure * 85),
    brightness: Math.round(params.brightness * 255),
    highlights: Math.round(params.highlights * 100 - params.brilliance * 10),
    shadows: Math.round(params.shadows * 100 + params.brilliance * 16),
    whites: -Math.round(
      clampNumber(
        hotHighlightProtection + Math.max(0, -params.highlights) * 0.24,
        0,
        0.58,
        0.08
      ) * 100
    ),
    blacks: -Math.round(params.blackPoint * 100),
    gamma: Math.round(
      params.brilliance * 8 +
        params.shadows * 6 -
        hotHighlightProtection * 28
    ),
    contrast: Math.round(params.contrast * 100 + params.brilliance * 8),
    temperature: Math.round(params.warmth * 100),
    tint: Math.round(params.tint * 100),
    saturation: Math.round(params.saturation * 100),
    vibrance: Math.round(params.vibrance * 100),
    clarity: Math.round(params.definition * 100 + params.brilliance * 5),
    texture: Math.round(params.definition * 65),
    sharpness: Math.round(params.sharpness * 100),
    denoise: Math.round(params.noiseReduction * 100),
    fade: 0,
    grain: 0,
    vignette: params.mode === 'neon' ? 4 : 0,
  });
}

function calculateSmartAutoPhotoEditValues({ strength = 1, baseValues = null } = {}) {
  const analysis = analyzePhotoEditorSourceImage();
  const normalizedStrength = normalizePhotoEditorAutoStrength(strength);

  if (!analysis) {
    const fallbackValues = clonePhotoEditValues({
      brightness: 2,
      exposure: 0,
      contrast: 6,
      highlights: -14,
      shadows: 8,
      whites: -8,
      blacks: -4,
      gamma: 3,
      saturation: 2,
      vibrance: 6,
      clarity: 4,
      texture: 3,
      sharpness: 6,
      denoise: 4,
      vignette: 0,
    });

    return blendPhotoEditValues(
      baseValues ? clonePhotoEditValues(baseValues) : clonePhotoEditValues(),
      fallbackValues,
      normalizedStrength
    );
  }

  const autoParams = generatePhotoEditorAutoEnhanceParams(analysis);
  const generatedValues = protectPhotoEditorSmartAutoHighlights(
    analysis,
    convertPhotoEditorAutoEnhanceParamsToValues(autoParams, analysis)
  );

  return blendPhotoEditValues(
    baseValues ? clonePhotoEditValues(baseValues) : clonePhotoEditValues(),
    generatedValues,
    normalizedStrength
  );
}

function calculatePhotoEditorUserPresetAverageValues() {
  const usablePresets = photoEditorUserPresets
    .map((preset) => normalizePhotoEditorUserPreset(preset))
    .filter(Boolean);

  if (usablePresets.length === 0) {
    return null;
  }

  const totals = { ...PHOTO_EDIT_DEFAULT_VALUES };

  for (const preset of usablePresets) {
    const values = clonePhotoEditValues(preset.values);

    for (const slider of PHOTO_EDIT_SLIDERS) {
      totals[slider.key] += values[slider.key];
    }
  }

  return clonePhotoEditValues(
    Object.fromEntries(
      PHOTO_EDIT_SLIDERS.map((slider) => [
        slider.key,
        Math.round(totals[slider.key] / usablePresets.length),
      ])
    )
  );
}

function blendPhotoEditValues(baseValues, overlayValues, weight) {
  const normalizedWeight = clampNumber(weight, 0, 1, 0);

  return clonePhotoEditValues(
    Object.fromEntries(
      PHOTO_EDIT_SLIDERS.map((slider) => [
        slider.key,
        Math.round(
          baseValues[slider.key] * (1 - normalizedWeight) +
            overlayValues[slider.key] * normalizedWeight
        ),
      ])
    )
  );
}

function calculateLearningPhotoEditValues() {
  const smartValues = calculateSmartAutoPhotoEditValues();
  const learnedValues = calculatePhotoEditorUserPresetAverageValues();

  if (!learnedValues) {
    showToast('保存済みプリセットがないためスマート自動補正を適用しました');
    return smartValues;
  }

  const learningWeight = Math.min(
    0.45,
    0.22 + photoEditorUserPresets.length * 0.035
  );

  return blendPhotoEditValues(smartValues, learnedValues, learningWeight);
}

function clearPhotoEditorAutoEnhanceState({ keepStrength = true } = {}) {
  if (!photoEditorState) {
    return;
  }

  const currentStrength = photoEditorState.autoEnhance?.strength;
  photoEditorState.autoEnhance = {
    ...getDefaultPhotoEditorAutoEnhanceState(),
    strength: keepStrength
      ? clampNumber(
          currentStrength,
          0,
          100,
          PHOTO_EDITOR_AUTO_ENHANCE_DEFAULT_STRENGTH
        )
      : PHOTO_EDITOR_AUTO_ENHANCE_DEFAULT_STRENGTH,
  };
}

function getPhotoEditorAutoEnhanceStrengthValue() {
  return clampNumber(
    photoEditorState?.autoEnhance?.strength,
    0,
    100,
    PHOTO_EDITOR_AUTO_ENHANCE_DEFAULT_STRENGTH
  );
}

function getPhotoEditorAutoEnhanceGeneratedValues(preset) {
  if (preset?.isLearningAuto) {
    return calculateLearningPhotoEditValues();
  }

  if (preset?.isAuto) {
    return calculateSmartAutoPhotoEditValues();
  }

  return clonePhotoEditValues(preset?.values);
}

function applyPhotoEditorAutoEnhancePreset(presetKey, preset) {
  if (!photoEditorState || !preset) {
    return;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  const normalizedState = normalizePhotoEditorAutoEnhanceState(
    photoEditorState.autoEnhance
  );
  const isTogglingOff =
    (preset.isLearningAuto || preset.isAuto) &&
    normalizedState.enabled &&
    normalizedState.presetKey === presetKey;

  beginPhotoEditorHistoryMutation();

  if (isTogglingOff) {
    photoEditorState.values = clonePhotoEditValues(
      normalizedState.originalValuesBeforeAuto
    );
    clearPhotoEditorAutoEnhanceState();
    syncPhotoEditorUi();
    clearPhotoEditorAdjustmentLivePreview();
    schedulePhotoEditorRender();
    commitPhotoEditorHistoryMutation();
    setPhotoEditorStatus('自動補正を解除しました');
    return;
  }

  const originalValues = clonePhotoEditValues(photoEditorState.values);
  const generatedValues = getPhotoEditorAutoEnhanceGeneratedValues(preset);
  const strength = getPhotoEditorAutoEnhanceStrengthValue();

  photoEditorState.autoEnhance = {
    enabled: true,
    originalValuesBeforeAuto: originalValues,
    generatedValues,
    strength,
    presetKey,
  };
  photoEditorState.values = blendPhotoEditValues(
    originalValues,
    generatedValues,
    normalizePhotoEditorAutoStrength(strength)
  );

  if (preset.crop) {
    photoEditorState.crop = normalizePhotoEditorCropState(preset.crop);
  }

  if (preset.blur) {
    photoEditorState.blur = normalizePhotoEditorBlurState(preset.blur);
  }

  if (preset.curve) {
    photoEditorState.curve = normalizePhotoEditorCurveState(preset.curve);
  }

  if (preset.exportSettings) {
    photoEditorState.exportSettings = normalizePhotoEditorExportSettings(
      preset.exportSettings
    );
  }

  syncPhotoEditorUi();
  clearPhotoEditorAdjustmentLivePreview();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
  setPhotoEditorStatus(
    preset.isLearningAuto
      ? '学習補正を適用しました'
      : preset.isAuto
        ? '自動補正を適用しました'
        : `プリセットを適用しました: ${preset.label || 'プリセット'}`
  );
}

function updatePhotoEditorAutoEnhanceStrength(value) {
  if (!photoEditorState) {
    return;
  }

  const strength = clampNumber(
    value,
    0,
    100,
    PHOTO_EDITOR_AUTO_ENHANCE_DEFAULT_STRENGTH
  );
  const normalizedState = normalizePhotoEditorAutoEnhanceState(
    photoEditorState.autoEnhance
  );

  if (!normalizedState.enabled) {
    photoEditorState.autoEnhance = {
      ...normalizedState,
      strength,
    };
    syncPhotoEditorAutoEnhanceControls();
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.autoEnhance = {
    ...normalizedState,
    strength,
  };

  photoEditorState.values = blendPhotoEditValues(
    normalizedState.originalValuesBeforeAuto,
    normalizedState.generatedValues,
    normalizePhotoEditorAutoStrength(strength)
  );
  syncPhotoEditorAdjustmentControls();
  clearPhotoEditorAdjustmentLivePreview();
  schedulePhotoEditorRender({
    debounceMs: PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS,
    interactive: true,
  });

  syncPhotoEditorAutoEnhanceControls();
  schedulePhotoEditorHistoryCommit();
}

function applyPhotoEditorPreset(presetKey) {
  if (!photoEditorState) {
    return;
  }

  const preset = getPhotoEditorPreset(presetKey);

  if (!preset) {
    return;
  }

  applyPhotoEditorAutoEnhancePreset(presetKey, preset);
}

function updatePhotoEditorSliderValue(key, value) {
  if (!photoEditorState) {
    return;
  }

  const slider = getPhotoEditorSliderMeta(key);

  if (!slider) {
    return;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.values[key] = clampNumber(
    value,
    slider.min,
    slider.max,
    slider.defaultValue
  );
  if (photoEditorState.autoEnhance?.enabled) {
    photoEditorState.autoEnhance = {
      ...normalizePhotoEditorAutoEnhanceState(photoEditorState.autoEnhance),
      generatedValues: clonePhotoEditValues(photoEditorState.values),
    };
  }
  syncPhotoEditorAdjustmentControls(key);
  syncPhotoEditorAutoEnhanceControls();
  applyPhotoEditorAdjustmentLivePreview();
  schedulePhotoEditorRender({
    debounceMs: PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS,
    interactive: true,
  });
  schedulePhotoEditorHistoryCommit();
}

function updatePhotoEditorCrop(nextCrop) {
  if (!photoEditorState) {
    return;
  }

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  beginPhotoEditorHistoryMutation();
  const currentCrop = normalizePhotoEditorCropState(photoEditorState.crop);

  photoEditorState.crop = {
    ...currentCrop,
    ...nextCrop,
    zoomMode: 'offset',
    zoom: clampNumber(
      nextCrop.zoom ?? currentCrop.zoom,
      0,
      PHOTO_EDITOR_CROP_ZOOM_MAX,
      0
    ),
    offsetX: clampNumber(
      nextCrop.offsetX ?? currentCrop.offsetX,
      -100,
      100,
      0
    ),
    offsetY: clampNumber(
      nextCrop.offsetY ?? currentCrop.offsetY,
      -100,
      100,
      0
    ),
    rotation: normalizePhotoEditorCropRotation(
      nextCrop.rotation ?? currentCrop.rotation
    ),
    flipX: Boolean(nextCrop.flipX ?? currentCrop.flipX),
    flipY: Boolean(nextCrop.flipY ?? currentCrop.flipY),
    tilt: clampNumber(nextCrop.tilt ?? currentCrop.tilt, -45, 45, 0),
  };
  photoEditorState.isCropInteractivePreview = true;
  syncPhotoEditorCropControls();
  schedulePhotoEditorRender({
    debounceMs: PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS,
    interactive: true,
  });
  schedulePhotoEditorHistoryCommit();
}

function applyPhotoEditorCropPreset(presetKey) {
  if (!photoEditorState) {
    return;
  }

  const preset = getPhotoEditorCropPreset(presetKey);

  if (photoEditorState.draftMask || photoEditorState.maskTool !== 'none') {
    cancelPhotoEditorPendingMask();
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.crop = normalizePhotoEditorCropState({
    ...photoEditorState.crop,
    preset: preset.key,
    zoom: preset.transparentPadding ? 0 : photoEditorState.crop?.zoom,
    offsetX: 0,
    offsetY: 0,
  });

  if (preset.exportMaxEdge || preset.exportFormat) {
    photoEditorState.exportSettings = normalizePhotoEditorExportSettings({
      ...photoEditorState.exportSettings,
      format: preset.exportFormat || photoEditorState.exportSettings?.format,
      maxEdge: preset.exportMaxEdge || photoEditorState.exportSettings?.maxEdge,
    });
  }

  syncPhotoEditorCropControls();
  syncPhotoEditorExportControls();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function getPhotoEditorPreviewMaxEdge() {
  const bounds = photoEditorCanvasWrap?.getBoundingClientRect();
  const deviceScale = Math.max(1, Math.min(2, window.devicePixelRatio || 1));
  const availableEdge =
    bounds && bounds.width > 0 && bounds.height > 0
      ? Math.max(bounds.width, bounds.height) * deviceScale
      : PHOTO_EDITOR_PREVIEW_MAX_EDGE;
  const dragMode = String(photoEditorState?.dragMode || '');
  const isInteractive =
    Boolean(photoEditorState?.isInteractivePreview) ||
    Boolean(dragMode);
  const hasHeavyInteractiveEffect =
    isInteractive &&
    (
      dragMode === 'mask' ||
      dragMode.startsWith('blur-') ||
      Number(photoEditorState?.blur?.amount) > 0 ||
      Boolean(photoEditorState?.draftMask) ||
      (Array.isArray(photoEditorState?.masks) && photoEditorState.masks.length > 0)
    );
  const isCropInteractive =
    isInteractive &&
    (dragMode === 'pan' || Boolean(photoEditorState?.isCropInteractivePreview));
  const maxEdge = dragMode === 'curve'
    ? PHOTO_EDITOR_CURVE_DRAG_PREVIEW_MAX_EDGE
    : isCropInteractive
      ? PHOTO_EDITOR_CROP_INTERACTIVE_PREVIEW_MAX_EDGE
    : hasHeavyInteractiveEffect
      ? PHOTO_EDITOR_HEAVY_INTERACTIVE_PREVIEW_MAX_EDGE
    : isInteractive
      ? PHOTO_EDITOR_INTERACTIVE_PREVIEW_MAX_EDGE
      : PHOTO_EDITOR_PREVIEW_MAX_EDGE;
  const minEdge = hasHeavyInteractiveEffect
    ? isCropInteractive
      ? 360
      : 640
    : isCropInteractive
      ? 360
    : isInteractive
      ? 760
      : 900;

  return Math.max(
    minEdge,
    Math.min(maxEdge, Math.round(availableEdge))
  );
}

function handlePhotoEditorCanvasWheel(event) {
  if (!photoEditorState?.sourceImage) {
    return;
  }

  event.preventDefault();

  const currentZoom = clampNumber(
    photoEditorState.crop.zoom,
    0,
    PHOTO_EDITOR_CROP_ZOOM_MAX,
    0
  );
  const currentScale = 1 + currentZoom / 100;
  const zoomFactor = event.deltaY < 0 ? 1.08 : 1 / 1.08;
  const nextZoom = Math.round((currentScale * zoomFactor - 1) * 100);

  updatePhotoEditorCrop({
    zoom: clampNumber(nextZoom, 0, PHOTO_EDITOR_CROP_ZOOM_MAX, 0),
  });
}

function rotatePhotoEditorCrop(deltaDegrees) {
  if (!photoEditorState) {
    return;
  }

  updatePhotoEditorCrop({
    rotation: normalizePhotoEditorCropRotation(
      (photoEditorState.crop.rotation || 0) + deltaDegrees
    ),
  });
}

function togglePhotoEditorCropFlipX() {
  if (!photoEditorState) {
    return;
  }

  updatePhotoEditorCrop({
    flipX: !photoEditorState.crop.flipX,
  });
}

function togglePhotoEditorCropFlipY() {
  if (!photoEditorState) {
    return;
  }

  updatePhotoEditorCrop({
    flipY: !photoEditorState.crop.flipY,
  });
}

function beginPhotoEditorPanDrag(point) {
  const image = photoEditorState?.sourceImage;

  if (!image || !photoEditorCanvas) {
    return;
  }

  photoEditorState.dragMode = 'pan';
  photoEditorState.dragStart = point;
  photoEditorState.dragInitialCrop = { ...photoEditorState.crop };
  photoEditorState.dragInitialSourceRect = getPhotoEditorSourceRect(image);
  photoEditorCanvas.classList.add('is-panning');
}

function getPanAdjustedCropOffset(axis, deltaRatio) {
  const image = photoEditorState?.sourceImage;
  const initialCrop = photoEditorState?.dragInitialCrop;
  const initialRect = photoEditorState?.dragInitialSourceRect;

  if (isPhotoEditorTransparentPaddingCrop(initialCrop)) {
    const outputSize = getPhotoEditorOutputSize(initialRect, null, initialCrop);
    const geometry = getPhotoEditorCropRenderGeometry(
      initialRect,
      outputSize.width,
      outputSize.height,
      initialCrop
    );
    const overflow = axis === 'x' ? geometry.overflowX : geometry.overflowY;

    if (!Number.isFinite(overflow) || overflow <= 0) {
      return 0;
    }

    const initialOffset = axis === 'x'
      ? Number(initialCrop?.offsetX) || 0
      : Number(initialCrop?.offsetY) || 0;
    const outputLength = axis === 'x' ? outputSize.width : outputSize.height;
    const nextTranslate = overflow * initialOffset / 100 + deltaRatio * outputLength;

    return clampNumber(nextTranslate / overflow * 100, -100, 100, 0);
  }

  const imageSize =
    axis === 'x'
      ? Number(image?.naturalWidth || image?.width) || 0
      : Number(image?.naturalHeight || image?.height) || 0;
  const cropSize = axis === 'x' ? initialRect?.width || 0 : initialRect?.height || 0;
  const maxOffset = Math.max(0, (imageSize - cropSize) / 2);

  if (!imageSize || !cropSize || maxOffset <= 0) {
    return 0;
  }

  const initialOffset = axis === 'x'
    ? Number(initialCrop?.offsetX) || 0
    : Number(initialCrop?.offsetY) || 0;
  const currentCenter = imageSize / 2 + maxOffset * initialOffset / 100;
  const nextCenter = currentCenter - deltaRatio * cropSize;

  return clampNumber((nextCenter - imageSize / 2) / maxOffset * 100, -100, 100, 0);
}

function updatePhotoEditorPanDrag(point) {
  if (
    !photoEditorState ||
    photoEditorState.dragMode !== 'pan' ||
    !photoEditorState.dragStart
  ) {
    return;
  }

  const deltaX = point.x - photoEditorState.dragStart.x;
  const deltaY = point.y - photoEditorState.dragStart.y;

  updatePhotoEditorCrop({
    offsetX: Math.round(getPanAdjustedCropOffset('x', deltaX)),
    offsetY: Math.round(getPanAdjustedCropOffset('y', deltaY)),
  });
}

function finishPhotoEditorPanDrag() {
  if (!photoEditorState || photoEditorState.dragMode !== 'pan') {
    return;
  }

  photoEditorState.dragMode = null;
  photoEditorState.dragStart = null;
  photoEditorState.dragInitialCrop = null;
  photoEditorState.dragInitialSourceRect = null;
  photoEditorCanvas?.classList.remove('is-panning');
  finishPhotoEditorInteractivePreview();
}

function canPhotoEditorDragTextOverlay() {
  return (
    getPhotoEditorTextCollectionFromState().textOverlays.length > 0 &&
    photoEditorState?.maskTool === 'none' &&
    !photoEditorState?.draftMask &&
    isPhotoEditorAccordionOpen('text')
  );
}

function canPhotoEditorDragImageOverlay() {
  return (
    getPhotoEditorImageOverlayCollectionFromState().imageOverlays.length > 0 &&
    photoEditorState?.maskTool === 'none' &&
    !photoEditorState?.draftMask &&
    isPhotoEditorAccordionOpen('imageOverlay')
  );
}

function getPhotoEditorImageOverlayDisplayMetrics(overlay) {
  const displaySize = getPhotoEditorCanvasDisplaySize();
  return getPhotoEditorImageOverlayCanvasMetrics(
    displaySize.width,
    displaySize.height,
    overlay
  );
}

function getPhotoEditorImageOverlayLocalPoint(point, metrics) {
  const displaySize = getPhotoEditorCanvasDisplaySize();
  const pointX = clampNumber(point?.x, -2, 3, 0) * displaySize.width;
  const pointY = clampNumber(point?.y, -2, 3, 0) * displaySize.height;

  return {
    x: pointX - metrics.centerX,
    y: pointY - metrics.centerY,
    canvasX: pointX,
    canvasY: pointY,
  };
}

function getPhotoEditorImageOverlayDragInfo(point) {
  if (!canPhotoEditorDragImageOverlay()) {
    return null;
  }

  const collection = getPhotoEditorImageOverlayCollectionFromState();
  const activeOverlay = collection.imageOverlays.find(
    (overlay) => overlay.id === collection.activeImageOverlayId
  );
  const orderedOverlays = [
    ...(activeOverlay ? [activeOverlay] : []),
    ...collection.imageOverlays
      .filter((overlay) => overlay.id !== activeOverlay?.id)
      .reverse(),
  ];

  for (const overlay of orderedOverlays) {
    const metrics = getPhotoEditorImageOverlayDisplayMetrics(overlay);

    if (!metrics) {
      continue;
    }

    const handles = getPhotoEditorImageOverlayHandles(metrics);
    const localPoint = getPhotoEditorImageOverlayLocalPoint(point, metrics);

    if (
      handles?.resize &&
      Math.hypot(localPoint.canvasX - handles.resize.x, localPoint.canvasY - handles.resize.y) <=
        PHOTO_EDITOR_IMAGE_OVERLAY_HANDLE_HIT_RADIUS
    ) {
      return {
        mode: 'image-overlay-resize',
        overlayId: overlay.id,
        metrics,
      };
    }

    if (
      Math.abs(localPoint.x) <= metrics.width / 2 &&
      Math.abs(localPoint.y) <= metrics.height / 2
    ) {
      return {
        mode: 'image-overlay-move',
        overlayId: overlay.id,
        metrics,
      };
    }
  }

  return null;
}

function translatePhotoEditorImageOverlay(overlay, deltaX, deltaY) {
  if (!overlay) {
    return null;
  }

  return normalizePhotoEditorImageOverlayState({
    ...overlay,
    x: overlay.x + deltaX,
    y: overlay.y + deltaY,
  });
}

function snapPhotoEditorImageOverlayToGuides(overlay) {
  if (!overlay) {
    return {
      overlay,
      snapX: null,
      snapY: null,
    };
  }

  const normalizedOverlay = normalizePhotoEditorImageOverlayState(overlay);
  const halfWidth = normalizedOverlay.width / 2;
  const halfHeight = normalizedOverlay.height / 2;
  const snapX = getPhotoEditorBestAxisSnap(
    [
      normalizedOverlay.x,
      normalizedOverlay.x - halfWidth,
      normalizedOverlay.x + halfWidth,
    ],
    'x'
  );
  const snapY = getPhotoEditorBestAxisSnap(
    [
      normalizedOverlay.y,
      normalizedOverlay.y - halfHeight,
      normalizedOverlay.y + halfHeight,
    ],
    'y'
  );
  const snappedOverlay =
    snapX || snapY
      ? translatePhotoEditorImageOverlay(
          normalizedOverlay,
          snapX ? snapX.value - snapX.originalValue : 0,
          snapY ? snapY.value - snapY.originalValue : 0
        )
      : normalizedOverlay;

  setPhotoEditorSnapGuideFromAxisSnaps(snapX, snapY);

  return {
    overlay: snappedOverlay,
    snapX,
    snapY,
  };
}

function resizePhotoEditorImageOverlay(
  overlay,
  point,
  { keepRatio = false } = {}
) {
  if (!overlay) {
    return null;
  }

  const normalizedOverlay = normalizePhotoEditorImageOverlayState(overlay);
  const metrics = getPhotoEditorImageOverlayDisplayMetrics(normalizedOverlay);

  if (!metrics) {
    return normalizedOverlay;
  }

  const displaySize = getPhotoEditorCanvasDisplaySize();
  const localPoint = getPhotoEditorImageOverlayLocalPoint(point, metrics);
  let nextWidth = Math.max(
    PHOTO_EDITOR_IMAGE_OVERLAY_MIN_SIZE,
    Math.abs(localPoint.x) * 2 / displaySize.width
  );
  let nextHeight = Math.max(
    PHOTO_EDITOR_IMAGE_OVERLAY_MIN_SIZE,
    Math.abs(localPoint.y) * 2 / displaySize.height
  );

  if (keepRatio) {
    const currentRatio =
      normalizedOverlay.height > 0
        ? normalizedOverlay.width / normalizedOverlay.height
        : getPhotoEditorImageOverlayAspectRatio(normalizedOverlay);

    if (Math.abs(localPoint.x) >= Math.abs(localPoint.y) * currentRatio) {
      nextHeight = nextWidth / Math.max(0.01, currentRatio);
    } else {
      nextWidth = nextHeight * currentRatio;
    }
  }

  return normalizePhotoEditorImageOverlayState({
    ...normalizedOverlay,
    width: clampNumber(nextWidth, PHOTO_EDITOR_IMAGE_OVERLAY_MIN_SIZE, 2, normalizedOverlay.width),
    height: clampNumber(nextHeight, PHOTO_EDITOR_IMAGE_OVERLAY_MIN_SIZE, 2, normalizedOverlay.height),
  });
}

function beginPhotoEditorImageOverlayDrag(point, dragInfo) {
  if (!photoEditorState || !point || !dragInfo?.overlayId) {
    return;
  }

  const collection = getPhotoEditorImageOverlayCollectionFromState();
  const initialOverlay = collection.imageOverlays.find(
    (overlay) => overlay.id === dragInfo.overlayId
  );

  if (!initialOverlay) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.dragMode = dragInfo.mode || 'image-overlay-move';
  photoEditorState.dragStart = point;
  photoEditorState.dragInitialImageOverlay = initialOverlay;
  photoEditorState.activeImageOverlayId = initialOverlay.id;
  photoEditorState.activeTextId = '';
  photoEditorCanvas?.classList.add('is-panning');
  syncPhotoEditorTextControls();
  syncPhotoEditorImageOverlayControls();
  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender({
      debounceMs: PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS,
      interactive: true,
    });
  }
}

function updatePhotoEditorImageOverlayDrag(point, event = null) {
  if (
    !photoEditorState ||
    !String(photoEditorState.dragMode || '').startsWith('image-overlay-') ||
    !photoEditorState.dragStart ||
    !photoEditorState.dragInitialImageOverlay
  ) {
    return;
  }

  const collection = getPhotoEditorImageOverlayCollectionFromState();
  let nextOverlay = photoEditorState.dragInitialImageOverlay;

  if (photoEditorState.dragMode === 'image-overlay-resize') {
    nextOverlay = snapPhotoEditorImageOverlayToGuides(
      resizePhotoEditorImageOverlay(photoEditorState.dragInitialImageOverlay, point, {
        keepRatio: Boolean(event?.shiftKey),
      })
    ).overlay;
  } else {
    const deltaX = point.x - photoEditorState.dragStart.x;
    const deltaY = point.y - photoEditorState.dragStart.y;
    nextOverlay = snapPhotoEditorImageOverlayToGuides(
      translatePhotoEditorImageOverlay(
        photoEditorState.dragInitialImageOverlay,
        deltaX,
        deltaY
      )
    ).overlay;
  }

  setPhotoEditorImageOverlayCollection(
    collection.imageOverlays.map((overlay) =>
      overlay.id === photoEditorState.dragInitialImageOverlay.id
        ? nextOverlay
        : overlay
    ),
    nextOverlay.id
  );

  syncPhotoEditorImageOverlayControls();
  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender({
      debounceMs: PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS,
      interactive: true,
    });
  }
}

function finishPhotoEditorImageOverlayDrag() {
  if (
    !photoEditorState ||
    !String(photoEditorState.dragMode || '').startsWith('image-overlay-')
  ) {
    return;
  }

  photoEditorState.dragMode = null;
  photoEditorState.dragStart = null;
  photoEditorState.dragInitialImageOverlay = null;
  photoEditorState.snapGuide = null;
  photoEditorCanvas?.classList.remove('is-panning');
  syncPhotoEditorImageOverlayControls();
  finishPhotoEditorInteractivePreview();
  commitPhotoEditorHistoryMutation();
}

function getPhotoEditorCanvasDisplaySize() {
  const bounds = photoEditorCanvas?.getBoundingClientRect();

  return {
    width: Math.max(1, Number(bounds?.width) || Number(photoEditorCanvas?.width) || 1),
    height: Math.max(1, Number(bounds?.height) || Number(photoEditorCanvas?.height) || 1),
  };
}

function getPhotoEditorRulerGuideDragMode(point) {
  if (!photoEditorState?.showRulers || !point || !photoEditorCanvas) {
    return null;
  }

  const displaySize = getPhotoEditorCanvasDisplaySize();
  const edge = getPhotoEditorRulerEdgeSize(displaySize.width, displaySize.height);
  const inTopRuler = point.y >= 0 && point.y <= edge / displaySize.height;
  const inLeftRuler = point.x >= 0 && point.x <= edge / displaySize.width;

  if (inTopRuler && !inLeftRuler) {
    return 'guide-y';
  }

  if (inLeftRuler && !inTopRuler) {
    return 'guide-x';
  }

  return null;
}

function getPhotoEditorExistingRulerGuideDragInfo(point) {
  if (!photoEditorState?.showRulers || !point || !photoEditorCanvas) {
    return null;
  }

  const displaySize = getPhotoEditorCanvasDisplaySize();
  const guides = normalizePhotoEditorRulerGuides(photoEditorState.rulerGuides);
  const hitRadiusPx = Math.max(
    8,
    Math.min(displaySize.width, displaySize.height) * 0.012
  );
  let bestGuide = null;
  let bestDistance = Infinity;

  guides.x.forEach((position, index) => {
    const distance = Math.abs(point.x - position) * displaySize.width;

    if (distance <= hitRadiusPx && distance < bestDistance) {
      bestDistance = distance;
      bestGuide = {
        axis: 'x',
        index,
        position,
      };
    }
  });

  guides.y.forEach((position, index) => {
    const distance = Math.abs(point.y - position) * displaySize.height;

    if (distance <= hitRadiusPx && distance < bestDistance) {
      bestDistance = distance;
      bestGuide = {
        axis: 'y',
        index,
        position,
      };
    }
  });

  return bestGuide;
}

function getPhotoEditorRulerGuideAxisFromDragMode(dragMode) {
  return dragMode === 'guide-x'
    ? 'x'
    : dragMode === 'guide-y'
      ? 'y'
      : '';
}

function getPhotoEditorRulerGuidePositionFromPoint(point, axis) {
  return clampNumber(axis === 'x' ? point?.x : point?.y, 0, 1, 0.5);
}

function addPhotoEditorRulerGuide(axis, position) {
  if (!photoEditorState || !['x', 'y'].includes(axis)) {
    return false;
  }

  const normalizedGuides = normalizePhotoEditorRulerGuides(
    photoEditorState.rulerGuides
  );
  const normalizedPosition = clampNumber(position, 0, 1, 0.5);
  const nextGuides = normalizedGuides[axis].filter(
    (guidePosition) => Math.abs(guidePosition - normalizedPosition) > 0.003
  );

  nextGuides.push(normalizedPosition);
  photoEditorState.rulerGuides = normalizePhotoEditorRulerGuides({
    ...normalizedGuides,
    [axis]: nextGuides,
  });

  return true;
}

function updatePhotoEditorRulerGuideAtIndex(axis, index, position) {
  if (!photoEditorState || !['x', 'y'].includes(axis)) {
    return false;
  }

  const normalizedGuides = normalizePhotoEditorRulerGuides(
    photoEditorState.rulerGuides
  );
  const guides = [...normalizedGuides[axis]];

  if (index < 0 || index >= guides.length) {
    return false;
  }

  guides[index] = clampNumber(position, 0, 1, guides[index]);
  photoEditorState.rulerGuides = normalizePhotoEditorRulerGuides({
    ...normalizedGuides,
    [axis]: guides,
  });

  return true;
}

function beginPhotoEditorRulerGuideDrag(point, dragMode, existingGuide = null) {
  const axis = getPhotoEditorRulerGuideAxisFromDragMode(dragMode);

  if (!photoEditorState || !axis) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.dragMode = dragMode;
  photoEditorState.dragStart = point;
  photoEditorState.draftRulerGuide = {
    axis,
    position: existingGuide
      ? existingGuide.position
      : getPhotoEditorRulerGuidePositionFromPoint(point, axis),
  };
  photoEditorState.dragInitialRulerGuide = existingGuide
    ? {
        axis: existingGuide.axis,
        index: existingGuide.index,
        position: existingGuide.position,
        pointerOffset:
          getPhotoEditorRulerGuidePositionFromPoint(point, axis) -
          existingGuide.position,
      }
    : null;
  photoEditorState.snapGuide = null;
  photoEditorCanvas?.classList.add('is-panning');
  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender({ interactive: true });
  }
}

function updatePhotoEditorRulerGuideDrag(point) {
  const axis = getPhotoEditorRulerGuideAxisFromDragMode(
    photoEditorState?.dragMode
  );

  if (!photoEditorState || !axis || !point) {
    return;
  }

  const initialGuide = photoEditorState.dragInitialRulerGuide;

  photoEditorState.draftRulerGuide = {
    axis,
    position: clampNumber(
      getPhotoEditorRulerGuidePositionFromPoint(point, axis) -
        (Number(initialGuide?.pointerOffset) || 0),
      0,
      1,
      0.5
    ),
  };
  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender({ interactive: true });
  }
}

function finishPhotoEditorRulerGuideDrag() {
  const axis = getPhotoEditorRulerGuideAxisFromDragMode(
    photoEditorState?.dragMode
  );

  if (!photoEditorState || !axis) {
    return;
  }

  const guide = photoEditorState.draftRulerGuide;
  const initialGuide = photoEditorState.dragInitialRulerGuide;
  const displaySize = getPhotoEditorCanvasDisplaySize();
  const movedPixels = axis === 'x'
    ? Math.abs((guide?.position || 0) - (photoEditorState.dragStart?.x || 0)) *
      displaySize.width
    : Math.abs((guide?.position || 0) - (photoEditorState.dragStart?.y || 0)) *
      displaySize.height;

  if (guide && initialGuide) {
    updatePhotoEditorRulerGuideAtIndex(
      axis,
      initialGuide.index,
      guide.position
    );
  } else if (guide && movedPixels >= PHOTO_EDITOR_RULER_GUIDE_MIN_DRAG_PX) {
    addPhotoEditorRulerGuide(axis, guide.position);
  }

  photoEditorState.dragMode = null;
  photoEditorState.dragStart = null;
  photoEditorState.draftRulerGuide = null;
  photoEditorState.dragInitialRulerGuide = null;
  photoEditorState.snapGuide = null;
  photoEditorCanvas?.classList.remove('is-panning');
  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender();
  }
  commitPhotoEditorHistoryMutation();
}

function getPhotoEditorTextLocalPoint(point, metrics, width, height) {
  const pointX = clampNumber(point?.x, -2, 3, 0) * width;
  const pointY = clampNumber(point?.y, -2, 3, 0) * height;
  const cos = Math.cos(-metrics.rotationRadians);
  const sin = Math.sin(-metrics.rotationRadians);
  const deltaX = pointX - metrics.centerX;
  const deltaY = pointY - metrics.centerY;

  return {
    x: deltaX * cos - deltaY * sin,
    y: deltaX * sin + deltaY * cos,
    canvasX: pointX,
    canvasY: pointY,
  };
}

function getPhotoEditorTextDragInfo(point) {
  if (!photoEditorCanvas || !canPhotoEditorDragTextOverlay()) {
    return null;
  }

  const width = Math.max(1, photoEditorCanvas.width || 1);
  const height = Math.max(1, photoEditorCanvas.height || 1);
  const ctx = photoEditorCanvas.getContext('2d');

  if (!ctx) {
    return null;
  }

  const collection = getPhotoEditorTextCollectionFromState();
  const activeText = collection.textOverlays.find(
    (textOverlay) => textOverlay.id === collection.activeTextId
  );
  const orderedTextOverlays = [
    ...(activeText ? [activeText] : []),
    ...collection.textOverlays
      .filter((textOverlay) => textOverlay.id !== activeText?.id)
      .reverse(),
  ];
  const pointX = clampNumber(point?.x, -2, 3, 0) * width;
  const pointY = clampNumber(point?.y, -2, 3, 0) * height;

  for (const textOverlay of orderedTextOverlays) {
    const metrics = getPhotoEditorTextCanvasMetrics(
      ctx,
      width,
      height,
      textOverlay
    );

    if (!metrics) {
      continue;
    }

    const handles = getPhotoEditorTextHandles(metrics);
    const handleRadius = Math.max(
      PHOTO_EDITOR_MASK_HANDLE_HIT_RADIUS,
      Math.min(34, metrics.height * 0.24)
    );

    if (
      handles?.rotate &&
      Math.hypot(pointX - handles.rotate.x, pointY - handles.rotate.y) <=
        handleRadius
    ) {
      return {
        mode: 'text-rotate',
        textId: textOverlay.id,
        metrics,
      };
    }

    const localPoint = getPhotoEditorTextLocalPoint(point, metrics, width, height);

    if (
      Math.abs(localPoint.x) <= metrics.width / 2 &&
      Math.abs(localPoint.y) <= metrics.height / 2
    ) {
      return {
        mode: 'text-move',
        textId: textOverlay.id,
        metrics,
      };
    }
  }

  return null;
}

function beginPhotoEditorTextDrag(point, dragInfo) {
  if (!photoEditorState || !point || !dragInfo?.textId) {
    return;
  }

  const collection = getPhotoEditorTextCollectionFromState();
  const initialText = collection.textOverlays.find(
    (textOverlay) => textOverlay.id === dragInfo.textId
  );

  if (!initialText) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.dragMode = dragInfo.mode || 'text-move';
  photoEditorState.dragStart = point;
  photoEditorState.dragInitialText = initialText;
  photoEditorState.activeTextId = initialText.id;
  photoEditorState.activeImageOverlayId = '';
  photoEditorCanvas?.classList.add('is-panning');
  syncPhotoEditorTextControls();
  syncPhotoEditorImageOverlayControls();
  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender({
      debounceMs: PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS,
      interactive: true,
    });
  }
}

function updatePhotoEditorTextDrag(point) {
  if (
    !photoEditorState ||
    !String(photoEditorState.dragMode || '').startsWith('text-') ||
    !photoEditorState.dragStart ||
    !photoEditorState.dragInitialText
  ) {
    return;
  }

  const collection = getPhotoEditorTextCollectionFromState();
  let nextText = photoEditorState.dragInitialText;

  if (photoEditorState.dragMode === 'text-rotate') {
    photoEditorState.snapGuide = null;
    const displaySize = getPhotoEditorCanvasDisplaySize();
    const centerX = photoEditorState.dragInitialText.x * displaySize.width;
    const centerY = photoEditorState.dragInitialText.y * displaySize.height;
    const getAngle = (targetPoint) =>
      Math.atan2(
        targetPoint.y * displaySize.height - centerY,
        targetPoint.x * displaySize.width - centerX
      );
    const deltaDegrees =
      (getAngle(point) - getAngle(photoEditorState.dragStart)) * 180 / Math.PI;

    nextText = normalizePhotoEditorTextState({
      ...photoEditorState.dragInitialText,
      rotation: photoEditorState.dragInitialText.rotation + deltaDegrees,
    });
  } else {
    const deltaX = point.x - photoEditorState.dragStart.x;
    const deltaY = point.y - photoEditorState.dragStart.y;
    nextText = snapPhotoEditorTextToGuides({
      ...photoEditorState.dragInitialText,
      x: photoEditorState.dragInitialText.x + deltaX,
      y: photoEditorState.dragInitialText.y + deltaY,
    });
  }

  setPhotoEditorTextCollection(
    collection.textOverlays.map((textOverlay) =>
      textOverlay.id === photoEditorState.dragInitialText.id ? nextText : textOverlay
    ),
    nextText.id
  );

  if (!paintPhotoEditorPreviewOverlayOnly()) {
    schedulePhotoEditorRender({
      debounceMs: PHOTO_EDITOR_INTERACTIVE_PREVIEW_DEBOUNCE_MS,
      interactive: true,
    });
  }
}

function finishPhotoEditorTextDrag() {
  if (
    !photoEditorState ||
    !String(photoEditorState.dragMode || '').startsWith('text-')
  ) {
    return;
  }

  photoEditorState.dragMode = null;
  photoEditorState.dragStart = null;
  photoEditorState.dragInitialText = null;
  photoEditorState.snapGuide = null;
  photoEditorCanvas?.classList.remove('is-panning');
  syncPhotoEditorTextControls();
  finishPhotoEditorInteractivePreview();
  commitPhotoEditorHistoryMutation();
}

function snapPhotoEditorPointToGuides(
  point,
  { allowOutside = false, includeCenter = true, updateGuide = true } = {}
) {
  const bounds = getPhotoEditorMaskCoordinateBounds(allowOutside);
  const fallbackX = Number(point?.x) || 0;
  const fallbackY = Number(point?.y) || 0;
  const snapX = getPhotoEditorAxisSnap(point?.x, 'x', { includeCenter });
  const snapY = getPhotoEditorAxisSnap(point?.y, 'y', { includeCenter });
  const snappedPoint = {
    x: clampNumber(snapX.value, bounds.min, bounds.max, fallbackX),
    y: clampNumber(snapY.value, bounds.min, bounds.max, fallbackY),
  };

  if (updateGuide) {
    setPhotoEditorSnapGuideFromAxisSnaps(snapX, snapY);
  }

  return {
    point: snappedPoint,
    snapX,
    snapY,
  };
}

function getPhotoEditorBestAxisSnap(candidates, axis) {
  let bestSnap = null;

  for (const candidate of Array.isArray(candidates) ? candidates : []) {
    const snap = getPhotoEditorAxisSnap(candidate, axis);

    if (
      snap.snapped &&
      (!bestSnap || snap.distance < bestSnap.distance)
    ) {
      bestSnap = snap;
    }
  }

  return bestSnap;
}

function snapPhotoEditorMaskToGuides(mask) {
  if (!mask) {
    return {
      mask,
      snapX: null,
      snapY: null,
    };
  }

  const rect = getPhotoEditorMaskNormalizedRect(mask);
  const center = getPhotoEditorMaskRectCenter(rect);
  const snapX = getPhotoEditorBestAxisSnap(
    [center.x, rect.x, rect.x + rect.width],
    'x'
  );
  const snapY = getPhotoEditorBestAxisSnap(
    [center.y, rect.y, rect.y + rect.height],
    'y'
  );
  const snappedMask =
    snapX || snapY
      ? translatePhotoEditorMask(
          mask,
          snapX ? snapX.value - snapX.originalValue : 0,
          snapY ? snapY.value - snapY.originalValue : 0
        )
      : mask;

  setPhotoEditorSnapGuideFromAxisSnaps(snapX, snapY);

  return {
    mask: snappedMask,
    snapX,
    snapY,
  };
}

function getPhotoEditorRadialBlurDragMode(point) {
  if (
    photoEditorState?.blur?.mode !== 'radial' ||
    photoEditorState.blur.isConfirmed ||
    !isPhotoEditorAccordionOpen('blur')
  ) {
    return null;
  }

  const blur = normalizePhotoEditorBlurState(photoEditorState.blur);
  const displaySize = getPhotoEditorCanvasDisplaySize();
  const minEdge = Math.max(1, Math.min(displaySize.width, displaySize.height));
  const centerX = blur.centerX * displaySize.width;
  const centerY = blur.centerY * displaySize.height;
  const pointX = clampNumber(point?.x, 0, 1, 0) * displaySize.width;
  const pointY = clampNumber(point?.y, 0, 1, 0) * displaySize.height;
  const distance = Math.hypot(pointX - centerX, pointY - centerY);
  const radius = blur.radius * minEdge;
  const outerRadius = blur.outerRadius * minEdge;
  const edgeTolerance = Math.max(14, minEdge * 0.025);

  if (Math.abs(distance - outerRadius) <= edgeTolerance) {
    return 'blur-outer-radius';
  }

  if (Math.abs(distance - radius) <= edgeTolerance) {
    return 'blur-inner-radius';
  }

  return distance <= outerRadius ? 'blur-center' : 'blur-center-snap';
}

function beginPhotoEditorBlurDrag(point, dragMode) {
  if (!photoEditorState || !dragMode) {
    return;
  }

  beginPhotoEditorHistoryMutation();

  if (dragMode === 'blur-center-snap') {
    photoEditorState.blur = normalizePhotoEditorBlurState({
      ...photoEditorState.blur,
      centerX: point.x,
      centerY: point.y,
    });
  }

  photoEditorState.dragMode = dragMode;
  photoEditorState.dragStart = point;
  photoEditorState.dragInitialBlur = normalizePhotoEditorBlurState(
    photoEditorState.blur
  );
  photoEditorCanvas?.classList.add('is-panning');
  syncPhotoEditorBlurControls();
  paintPhotoEditorPreviewOverlayOnly();
  schedulePhotoEditorRender({
    debounceMs: PHOTO_EDITOR_EFFECT_DRAG_PREVIEW_DEBOUNCE_MS,
    interactive: true,
  });
}

function updatePhotoEditorBlurDrag(point) {
  if (
    !photoEditorState ||
    !photoEditorState.dragStart ||
    !photoEditorState.dragInitialBlur
  ) {
    return;
  }

  const initialBlur = photoEditorState.dragInitialBlur;

  if (photoEditorState.dragMode === 'blur-inner-radius') {
    const displaySize = getPhotoEditorCanvasDisplaySize();
    const minEdge = Math.max(1, Math.min(displaySize.width, displaySize.height));
    const distance = Math.hypot(
      (point.x - initialBlur.centerX) * displaySize.width,
      (point.y - initialBlur.centerY) * displaySize.height
    );

    photoEditorState.blur = normalizePhotoEditorBlurState({
      ...photoEditorState.blur,
      radius: Math.min(
        distance / minEdge,
        initialBlur.outerRadius - PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER
      ),
      outerRadius: initialBlur.outerRadius,
    });
  } else if (photoEditorState.dragMode === 'blur-outer-radius') {
    const displaySize = getPhotoEditorCanvasDisplaySize();
    const minEdge = Math.max(1, Math.min(displaySize.width, displaySize.height));
    const distance = Math.hypot(
      (point.x - initialBlur.centerX) * displaySize.width,
      (point.y - initialBlur.centerY) * displaySize.height
    );

    photoEditorState.blur = normalizePhotoEditorBlurState({
      ...photoEditorState.blur,
      outerRadius: Math.max(
        distance / minEdge,
        initialBlur.radius + PHOTO_EDITOR_RADIAL_BLUR_MIN_FEATHER
      ),
      radius: initialBlur.radius,
    });
  } else {
    const deltaX = point.x - photoEditorState.dragStart.x;
    const deltaY = point.y - photoEditorState.dragStart.y;

    photoEditorState.blur = normalizePhotoEditorBlurState({
      ...photoEditorState.blur,
      centerX: initialBlur.centerX + deltaX,
      centerY: initialBlur.centerY + deltaY,
    });
  }

  syncPhotoEditorBlurControls();
  paintPhotoEditorPreviewOverlayOnly();
  schedulePhotoEditorRender({
    debounceMs: PHOTO_EDITOR_EFFECT_DRAG_PREVIEW_DEBOUNCE_MS,
    interactive: true,
  });
}

function finishPhotoEditorBlurDrag() {
  if (!photoEditorState || !String(photoEditorState.dragMode || '').startsWith('blur-')) {
    return;
  }

  photoEditorState.dragMode = null;
  photoEditorState.dragStart = null;
  photoEditorState.dragInitialBlur = null;
  photoEditorCanvas?.classList.remove('is-panning');
  syncPhotoEditorBlurControls();
  finishPhotoEditorInteractivePreview();
  commitPhotoEditorHistoryMutation();
}

function normalizePhotoEditorMaskRect(
  startPoint,
  endPoint,
  { constrainSquare = false, allowOutside = false } = {}
) {
  let nextEndPoint = endPoint;
  const bounds = getPhotoEditorMaskCoordinateBounds(allowOutside);

  if (constrainSquare) {
    const displaySize = getPhotoEditorCanvasDisplaySize();
    const deltaX = endPoint.x - startPoint.x;
    const deltaY = endPoint.y - startPoint.y;
    const directionX = deltaX < 0 ? -1 : 1;
    const directionY = deltaY < 0 ? -1 : 1;
    const maxWidth =
      directionX > 0 ? bounds.max - startPoint.x : startPoint.x - bounds.min;
    const maxHeight =
      directionY > 0 ? bounds.max - startPoint.y : startPoint.y - bounds.min;
    const sizeInPixels = Math.min(
      Math.abs(deltaX) * displaySize.width,
      Math.abs(deltaY) * displaySize.height,
      maxWidth * displaySize.width,
      maxHeight * displaySize.height
    );

    nextEndPoint = {
      x: startPoint.x + directionX * sizeInPixels / displaySize.width,
      y: startPoint.y + directionY * sizeInPixels / displaySize.height,
    };
  }

  const left = clampNumber(
    Math.min(startPoint.x, nextEndPoint.x),
    bounds.min,
    bounds.max,
    0
  );
  const top = clampNumber(
    Math.min(startPoint.y, nextEndPoint.y),
    bounds.min,
    bounds.max,
    0
  );
  const right = clampNumber(
    Math.max(startPoint.x, nextEndPoint.x),
    bounds.min,
    bounds.max,
    0
  );
  const bottom = clampNumber(
    Math.max(startPoint.y, nextEndPoint.y),
    bounds.min,
    bounds.max,
    0
  );

  return {
    x: left,
    y: top,
    width: Math.max(0, right - left),
    height: Math.max(0, bottom - top),
  };
}

function getPhotoEditorPointsMaskRect(points, { allowOutside = false } = {}) {
  const validPoints = (Array.isArray(points) ? points : []).filter(
    (point) => Number.isFinite(point?.x) && Number.isFinite(point?.y)
  );

  if (validPoints.length === 0) {
    return {
      x: 0,
      y: 0,
      width: 0,
      height: 0,
    };
  }

  const bounds = getPhotoEditorMaskCoordinateBounds(allowOutside);
  const xs = validPoints.map((point) =>
    clampNumber(point.x, bounds.min, bounds.max, 0)
  );
  const ys = validPoints.map((point) =>
    clampNumber(point.y, bounds.min, bounds.max, 0)
  );
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  const right = Math.max(...xs);
  const bottom = Math.max(...ys);

  return {
    x: left,
    y: top,
    width: Math.max(0, right - left),
    height: Math.max(0, bottom - top),
  };
}

function getPhotoEditorMaskNormalizedRect(mask) {
  if (mask?.shape === 'freehand') {
    return getPhotoEditorPointsMaskRect(mask.points, {
      allowOutside: canPhotoEditorMaskExtendOutside(mask),
    });
  }

  const rect = mask?.rect || {};

  return {
    x: Number(rect.x) || 0,
    y: Number(rect.y) || 0,
    width: Math.max(0, Number(rect.width) || 0),
    height: Math.max(0, Number(rect.height) || 0),
  };
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

function rotateNormalizedPointAroundCenterForSize(
  point,
  center,
  rotationRadians,
  width,
  height
) {
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
        rotateNormalizedPointAroundCenterForSize(
          corner,
          center,
          rotationRadians,
          width,
          height
        )
      );
}

function getPhotoEditorConstrainedMaskDelta(mask, deltaX, deltaY) {
  const allowOutside = canPhotoEditorMaskExtendOutside(mask);
  const bounds = getPhotoEditorMaskCoordinateBounds(allowOutside);
  const rect = getPhotoEditorMaskNormalizedRect(mask);
  const minDeltaX = bounds.min - rect.x;
  const maxDeltaX = bounds.max - (rect.x + rect.width);
  const minDeltaY = bounds.min - rect.y;
  const maxDeltaY = bounds.max - (rect.y + rect.height);

  return {
    x: clampNumber(deltaX, minDeltaX, maxDeltaX, 0),
    y: clampNumber(deltaY, minDeltaY, maxDeltaY, 0),
  };
}

function translatePhotoEditorMask(mask, deltaX, deltaY) {
  if (!mask) {
    return null;
  }

  const constrainedDelta = getPhotoEditorConstrainedMaskDelta(
    mask,
    deltaX,
    deltaY
  );
  const rect = mask.rect || {};
  const points = Array.isArray(mask.points)
    ? mask.points.map((point) => ({
        x: point.x + constrainedDelta.x,
        y: point.y + constrainedDelta.y,
      }))
    : null;

  return {
    ...mask,
    rect: {
      x: (Number(rect.x) || 0) + constrainedDelta.x,
      y: (Number(rect.y) || 0) + constrainedDelta.y,
      width: Math.max(0, Number(rect.width) || 0),
      height: Math.max(0, Number(rect.height) || 0),
    },
    points,
  };
}

function getPhotoEditorDraftMaskHandles(width, height, sourceRect = null, sourceImage = null) {
  const mask = photoEditorState?.draftMask;

  if (!mask) {
    return null;
  }

  const maskRect = getRawCanvasRectFromMask(mask, width, height, sourceRect, sourceImage);

  if (!maskRect || maskRect.width <= 0 || maskRect.height <= 0) {
    return null;
  }

  const rect = getPhotoEditorMaskNormalizedRect(mask);
  const center = getPhotoEditorMaskRectCenter(rect);
  const rotationRadians = getPhotoEditorMaskRotation(mask) * Math.PI / 180;
  const centerPoint = mapPhotoEditorMaskPointToCanvas(
    center,
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );
  const resizePoint = mapPhotoEditorMaskPointToCanvas(
    rotateNormalizedPointAroundCenterForSize(
      {
        x: rect.x + rect.width,
        y: rect.y + rect.height,
      },
      center,
      rotationRadians,
      width,
      height
    ),
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );
  const topPoint = mapPhotoEditorMaskPointToCanvas(
    rotateNormalizedPointAroundCenterForSize(
      {
        x: rect.x + rect.width / 2,
        y: rect.y,
      },
      center,
      rotationRadians,
      width,
      height
    ),
    mask,
    width,
    height,
    sourceRect,
    sourceImage
  );
  const handleMargin = PHOTO_EDITOR_MASK_HANDLE_HIT_RADIUS * 0.7;
  const rawRotatePoint = {
    x:
      topPoint.x +
      Math.sin(rotationRadians) * PHOTO_EDITOR_MASK_ROTATE_HANDLE_OFFSET,
    y:
      topPoint.y -
      Math.cos(rotationRadians) * PHOTO_EDITOR_MASK_ROTATE_HANDLE_OFFSET,
  };
  const rotatePoint = {
    x: clampNumber(rawRotatePoint.x, handleMargin, width - handleMargin, rawRotatePoint.x),
    y: clampNumber(rawRotatePoint.y, handleMargin, height - handleMargin, rawRotatePoint.y),
  };

  return {
    move: centerPoint,
    resize: resizePoint,
    rotate: rotatePoint,
    top: topPoint,
    bounds: maskRect,
  };
}

function getPhotoEditorMaskLocalPointFromCanvasPoint(mask, point) {
  const displaySize = getPhotoEditorCanvasDisplaySize();
  const rect = getPhotoEditorMaskNormalizedRect(mask);
  const center = getPhotoEditorMaskRectCenter(rect);
  const centerX = center.x * displaySize.width;
  const centerY = center.y * displaySize.height;
  const pointX = clampNumber(point?.x, -10, 10, 0) * displaySize.width;
  const pointY = clampNumber(point?.y, -10, 10, 0) * displaySize.height;
  const rotationRadians = -getPhotoEditorMaskRotation(mask) * Math.PI / 180;
  const rotatedPoint = rotatePointAroundCenter(
    { x: pointX, y: pointY },
    { x: centerX, y: centerY },
    rotationRadians
  );

  return {
    x: (rotatedPoint.x - centerX) / displaySize.width,
    y: (rotatedPoint.y - centerY) / displaySize.height,
  };
}

function resizePhotoEditorMask(mask, point, { keepRatio = false } = {}) {
  if (!mask) {
    return null;
  }

  const rect = getPhotoEditorMaskNormalizedRect(mask);
  const center = getPhotoEditorMaskRectCenter(rect);
  const localPoint = getPhotoEditorMaskLocalPointFromCanvasPoint(mask, point);
  let nextWidth = Math.max(PHOTO_EDITOR_MIN_MASK_SIZE, Math.abs(localPoint.x) * 2);
  let nextHeight = Math.max(PHOTO_EDITOR_MIN_MASK_SIZE, Math.abs(localPoint.y) * 2);

  if (keepRatio || mask.shape === 'ellipse') {
    const currentRatio = rect.height > 0 ? rect.width / rect.height : 1;

    if (Math.abs(localPoint.x) >= Math.abs(localPoint.y) * currentRatio) {
      nextHeight = nextWidth / Math.max(0.01, currentRatio);
    } else {
      nextWidth = nextHeight * currentRatio;
    }
  }

  const allowOutside = canPhotoEditorMaskExtendOutside(mask);
  const bounds = getPhotoEditorMaskCoordinateBounds(allowOutside);
  const maxWidth = Math.max(
    PHOTO_EDITOR_MIN_MASK_SIZE,
    Math.min(
      (center.x - bounds.min) * 2,
      (bounds.max - center.x) * 2
    )
  );
  const maxHeight = Math.max(
    PHOTO_EDITOR_MIN_MASK_SIZE,
    Math.min(
      (center.y - bounds.min) * 2,
      (bounds.max - center.y) * 2
    )
  );
  nextWidth = clampNumber(nextWidth, PHOTO_EDITOR_MIN_MASK_SIZE, maxWidth, rect.width);
  nextHeight = clampNumber(nextHeight, PHOTO_EDITOR_MIN_MASK_SIZE, maxHeight, rect.height);

  if (mask.shape === 'freehand' && Array.isArray(mask.points) && rect.width > 0 && rect.height > 0) {
    const widthScale = nextWidth / rect.width;
    const heightScale = nextHeight / rect.height;
    const points = mask.points.map((maskPoint) => ({
      x: center.x + (maskPoint.x - center.x) * widthScale,
      y: center.y + (maskPoint.y - center.y) * heightScale,
    }));

    return {
      ...mask,
      rect: getPhotoEditorPointsMaskRect(points, {
        allowOutside: canPhotoEditorMaskExtendOutside(mask),
      }),
      points,
    };
  }

  return {
    ...mask,
    rect: {
      x: center.x - nextWidth / 2,
      y: center.y - nextHeight / 2,
      width: nextWidth,
      height: nextHeight,
    },
  };
}

function rotatePhotoEditorMask(mask, startPoint, currentPoint) {
  if (!mask) {
    return null;
  }

  const displaySize = getPhotoEditorCanvasDisplaySize();
  const rect = getPhotoEditorMaskNormalizedRect(mask);
  const center = getPhotoEditorMaskRectCenter(rect);
  const centerX = center.x * displaySize.width;
  const centerY = center.y * displaySize.height;
  const getAngle = (point) =>
    Math.atan2(
      point.y * displaySize.height - centerY,
      point.x * displaySize.width - centerX
    );
  const deltaDegrees = (getAngle(currentPoint) - getAngle(startPoint)) * 180 / Math.PI;

  if (mask.shape === 'freehand' && Array.isArray(mask.points)) {
    const rotationRadians = deltaDegrees * Math.PI / 180;
    const points = mask.points.map((maskPoint) =>
      rotatePointAroundCenter(maskPoint, center, rotationRadians)
    );

    return {
      ...mask,
      rect: getPhotoEditorPointsMaskRect(points, {
        allowOutside: canPhotoEditorMaskExtendOutside(mask),
      }),
      points,
    };
  }

  return {
    ...mask,
    rotation: normalizePhotoEditorMaskRotation(
      getPhotoEditorMaskRotation(mask) + deltaDegrees
    ),
  };
}

function getPhotoEditorCanvasPointerPoint(event, { allowOutside = false } = {}) {
  if (!photoEditorCanvas) {
    return null;
  }

  const bounds = photoEditorCanvas.getBoundingClientRect();

  if (bounds.width <= 0 || bounds.height <= 0) {
    return null;
  }

  const boundsLimit = getPhotoEditorMaskCoordinateBounds(allowOutside);

  return {
    x: clampNumber(
      (event.clientX - bounds.left) / bounds.width,
      boundsLimit.min,
      boundsLimit.max,
      0
    ),
    y: clampNumber(
      (event.clientY - bounds.top) / bounds.height,
      boundsLimit.min,
      boundsLimit.max,
      0
    ),
  };
}

function mapPhotoEditorOutputPointToSource(
  point,
  sourceRect,
  image,
  { allowOutside = false } = {}
) {
  const imageSize = getPhotoEditorImageSize(image);
  const bounds = getPhotoEditorMaskCoordinateBounds(allowOutside);

  if (
    !sourceRect ||
    sourceRect.width <= 0 ||
    sourceRect.height <= 0 ||
    imageSize.width <= 0 ||
    imageSize.height <= 0
  ) {
    return {
      x: clampNumber(point?.x, bounds.min, bounds.max, 0),
      y: clampNumber(point?.y, bounds.min, bounds.max, 0),
    };
  }

  const outputSize = getPhotoEditorOutputSize(sourceRect, null, photoEditorState?.crop);
  const geometry = getPhotoEditorCropRenderGeometry(
    sourceRect,
    outputSize.width,
    outputSize.height,
    photoEditorState?.crop || getDefaultPhotoEditorCropState()
  );
  const outputX =
    clampNumber(point?.x, bounds.min, bounds.max, 0) * outputSize.width -
    outputSize.width / 2 -
    geometry.translateX;
  const outputY =
    clampNumber(point?.y, bounds.min, bounds.max, 0) * outputSize.height -
    outputSize.height / 2 -
    geometry.translateY;
  const cos = Math.cos(geometry.rotationRadians);
  const sin = Math.sin(geometry.rotationRadians);
  let localX = outputX * cos + outputY * sin;
  let localY = -outputX * sin + outputY * cos;

  if (geometry.flipX) {
    localX = -localX;
  }

  if (geometry.flipY) {
    localY = -localY;
  }

  const sourceX =
    sourceRect.x +
    (localX / (geometry.drawWidth * geometry.coverScale) + 0.5) *
      sourceRect.width;
  const sourceY =
    sourceRect.y +
    (localY / (geometry.drawHeight * geometry.coverScale) + 0.5) *
      sourceRect.height;

  return {
    x: clampNumber(sourceX / imageSize.width, bounds.min, bounds.max, 0),
    y: clampNumber(sourceY / imageSize.height, bounds.min, bounds.max, 0),
  };
}

function getPhotoEditorOutputMaskOutlinePoints(mask, outputSize = null) {
  if (!mask || mask.space === 'source') {
    return [];
  }

  const width = Math.max(1, Number(outputSize?.width) || 1);
  const height = Math.max(1, Number(outputSize?.height) || 1);
  const rect = getPhotoEditorMaskNormalizedRect(mask);
  const center = getPhotoEditorMaskRectCenter(rect);
  const rotationRadians = getPhotoEditorMaskRotation(mask) * Math.PI / 180;

  if (mask.shape === 'ellipse') {
    const points = [];
    const pointCount = 48;

    for (let index = 0; index < pointCount; index += 1) {
      const angle = index / pointCount * Math.PI * 2;
      const point = {
        x: center.x + Math.cos(angle) * rect.width / 2,
        y: center.y + Math.sin(angle) * rect.height / 2,
      };
      points.push(
        rotationRadians === 0
          ? point
          : rotateNormalizedPointAroundCenterForSize(
              point,
              center,
              rotationRadians,
              width,
              height
            )
      );
    }

    return points;
  }

  if (mask.shape === 'rect') {
    return getPhotoEditorRotatedMaskRectCorners(mask, width, height);
  }

  return Array.isArray(mask.points) ? mask.points : [];
}

function convertPhotoEditorMaskToSourceSpace(mask) {
  if (!photoEditorState?.sourceImage || !mask || mask.space === 'source') {
    return mask;
  }

  const sourceImage = photoEditorState.sourceImage;
  const sourceRect = getPhotoEditorSourceRect(sourceImage);
  const outputSize = getPhotoEditorOutputSize(
    sourceRect,
    null,
    photoEditorState.crop
  );
  const rect = mask.rect || {};
  const allowOutside = canPhotoEditorMaskExtendOutside(mask);
  const shouldConvertAsOutline =
    ['rect', 'ellipse'].includes(mask.shape) &&
    getPhotoEditorMaskRotation(mask) !== 0;

  if (shouldConvertAsOutline) {
    const sourcePoints = getPhotoEditorOutputMaskOutlinePoints(
      mask,
      outputSize
    ).map((point) =>
      mapPhotoEditorOutputPointToSource(point, sourceRect, sourceImage, {
        allowOutside,
      })
    );

    return {
      ...mask,
      shape: 'freehand',
      space: 'source',
      allowOutside,
      rotation: 0,
      rect: getPhotoEditorPointsMaskRect(sourcePoints, { allowOutside }),
      points: sourcePoints,
    };
  }

  const topLeft = mapPhotoEditorOutputPointToSource(rect, sourceRect, sourceImage, {
    allowOutside,
  });
  const bottomRight = mapPhotoEditorOutputPointToSource(
    {
      x: rect.x + rect.width,
      y: rect.y + rect.height,
    },
    sourceRect,
    sourceImage,
    { allowOutside }
  );
  const points = Array.isArray(mask.points)
    ? mask.points.map((point) =>
        mapPhotoEditorOutputPointToSource(point, sourceRect, sourceImage, {
          allowOutside,
        })
      )
    : null;

  return {
    ...mask,
    space: 'source',
    rect: normalizePhotoEditorMaskRect(topLeft, bottomRight, { allowOutside }),
    points,
  };
}

function cancelPhotoEditorPendingMask({ deactivateTool = true } = {}) {
  if (!photoEditorState) {
    return false;
  }

  const hadDraft = Boolean(photoEditorState.draftMask);
  const wasMaskDrag = String(photoEditorState.dragMode || '').startsWith('mask');

  photoEditorState.dragMode = wasMaskDrag ? null : photoEditorState.dragMode;
  photoEditorState.dragStart = wasMaskDrag ? null : photoEditorState.dragStart;
  photoEditorState.dragInitialMask = wasMaskDrag
    ? null
    : photoEditorState.dragInitialMask;
  photoEditorState.draftMask = null;
  photoEditorState.snapGuide = null;
  photoEditorCanvas?.classList.remove('is-panning');

  if (deactivateTool) {
    photoEditorState.maskTool = 'none';
  }

  if (hadDraft) {
    syncPhotoEditorMaskToolUi();
    schedulePhotoEditorRender();
  } else {
    syncPhotoEditorMaskToolUi();
  }

  return hadDraft;
}

function updatePhotoEditorDraftMask(values = {}) {
  if (!photoEditorState?.draftMask) {
    return;
  }

  photoEditorState.draftMask = {
    ...photoEditorState.draftMask,
    ...values,
  };
  schedulePhotoEditorRender({
    debounceMs: PHOTO_EDITOR_EFFECT_DRAG_PREVIEW_DEBOUNCE_MS,
    interactive: true,
  });
}

function cancelPhotoEditorPendingMaskFromControl() {
  if (
    photoEditorState?.draftMask ||
    (photoEditorState && photoEditorState.maskTool !== 'none')
  ) {
    cancelPhotoEditorPendingMask();
  }
}

function getPhotoEditorDraftMaskDragMode(point) {
  if (!photoEditorState?.draftMask || !photoEditorCanvas) {
    return null;
  }

  const sourceImage = photoEditorState.sourceImage;
  const sourceRect = sourceImage ? getPhotoEditorSourceRect(sourceImage) : null;
  const handles = getPhotoEditorDraftMaskHandles(
    photoEditorCanvas.width,
    photoEditorCanvas.height,
    sourceRect,
    sourceImage
  );
  const pointX = clampNumber(point?.x, -10, 10, 0) * photoEditorCanvas.width;
  const pointY = clampNumber(point?.y, -10, 10, 0) * photoEditorCanvas.height;
  const isNearHandle = (handlePoint) =>
    Boolean(
      handlePoint &&
        Math.hypot(pointX - handlePoint.x, pointY - handlePoint.y) <=
          PHOTO_EDITOR_MASK_HANDLE_HIT_RADIUS
    );

  if (isNearHandle(handles?.rotate)) {
    return 'mask-rotate';
  }

  if (isNearHandle(handles?.resize)) {
    return 'mask-resize';
  }

  if (isNearHandle(handles?.move)) {
    return 'mask-move';
  }

  return isPhotoEditorPointInsideMask(
    point,
    photoEditorState.draftMask,
    photoEditorCanvas.width,
    photoEditorCanvas.height,
    sourceRect,
    sourceImage
  )
    ? 'mask-move'
    : null;
}

function beginPhotoEditorDraftMaskTransform(point, dragMode) {
  if (!photoEditorState?.draftMask) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.dragMode = dragMode || 'mask-move';
  photoEditorState.dragStart = point;
  photoEditorState.dragInitialMask = clonePhotoEditorPlainValue(
    photoEditorState.draftMask
  );
  photoEditorCanvas?.classList.add('is-panning');
  paintPhotoEditorPreviewOverlayOnly();
  schedulePhotoEditorRender({ interactive: true });
}

function beginPhotoEditorMaskDrag(event) {
  if (
    !photoEditorState ||
    !photoEditorState.sourceImage
  ) {
    return;
  }

  const point = getPhotoEditorCanvasPointerPoint(event, {
    allowOutside: Boolean(photoEditorState.draftMask),
  });

  if (!point) {
    return;
  }

  const rulerGuideDragMode = getPhotoEditorRulerGuideDragMode(point);

  if (rulerGuideDragMode) {
    event.preventDefault();
    safelySetPointerCapture(photoEditorCanvas, event.pointerId);
    beginPhotoEditorRulerGuideDrag(point, rulerGuideDragMode);
    return;
  }

  event.preventDefault();
  safelySetPointerCapture(photoEditorCanvas, event.pointerId);

  const draftMaskDragMode = getPhotoEditorDraftMaskDragMode(point);

  if (draftMaskDragMode) {
    beginPhotoEditorDraftMaskTransform(point, draftMaskDragMode);
    return;
  }

  if (photoEditorState.maskTool === 'none') {
    const imageOverlayDragInfo = getPhotoEditorImageOverlayDragInfo(point);

    if (imageOverlayDragInfo) {
      beginPhotoEditorImageOverlayDrag(point, imageOverlayDragInfo);
      return;
    }

    const textDragInfo = getPhotoEditorTextDragInfo(point);

    if (textDragInfo) {
      beginPhotoEditorTextDrag(point, textDragInfo);
      return;
    }

    const existingGuide = getPhotoEditorExistingRulerGuideDragInfo(point);

    if (existingGuide) {
      beginPhotoEditorRulerGuideDrag(
        point,
        `guide-${existingGuide.axis}`,
        existingGuide
      );
      return;
    }

    if (
      isPhotoEditorAccordionOpen('imageOverlay') &&
      clearPhotoEditorImageOverlaySelection()
    ) {
      return;
    }

    if (isPhotoEditorAccordionOpen('text') && clearPhotoEditorTextSelection()) {
      return;
    }

    const blurDragMode = getPhotoEditorRadialBlurDragMode(point);

    if (blurDragMode) {
      beginPhotoEditorBlurDrag(point, blurDragMode);
      return;
    }

    beginPhotoEditorPanDrag(point);
    return;
  }

  beginPhotoEditorHistoryMutation();
  const snappedStart = snapPhotoEditorPointToGuides(point, {
    allowOutside: canPhotoEditorMaskExtendOutside(photoEditorState.maskShape),
  }).point;
  photoEditorState.dragMode = 'mask';
  photoEditorState.dragStart = snappedStart;
  photoEditorState.draftMask = {
    type: photoEditorState.maskTool,
    shape: photoEditorState.maskShape,
    space: 'output',
    strength: photoEditorState.blurStrength,
    color: photoEditorState.fillColor,
    rotation: 0,
    allowOutside: canPhotoEditorMaskExtendOutside(photoEditorState.maskShape),
    points: photoEditorState.maskShape === 'freehand' ? [snappedStart] : null,
    rect: {
      x: snappedStart.x,
      y: snappedStart.y,
      width: 0,
      height: 0,
    },
  };
  paintPhotoEditorPreviewOverlayOnly();
  schedulePhotoEditorRender({
    debounceMs: PHOTO_EDITOR_EFFECT_DRAG_PREVIEW_DEBOUNCE_MS,
    interactive: true,
  });
}

function updatePhotoEditorMaskDrag(event) {
  if (!photoEditorState?.dragStart) {
    return;
  }

  const isMaskTransform = ['mask-move', 'mask-resize', 'mask-rotate'].includes(
    String(photoEditorState.dragMode || '')
  );
  const allowOutside =
    isMaskTransform ||
    canPhotoEditorMaskExtendOutside(photoEditorState.draftMask?.shape);
  const point = getPhotoEditorCanvasPointerPoint(event, { allowOutside });

  if (!point) {
    return;
  }

  event.preventDefault();

  if (photoEditorState.dragMode === 'pan') {
    updatePhotoEditorPanDrag(point);
    return;
  }

  if (String(photoEditorState.dragMode || '').startsWith('guide-')) {
    updatePhotoEditorRulerGuideDrag(point);
    return;
  }

  if (String(photoEditorState.dragMode || '').startsWith('text-')) {
    updatePhotoEditorTextDrag(point);
    return;
  }

  if (String(photoEditorState.dragMode || '').startsWith('image-overlay-')) {
    updatePhotoEditorImageOverlayDrag(point, event);
    return;
  }

  if (String(photoEditorState.dragMode || '').startsWith('blur-')) {
    updatePhotoEditorBlurDrag(point);
    return;
  }

  if (
    photoEditorState.dragMode === 'mask-move' ||
    photoEditorState.dragMode === 'mask-resize' ||
    photoEditorState.dragMode === 'mask-rotate'
  ) {
    const initialMask = photoEditorState.dragInitialMask;

    if (!initialMask) {
      return;
    }

    if (photoEditorState.dragMode === 'mask-resize') {
      const snappedPoint = snapPhotoEditorPointToGuides(point, {
        allowOutside,
      }).point;
      photoEditorState.draftMask = resizePhotoEditorMask(initialMask, snappedPoint, {
        keepRatio: Boolean(event.shiftKey),
      });
    } else if (photoEditorState.dragMode === 'mask-rotate') {
      photoEditorState.snapGuide = null;
      photoEditorState.draftMask = rotatePhotoEditorMask(
        initialMask,
        photoEditorState.dragStart,
        point
      );
    } else {
      const translatedMask = translatePhotoEditorMask(
        initialMask,
        point.x - photoEditorState.dragStart.x,
        point.y - photoEditorState.dragStart.y
      );
      photoEditorState.draftMask = snapPhotoEditorMaskToGuides(
        translatedMask
      ).mask;
    }

    paintPhotoEditorPreviewOverlayOnly();
    schedulePhotoEditorRender({
      debounceMs: PHOTO_EDITOR_EFFECT_DRAG_PREVIEW_DEBOUNCE_MS,
      interactive: true,
    });
    return;
  }

  if (!photoEditorState.draftMask) {
    return;
  }

  if (photoEditorState.draftMask.shape === 'freehand') {
    const snappedPoint = snapPhotoEditorPointToGuides(point, {
      allowOutside,
    }).point;
    const points = Array.isArray(photoEditorState.draftMask.points)
      ? photoEditorState.draftMask.points
      : [];
    const lastPoint = points[points.length - 1];
    const shouldAddPoint =
      !lastPoint ||
      Math.hypot(snappedPoint.x - lastPoint.x, snappedPoint.y - lastPoint.y) >
        0.006;

    photoEditorState.draftMask = {
      ...photoEditorState.draftMask,
      points: shouldAddPoint ? [...points, snappedPoint] : points,
      rect: normalizePhotoEditorMaskRect(photoEditorState.dragStart, snappedPoint),
    };
  } else {
    const snappedPoint = snapPhotoEditorPointToGuides(point, {
      allowOutside,
    }).point;
    photoEditorState.draftMask = {
      ...photoEditorState.draftMask,
      rect: normalizePhotoEditorMaskRect(photoEditorState.dragStart, snappedPoint, {
        constrainSquare: Boolean(event.shiftKey),
        allowOutside,
      }),
    };
  }

  paintPhotoEditorPreviewOverlayOnly();
  schedulePhotoEditorRender({
    debounceMs: PHOTO_EDITOR_EFFECT_DRAG_PREVIEW_DEBOUNCE_MS,
    interactive: true,
  });
}

function finishPhotoEditorMaskDrag(event) {
  if (!photoEditorState?.dragStart) {
    return;
  }

  if (photoEditorState.dragMode === 'pan') {
    finishPhotoEditorPanDrag();
    return;
  }

  if (String(photoEditorState.dragMode || '').startsWith('guide-')) {
    finishPhotoEditorRulerGuideDrag();
    return;
  }

  if (String(photoEditorState.dragMode || '').startsWith('text-')) {
    finishPhotoEditorTextDrag();
    return;
  }

  if (String(photoEditorState.dragMode || '').startsWith('image-overlay-')) {
    finishPhotoEditorImageOverlayDrag();
    return;
  }

  if (String(photoEditorState.dragMode || '').startsWith('blur-')) {
    finishPhotoEditorBlurDrag();
    return;
  }

  if (
    photoEditorState.dragMode === 'mask-move' ||
    photoEditorState.dragMode === 'mask-resize' ||
    photoEditorState.dragMode === 'mask-rotate'
  ) {
    photoEditorState.dragStart = null;
    photoEditorState.dragMode = null;
    photoEditorState.dragInitialMask = null;
    photoEditorState.snapGuide = null;
    photoEditorCanvas?.classList.remove('is-panning');
    syncPhotoEditorMaskToolUi();
    finishPhotoEditorInteractivePreview();
    return;
  }

  if (!photoEditorState.draftMask) {
    return;
  }

  const isMaskTransform = ['mask-move', 'mask-resize', 'mask-rotate'].includes(
    String(photoEditorState.dragMode || '')
  );
  const allowOutside =
    isMaskTransform ||
    canPhotoEditorMaskExtendOutside(photoEditorState.draftMask?.shape);
  const point =
    getPhotoEditorCanvasPointerPoint(event, { allowOutside }) ||
    photoEditorState.dragStart;
  const snappedPoint = snapPhotoEditorPointToGuides(point, {
    allowOutside,
  }).point;
  const draftMask = photoEditorState.draftMask;
  const points =
    draftMask.shape === 'freehand'
      ? [
          ...(Array.isArray(draftMask.points) ? draftMask.points : []),
          snappedPoint,
        ]
      : null;
  const rect =
    draftMask.shape === 'freehand'
      ? getPhotoEditorPointsMaskRect(points)
      : normalizePhotoEditorMaskRect(photoEditorState.dragStart, snappedPoint, {
          constrainSquare: Boolean(event.shiftKey),
          allowOutside,
        });
  const shouldSaveMask =
    rect.width >= PHOTO_EDITOR_MIN_MASK_SIZE &&
    rect.height >= PHOTO_EDITOR_MIN_MASK_SIZE;

  const isFreehandMask =
    draftMask.shape === 'freehand' &&
    Array.isArray(points) &&
    points.length > 2 &&
    shouldSaveMask;

  if (!shouldSaveMask) {
    photoEditorState.dragStart = null;
    photoEditorState.dragMode = null;
    photoEditorState.dragInitialMask = null;
    photoEditorState.draftMask = null;
    photoEditorState.snapGuide = null;
    syncPhotoEditorMaskToolUi();
    finishPhotoEditorInteractivePreview();
    return;
  }

  photoEditorState.draftMask = {
    type: draftMask.type,
    shape: isFreehandMask
      ? 'freehand'
      : draftMask.shape === 'ellipse'
        ? 'ellipse'
        : 'rect',
    space: 'output',
    strength: draftMask.strength,
    color: draftMask.color,
    rotation: getPhotoEditorMaskRotation(draftMask),
    allowOutside: canPhotoEditorMaskExtendOutside(draftMask),
    points: isFreehandMask ? points : null,
    rect,
  };

  photoEditorState.dragStart = null;
  photoEditorState.dragMode = null;
  photoEditorState.dragInitialMask = null;
  photoEditorState.snapGuide = null;
  photoEditorCanvas?.classList.remove('is-panning');
  syncPhotoEditorMaskToolUi();
  finishPhotoEditorInteractivePreview();
}

function confirmPhotoEditorMask() {
  if (
    !photoEditorState?.draftMask ||
    String(photoEditorState.dragMode || '').startsWith('mask')
  ) {
    return;
  }

  const committedMask = convertPhotoEditorMaskToSourceSpace(photoEditorState.draftMask);

  if (!committedMask) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.masks.push(committedMask);
  photoEditorState.draftMask = null;
  photoEditorState.maskTool = 'none';
  photoEditorState.dragMode = null;
  photoEditorState.dragStart = null;
  photoEditorState.snapGuide = null;
  syncPhotoEditorMaskToolUi();
  finishPhotoEditorInteractivePreview();
  commitPhotoEditorHistoryMutation();
}

function undoPhotoEditorMask() {
  if (!photoEditorState) {
    return;
  }

  if (photoEditorState.draftMask) {
    cancelPhotoEditorPendingMask();
    return;
  }

  if (photoEditorState.masks.length === 0) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.masks.pop();
  syncPhotoEditorMaskToolUi();
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function clearPhotoEditorMasks() {
  if (!photoEditorState) {
    return;
  }

  beginPhotoEditorHistoryMutation();
  photoEditorState.masks = [];
  photoEditorState.draftMask = null;
  photoEditorState.dragMode = null;
  photoEditorState.dragStart = null;
  photoEditorState.dragInitialMask = null;
  photoEditorState.maskTool = 'none';
  photoEditorState.maskShape = 'rect';
  photoEditorState.maskStrengths = { ...PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS };
  photoEditorState.blurStrength = PHOTO_EDITOR_MASK_DEFAULT_STRENGTHS.blur;
  photoEditorState.fillColor = '#111827';
  syncPhotoEditorMaskToolUi();
  if (photoEditorFillColorInput) {
    photoEditorFillColorInput.value = photoEditorState.fillColor;
  }
  schedulePhotoEditorRender();
  commitPhotoEditorHistoryMutation();
}

function closePhotoEditorModal() {
  if (!photoEditorModal) {
    return;
  }

  photoEditorPreviewRenderToken += 1;
  photoEditorPreviewCommittedValues = null;
  clearPhotoEditorAdjustmentLivePreview();
  cancelPhotoEditorPreviewWorkerRequests('画像編集モーダルを閉じました');

  if (photoEditorRenderFrame) {
    cancelAnimationFrame(photoEditorRenderFrame);
    photoEditorRenderFrame = 0;
  }

  if (photoEditorRenderDebounceTimer) {
    clearTimeout(photoEditorRenderDebounceTimer);
    photoEditorRenderDebounceTimer = 0;
  }

  if (photoEditorPreviewSettleTimer) {
    clearTimeout(photoEditorPreviewSettleTimer);
    photoEditorPreviewSettleTimer = 0;
  }

  const closingState = photoEditorState;

  closeSubModalElement(photoEditorModal, {
    onClosed: () => {
      if (photoEditorState !== closingState) {
        return;
      }

      photoEditorState = null;
      if (photoEditorCanvas) {
        const ctx = photoEditorCanvas.getContext('2d');
        ctx?.clearRect(0, 0, photoEditorCanvas.width, photoEditorCanvas.height);
        photoEditorCanvas.width = 0;
        photoEditorCanvas.height = 0;
        photoEditorCanvas.style.width = '';
        photoEditorCanvas.style.height = '';
        photoEditorCanvas.style.aspectRatio = '';
        photoEditorCanvas.style.filter = '';
        photoEditorCanvas.style.willChange = '';
        photoEditorCanvas.classList.remove(
          'is-blur-tool-active',
          'is-mask-tool-active',
          'is-mask-draft-active',
          'is-panning',
          'is-text-tool-active',
          'is-image-overlay-tool-active'
        );
      }
      if (photoEditorPreviewWorkCanvas) {
        photoEditorPreviewWorkCanvas.width = 0;
        photoEditorPreviewWorkCanvas.height = 0;
      }
      if (photoEditorPreviewOverlayCanvas) {
        photoEditorPreviewOverlayCanvas.width = 0;
        photoEditorPreviewOverlayCanvas.height = 0;
      }
      photoEditorSubjectMaskImageCache.clear();
      photoEditorPreviewOverlayMeta = null;
      setPhotoEditorStatus('');
      if (photoEditorFileName) {
        photoEditorFileName.textContent = '-';
      }
    },
  });
}

function loadPhotoEditorSourceImage(photo, state) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    const loadToken = Symbol('photo-editor-load');
    state.loadToken = loadToken;

    image.onload = () => {
      if (photoEditorState !== state || state.loadToken !== loadToken) {
        reject(new Error('画像読み込みがキャンセルされました'));
        return;
      }

      resolve(image);
    };
    image.onerror = () => {
      reject(new Error('画像を読み込めませんでした'));
    };
    image.decoding = 'async';
    image.src = photo.fileUrl;
  });
}

async function openPhotoEditorModal() {
  if (!currentModalPhoto || !photoEditorModal) {
    return;
  }

  if (!currentModalPhoto.fileUrl) {
    showToast('画像を読み込めませんでした');
    return;
  }

  initializePhotoEditorUi();
  const nextState = createPhotoEditorState(currentModalPhoto);
  photoEditorPreviewCommittedValues = null;
  clearPhotoEditorAdjustmentLivePreview();
  photoEditorState = nextState;

  if (photoEditorFileName) {
    photoEditorFileName.textContent = currentModalPhoto.fileName || '-';
  }

  setPhotoEditorStatus('読み込み中...');
  syncPhotoEditorUi();
  openSubModalElement(photoEditorModal);
  refreshPhotoEditorAiSubjectModelStatus({ silent: true });

  try {
    const image = await loadPhotoEditorSourceImage(currentModalPhoto, nextState);

    if (photoEditorState !== nextState) {
      return;
    }

    photoEditorState.sourceImage = image;
    syncPhotoEditorUi();
    schedulePhotoEditorRender();
  } catch (error) {
    if (photoEditorState === nextState) {
      setPhotoEditorStatus(error.message || '画像を読み込めませんでした');
    }
  }
}

function getPhotoEditorExportFormatMeta(settings = null) {
  const exportSettings = normalizePhotoEditorExportSettings(
    settings || photoEditorState?.exportSettings
  );

  return PHOTO_EDITOR_EXPORT_FORMATS[exportSettings.format] ||
    PHOTO_EDITOR_EXPORT_FORMATS.png;
}

function getPhotoEditorEffectiveExportSettings(settings = null) {
  const exportSettings = normalizePhotoEditorExportSettings(
    settings || photoEditorState?.exportSettings
  );
  const cropPreset = getPhotoEditorCropPreset(photoEditorState?.crop?.preset);

  if (!cropPreset.exportFormat && !cropPreset.exportMaxEdge) {
    return exportSettings;
  }

  return normalizePhotoEditorExportSettings({
    ...exportSettings,
    format: cropPreset.exportFormat || exportSettings.format,
    maxEdge: cropPreset.exportMaxEdge || exportSettings.maxEdge,
  });
}

function getPhotoEditorExportDataUrl(canvas, settings = null) {
  const exportSettings = normalizePhotoEditorExportSettings(
    settings || photoEditorState?.exportSettings
  );
  const formatMeta = getPhotoEditorExportFormatMeta(exportSettings);
  const quality = getPhotoEditorExportQuality(exportSettings, formatMeta);

  return canvas.toDataURL(formatMeta.mimeType, quality);
}

function getPhotoEditorExportQuality(exportSettings, formatMeta) {
  return formatMeta.supportsQuality
    ? clampNumber(exportSettings.quality, 60, 100, 92) / 100
    : undefined;
}

function applyPhotoEditorSubjectTransparencyToCanvas(ctx, outputSize, sourceRect) {
  if (!ctx || !outputSize || !sourceRect) {
    throw new Error('背景透過用の描画を準備できませんでした');
  }

  const maskCanvas = createPhotoEditorSubjectMaskOutputCanvas(outputSize, sourceRect);

  if (!maskCanvas) {
    throw new Error('背景透過に使う被写体マスクを読み込めませんでした');
  }

  ctx.save();
  ctx.globalCompositeOperation = 'destination-in';
  ctx.drawImage(maskCanvas, 0, 0);
  ctx.restore();
}

function buildPhotoEditorExportRenderResult(
  exportCanvas,
  exportSettings,
  renderBase,
  { transparentSubjectBackground = false } = {}
) {
  if (transparentSubjectBackground) {
    applyPhotoEditorSubjectTransparencyToCanvas(
      renderBase.ctx,
      renderBase.outputSize,
      renderBase.sourceRect
    );
  }

  return {
    dataUrl: getPhotoEditorExportDataUrl(exportCanvas, exportSettings),
    outputSize: renderBase.outputSize,
  };
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = '';

  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(
      ...bytes.subarray(index, index + chunkSize)
    );
  }

  return btoa(binary);
}

async function renderPhotoEditorExportDataUrl(
  exportCanvas,
  exportSettings,
  { transparentSubjectBackground = false } = {}
) {
  const fixedOutputSize = getPhotoEditorFixedOutputSize(photoEditorState?.crop);
  const renderBase = drawPhotoEditorBaseToCanvas(exportCanvas, {
    maxEdge: exportSettings.maxEdge > 0 ? exportSettings.maxEdge : null,
    fixedOutputSize,
  });

  if (!renderBase) {
    throw new Error('編集結果を描画できませんでした');
  }

  const formatMeta = getPhotoEditorExportFormatMeta(exportSettings);
  const quality = getPhotoEditorExportQuality(exportSettings, formatMeta);
  const textOverlays = getPhotoEditorTextCollectionFromState().textOverlays;
  const imageOverlays = getPhotoEditorImageOverlayCollectionFromState().imageOverlays;
  const hasTextOverlays = textOverlays.length > 0;
  const hasImageOverlays = imageOverlays.length > 0;
  const hasSubjectMask = hasPhotoEditorReadySubjectMask();
  const needsMainOverlayPass =
    hasTextOverlays || hasImageOverlays || transparentSubjectBackground;

  if (transparentSubjectBackground && !hasSubjectMask) {
    throw new Error('背景透過で保存するにはAI被写体選択のマスクが必要です');
  }

  if (hasImageOverlays) {
    await waitForPhotoEditorImageOverlaysToLoad(imageOverlays);
  }

  if (hasSubjectMask) {
    await waitForPhotoEditorSubjectMaskToLoad(photoEditorState.subjectMask);
  }

  try {
    if (hasSubjectMask) {
      throw new Error('Subject mask composition requires main renderer');
    }

    const workerResult = await applyPhotoEditorEffectsWithWorker(
      exportCanvas,
      renderBase,
      {
        includeDraft: false,
        responseType: needsMainOverlayPass ? 'bitmap' : 'blob',
        exportSettings: {
          mimeType: formatMeta.mimeType,
          quality,
        },
        purpose: 'export',
      }
    );

    if (needsMainOverlayPass && workerResult?.bitmap) {
      renderBase.ctx.clearRect(
        0,
        0,
        renderBase.outputSize.width,
        renderBase.outputSize.height
      );
      renderBase.ctx.drawImage(workerResult.bitmap, 0, 0);
      workerResult.bitmap.close?.();
      applyPhotoEditorImageOverlayToCanvas(
        renderBase.ctx,
        renderBase.outputSize.width,
        renderBase.outputSize.height,
        imageOverlays
      );
      applyPhotoEditorTextOverlayToCanvas(
        renderBase.ctx,
        renderBase.outputSize.width,
        renderBase.outputSize.height,
        textOverlays
      );

      return buildPhotoEditorExportRenderResult(
        exportCanvas,
        exportSettings,
        renderBase,
        { transparentSubjectBackground }
      );
    }

    if (!transparentSubjectBackground && workerResult?.buffer) {
      return {
        dataUrl: `data:${workerResult.mimeType || formatMeta.mimeType};base64,${arrayBufferToBase64(workerResult.buffer)}`,
        outputSize: renderBase.outputSize,
      };
    }
  } catch (error) {
    applyPhotoEditorEffectsToCanvas(
      renderBase.ctx,
      renderBase.outputSize,
      renderBase.sourceRect,
      {
        includeDraft: false,
        drawSubjectMaskPreview: false,
      }
    );

    return buildPhotoEditorExportRenderResult(
      exportCanvas,
      exportSettings,
      renderBase,
      { transparentSubjectBackground }
    );
  }

  applyPhotoEditorEffectsToCanvas(
    renderBase.ctx,
    renderBase.outputSize,
    renderBase.sourceRect,
    {
      includeDraft: false,
      drawSubjectMaskPreview: false,
    }
  );

  return buildPhotoEditorExportRenderResult(
    exportCanvas,
    exportSettings,
    renderBase,
    { transparentSubjectBackground }
  );
}

function buildPhotoEditorSourceMetadataPayload(sourcePhoto) {
  if (!sourcePhoto) {
    return null;
  }

  const labels = [
    ...(Array.isArray(sourcePhoto.photoLabels) ? sourcePhoto.photoLabels : []),
    ...(Array.isArray(currentModalPhotoLabels) && currentModalPhoto?.id === sourcePhoto.id
      ? currentModalPhotoLabels
      : []),
  ];

  return {
    takenAt: sourcePhoto.takenAt,
    takenAtTimestamp: sourcePhoto.takenAtTimestamp,
    groupDate: sourcePhoto.groupDate,
    year: sourcePhoto.year,
    month: sourcePhoto.month,
    day: sourcePhoto.day,
    worldId: sourcePhoto.worldId,
    worldName: sourcePhoto.rawWorldName || sourcePhoto.worldName,
    worldNameManual: sourcePhoto.worldNameManual,
    worldUrl: sourcePhoto.worldUrl,
    isFavorite: Boolean(sourcePhoto.isFavorite),
    printNoteText: sourcePhoto.printNoteText,
    memoText: sourcePhoto.memoText,
    labels: labels.map((label) => label?.name).filter(Boolean),
  };
}

async function savePhotoEditorImage({ transparentSubjectBackground = false } = {}) {
  if (!photoEditorState?.sourceImage || photoEditorState.isSaving) {
    return;
  }

  if (!window.electronAPI.saveEditedPhoto) {
    setPhotoEditorStatus('保存機能を利用できません');
    return;
  }

  if (transparentSubjectBackground && !hasPhotoEditorReadySubjectMask()) {
    setPhotoEditorStatus('背景透過で保存するにはAI被写体選択のマスクが必要です');
    return;
  }

  setPhotoEditorSaving(true);
  setPhotoEditorStatus(
    transparentSubjectBackground ? '背景透過PNGを保存中...' : '保存中...'
  );

  try {
    const exportCanvas = document.createElement('canvas');
    const baseExportSettings = getPhotoEditorEffectiveExportSettings(
      photoEditorState.exportSettings
    );
    const exportSettings = transparentSubjectBackground
      ? normalizePhotoEditorExportSettings({
          ...baseExportSettings,
          format: 'png',
        })
      : baseExportSettings;
    const renderResult = await renderPhotoEditorExportDataUrl(
      exportCanvas,
      exportSettings,
      { transparentSubjectBackground }
    );

    if (!renderResult?.dataUrl) {
      throw new Error('編集結果を描画できませんでした');
    }

    const sourcePhoto = photoEditorState.sourcePhoto;
    const result = await window.electronAPI.saveEditedPhoto({
      sourceFilePath: sourcePhoto.filePath,
      sourceFileName: sourcePhoto.fileName,
      sourceTakenAtTimestamp: sourcePhoto.takenAtTimestamp,
      sourcePhotoMetadata: buildPhotoEditorSourceMetadataPayload(sourcePhoto),
      outputFormat: exportSettings.format,
      outputQuality: exportSettings.quality,
      dataUrl: renderResult.dataUrl,
    });

    if (result?.canceled) {
      setPhotoEditorStatus('保存をキャンセルしました');
      return;
    }

    if (!result?.ok) {
      setPhotoEditorStatus(result?.message || '保存に失敗しました');
      return;
    }

    if (result.importResult && !result.importFailed) {
      await restorePhotoDataSelectionFromResult(result.importResult, currentSelection);
      await queueWorldMetadataSyncForResult(result.importResult);
    }

    closePhotoEditorModal();
    showToast(
      result.importFailed
        ? `保存しました（登録は未反映）: ${result.fileName || ''}`
        : transparentSubjectBackground
          ? `背景透過PNGを保存しました: ${result.fileName || ''}`
          : `編集済み画像を保存しました: ${result.fileName || ''}`
    );
  } catch (error) {
    setPhotoEditorStatus(`保存に失敗しました: ${error.message}`);
  } finally {
    setPhotoEditorSaving(false);
  }
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

  if (photoEditorModal && !photoEditorModal.classList.contains('hidden')) {
    closePhotoEditorModal();
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
  closePhotoEditorModal();

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
  cancelText = 'キャンセル',
  showCancel = true,
} = {}) {
  if (!confirmModal) {
    return Promise.resolve(false);
  }

  confirmModalTitle.textContent = title;
  confirmModalMessage.textContent = message;
  confirmModalConfirmButton.textContent = confirmText;

  if (confirmModalCancelButton) {
    confirmModalCancelButton.textContent = cancelText;
    confirmModalCancelButton.hidden = !showCancel;
  }

  openSubModalElement(confirmModal);

  return new Promise((resolve) => {
    confirmModalResolver = resolve;
  });
}

function getAppUpdateHighlights(payload) {
  const highlights = Array.isArray(payload?.highlights)
    ? payload.highlights
        .map((item) => (typeof item === 'string' ? item.trim() : ''))
        .filter(Boolean)
        .slice(0, 4)
    : [];

  return highlights;
}

function buildAppUpdatePromptConfig(payload) {
  const kind =
    typeof payload?.kind === 'string' ? payload.kind.trim().toLowerCase() : '';
  const version =
    typeof payload?.version === 'string' ? payload.version.trim() : '';
  const versionLabel = version || '最新バージョン';

  if (kind === 'downloaded') {
    return {
      title: translateUiText('アップデートの準備ができました'),
      message:
        `${versionLabel} のダウンロードが完了しました。再起動して更新しますか？`,
      confirmText: '再起動して更新',
    };
  }

  if (kind === 'installed') {
    const highlights = getAppUpdateHighlights(payload);

    return {
      title: translateUiText('主な更新内容'),
      message:
        highlights.length > 0
          ? highlights.map((item) => `・${item}`).join('\n')
          : `${versionLabel} へのアップデートが完了しました。`,
      confirmText: 'OK',
      showCancel: false,
    };
  }

  if (kind === 'available') {
    return {
      title: translateUiText('アップデートがあります'),
      message:
        `新しいバージョン ${versionLabel} が利用できます。今すぐダウンロードしますか？`,
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

function setSettingsDataStatus(message = '', tone = 'default') {
  if (!settingsDataStatus) {
    return;
  }

  settingsDataStatus.textContent = message;
  settingsDataStatus.classList.remove('is-success', 'is-error', 'is-busy');

  if (tone === 'success') {
    settingsDataStatus.classList.add('is-success');
  } else if (tone === 'error') {
    settingsDataStatus.classList.add('is-error');
  } else if (tone === 'busy') {
    settingsDataStatus.classList.add('is-busy');
  }
}

function setSettingsAiModelStatus(message = '', tone = 'default') {
  if (!settingsAiModelStatus) {
    return;
  }

  settingsAiModelStatus.textContent = message;
  settingsAiModelStatus.classList.remove('is-success', 'is-error', 'is-busy');

  if (tone === 'success') {
    settingsAiModelStatus.classList.add('is-success');
  } else if (tone === 'error') {
    settingsAiModelStatus.classList.add('is-error');
  } else if (tone === 'busy') {
    settingsAiModelStatus.classList.add('is-busy');
  }
}

function formatAiSubjectModelSize(sizeBytes, fallback = '') {
  const size = Number(sizeBytes) || 0;

  if (size <= 0) {
    return formatAiSubjectModelSizeLabel(fallback);
  }

  if (size >= 1024 * 1024) {
    return `${Math.round((size / 1024 / 1024) * 10) / 10}MB`;
  }

  return `${Math.round(size / 1024)}KB`;
}

function formatAiSubjectModelSizeLabel(sizeLabel = '') {
  const label = String(sizeLabel || '');

  if (!label) {
    return '';
  }

  const approximateMatch = label.match(/^約(.+)$/);

  if (approximateMatch) {
    const language = window.WorldShotI18n?.getLanguage?.() || 'ja';

    if (language === 'en') {
      return `approx. ${approximateMatch[1]}`;
    }

    if (language === 'ko') {
      return `약 ${approximateMatch[1]}`;
    }
  }

  return translateUiText(label);
}

function renderSettingsAiSubjectModels() {
  if (!settingsAiModelList) {
    return;
  }

  const allModels = Array.isArray(photoEditorAiSubjectModelStatus?.models)
    ? photoEditorAiSubjectModelStatus.models
    : [];
  const models = allModels.filter((model) => {
    if (model.deprecated) {
      return model.installed;
    }

    return model.visible;
  });
  settingsAiModelList.innerHTML = '';

  if (models.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'settings-section-meta';
    empty.textContent = 'AIモデル情報を取得できませんでした';
    settingsAiModelList.appendChild(empty);
    return;
  }

  for (const model of models) {
    const item = document.createElement('div');
    item.className = 'settings-ai-model-item';
    item.dataset.aiSubjectModelTier = model.tier || '';

    const body = document.createElement('div');
    const title = document.createElement('p');
    title.className = 'settings-ai-model-name';
    const modelDisplayName = translateUiText(model.displayName || model.label);
    const modelLabel = model.label || model.modelName;
    title.textContent = model.deprecated
      ? `${translateUiText('旧モデル')}: ${modelLabel}`
      : `${modelDisplayName} / ${modelLabel}`;

    const meta = document.createElement('p');
    meta.className = 'settings-ai-model-meta';
    const statusLabel = model.ready
      ? '利用可能'
      : model.installed && !model.verified
        ? '検証失敗'
        : model.licenseRestricted
          ? '利用停止中'
        : model.future
          ? '準備中'
        : model.partial
          ? '一部不足'
          : model.installed
            ? '利用停止中'
            : model.bundled
              ? '同梱'
              : '未ダウンロード';
    const sizeText = formatAiSubjectModelSize(
      model.sizeBytes,
      model.sizeLabel || ''
    );
    meta.textContent = [
      translateUiText(statusLabel),
      sizeText,
      model.license,
      translateUiText(model.bundled ? '同梱' : '管理フォルダ'),
      translateUiText(model.description),
    ]
      .filter(Boolean)
      .join(' / ');

    body.appendChild(title);
    body.appendChild(meta);

    const actions = document.createElement('div');
    actions.className = 'settings-ai-model-actions';

    if (model.canDownload && !model.ready) {
      const downloadButton = document.createElement('button');
      downloadButton.type = 'button';
      downloadButton.className = 'secondary';
      downloadButton.dataset.aiSubjectModelDownload = model.id;
      downloadButton.textContent = translateUiText('ダウンロード');
      actions.appendChild(downloadButton);
    }

    if (!model.bundled && (!model.future || model.installed)) {
      const folderButton = document.createElement('button');
      folderButton.type = 'button';
      folderButton.className = 'secondary';
      folderButton.dataset.aiSubjectModelFolder = model.id;
      folderButton.textContent = translateUiText('保存場所');
      actions.appendChild(folderButton);
    }

    if (model.canDelete) {
      const deleteButton = document.createElement('button');
      deleteButton.type = 'button';
      deleteButton.className = 'danger-button';
      deleteButton.dataset.aiSubjectModelDelete = model.id;
      deleteButton.textContent = translateUiText('削除');
      actions.appendChild(deleteButton);
    }

    if (!actions.children.length) {
      const badge = document.createElement('span');
      badge.className = 'settings-ai-model-meta';
      badge.textContent = model.bundled ? translateUiText('同梱モデル') : '-';
      actions.appendChild(badge);
    }

    item.appendChild(body);
    item.appendChild(actions);
    settingsAiModelList.appendChild(item);
  }
}

async function refreshSettingsAiSubjectModelStatus({ silent = false } = {}) {
  const result = await refreshPhotoEditorAiSubjectModelStatus({ silent: true });

  if (!result?.ok && !silent) {
    setSettingsAiModelStatus(
      result?.message || 'AIモデル情報を取得できませんでした',
      'error'
    );
  }

  renderSettingsAiSubjectModels();
  return result;
}

async function downloadSettingsAiSubjectModel(modelId) {
  const model = getPhotoEditorSubjectModel(modelId);

  if (!window.electronAPI.downloadAiSubjectModel) {
    setSettingsAiModelStatus('AIモデルのダウンロード機能を利用できません', 'error');
    return;
  }

  if (!model?.canDownload) {
    setSettingsAiModelStatus('このAIモデルはダウンロードできません', 'error');
    return;
  }

  const confirmed = await openConfirmModal({
    title: `${model.displayName || model.label}をダウンロード`,
    message:
      `モデル: ${model.label || model.modelName}\n` +
      `容量: ${model.sizeLabel || formatAiSubjectModelSize(model.sizeBytes)}\n` +
      `ライセンス: ${model.license}\n` +
      `${model.description || ''}\n\n` +
      (model.id === 'withoutbg-focus'
        ? 'このモデルは標準AIより処理時間やメモリ使用量が大きくなる場合があります。\n'
        : '') +
      '画像は外部サーバーに送信されません。AI処理はPC内で実行されます。',
    confirmText: 'ダウンロード',
  });

  if (!confirmed) {
    return;
  }

  setSettingsAiModelStatus(
    `${model.displayName || model.label}をダウンロードしています...`,
    'busy'
  );

  try {
    const result = await window.electronAPI.downloadAiSubjectModel(modelId);

    if (!result?.ok) {
      throw new Error(result?.message || 'AIモデルをダウンロードできませんでした');
    }

    photoEditorAiSubjectModelStatus = result;
    renderSettingsAiSubjectModels();
    syncPhotoEditorSubjectMaskControls();
    setSettingsAiModelStatus(
      `${model.displayName || model.label}をダウンロードしました`,
      'success'
    );
    showToast(`${model.displayName || model.label}をダウンロードしました`);
  } catch (error) {
    setSettingsAiModelStatus(error.message, 'error');
    showToast(`AIモデルのダウンロードに失敗しました: ${error.message}`);
    await refreshSettingsAiSubjectModelStatus({ silent: true });
  }
}

async function openSettingsAiSubjectModelFolder(modelId) {
  if (!window.electronAPI.openAiSubjectModelFolder) {
    setSettingsAiModelStatus('保存場所を開けません', 'error');
    return;
  }

  const result = await window.electronAPI.openAiSubjectModelFolder(modelId);

  if (!result?.ok) {
    setSettingsAiModelStatus(
      result?.message || '保存場所を開けませんでした',
      'error'
    );
  }
}

async function deleteSettingsAiSubjectModel(modelId) {
  const model = getPhotoEditorSubjectModel(modelId);
  const confirmed = await openConfirmModal({
    title: 'AIモデルを削除',
    message: `${model?.label || 'AIモデル'}を管理フォルダから削除します。必要になったら設定から再ダウンロードできます。続行しますか？`,
    confirmText: '削除する',
  });

  if (!confirmed) {
    return;
  }

  if (!window.electronAPI.deleteAiSubjectModel) {
    setSettingsAiModelStatus('AIモデルの削除機能を利用できません', 'error');
    return;
  }

  setSettingsAiModelStatus('AIモデルを削除しています...', 'busy');

  try {
    const result = await window.electronAPI.deleteAiSubjectModel(modelId);

    if (!result?.ok) {
      throw new Error(result?.message || 'AIモデルを削除できませんでした');
    }

    for (const [sessionKey, session] of photoEditorSubjectModelSessionCache) {
      if (sessionKey.startsWith(`${modelId}:`)) {
        session.release?.();
        photoEditorSubjectModelSessionCache.delete(sessionKey);
      }
    }

    photoEditorAiSubjectModelStatus = result;
    renderSettingsAiSubjectModels();
    syncPhotoEditorSubjectMaskControls();
    setSettingsAiModelStatus('AIモデルを削除しました', 'success');
    showToast('AIモデルを削除しました');
  } catch (error) {
    setSettingsAiModelStatus(error.message, 'error');
    showToast(`AIモデルの削除に失敗しました: ${error.message}`);
    await refreshSettingsAiSubjectModelStatus({ silent: true });
  }
}

function setSettingsDataSectionOpen(isOpen) {
  isSettingsDataSectionOpen = Boolean(isOpen);
  settingsDataSection?.classList.toggle('is-open', isSettingsDataSectionOpen);
  settingsDataToggleButton?.setAttribute(
    'aria-expanded',
    String(isSettingsDataSectionOpen)
  );
  settingsDataPanel?.setAttribute(
    'aria-hidden',
    String(!isSettingsDataSectionOpen)
  );
}

function syncSettingsDataUi() {
  const buttons = [
    createAppDataBackupButton,
    checkAppDataHealthButton,
    showMissingOriginalFilesButton,
    showMissingThumbnailsButton,
    showMissingWorldInfoButton,
    showWorldMetadataIssuesButton,
    regenerateMissingThumbnailsButton,
    refreshWorldMetadataIssuesButton,
    restoreAppDataBackupButton,
    exportPhotoCatalogCsvButton,
    exportPhotoCatalogJsonButton,
  ].filter(Boolean);
  const disabled = Boolean(isImporting);

  for (const button of buttons) {
    button.disabled = disabled;
  }

  settingsDataToggleButton?.setAttribute(
    'title',
    isSettingsDataSectionOpen ? 'データ管理を閉じる' : 'データ管理を開く'
  );
  settingsDataToggleButton?.setAttribute(
    'aria-label',
    isSettingsDataSectionOpen ? 'データ管理を閉じる' : 'データ管理を開く'
  );
  createAppDataBackupButton?.setAttribute(
    'title',
    disabled ? '処理中はバックアップできません' : 'アプリデータをJSONでバックアップ'
  );
  checkAppDataHealthButton?.setAttribute(
    'title',
    disabled ? '処理中は状態チェックできません' : '登録データの状態をチェック'
  );
  showMissingOriginalFilesButton?.setAttribute(
    'title',
    disabled ? '処理中は抽出できません' : '元画像なしの写真だけを表示'
  );
  showMissingThumbnailsButton?.setAttribute(
    'title',
    disabled ? '処理中は抽出できません' : 'サムネイルなしの写真だけを表示'
  );
  showMissingWorldInfoButton?.setAttribute(
    'title',
    disabled ? '処理中は抽出できません' : 'World情報未取得の写真だけを表示'
  );
  showWorldMetadataIssuesButton?.setAttribute(
    'title',
    disabled
      ? '処理中は抽出できません'
      : 'Worldメタデータ要確認の写真だけを表示'
  );
  regenerateMissingThumbnailsButton?.setAttribute(
    'title',
    disabled
      ? '処理中は再生成できません'
      : '欠損しているサムネイルだけを再生成'
  );
  refreshWorldMetadataIssuesButton?.setAttribute(
    'title',
    disabled
      ? '処理中は再取得できません'
      : 'Worldメタデータ要確認の該当分だけを再取得'
  );
  restoreAppDataBackupButton?.setAttribute(
    'title',
    disabled ? '処理中は復元できません' : 'バックアップJSONから復元'
  );
  exportPhotoCatalogCsvButton?.setAttribute(
    'title',
    disabled ? '処理中はエクスポートできません' : '写真一覧をCSVで書き出し'
  );
  exportPhotoCatalogJsonButton?.setAttribute(
    'title',
    disabled ? '処理中はエクスポートできません' : '写真一覧をJSONで書き出し'
  );
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
    photoEditorClose,
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

  if (closeButton === photoEditorClose) {
    event.preventDefault();
    event.stopPropagation();
    closePhotoEditorModal();
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

  if (modalEditPhotoButton) {
    modalEditPhotoButton.textContent = '画像を加工する';
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

function initializePhotoEditorUi() {
  initializePhotoEditorAccordions();
  photoEditorUserPresets = loadPhotoEditorUserPresets();
  photoEditorTextRecentFonts = loadPhotoEditorRecentTextFonts();
  renderPhotoEditorTextFontOptions();
  renderPhotoEditorPresetButtons();
  void refreshPhotoEditorOverlayAssets({ silent: true });

  if (photoEditorCropPresetList && !photoEditorCropPresetList.dataset.initialized) {
    photoEditorCropPresetList.innerHTML = '';

    for (const preset of PHOTO_EDITOR_CROP_PRESETS) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'photo-editor-crop-button';
      button.dataset.photoEditorCrop = preset.key;
      button.setAttribute('aria-pressed', 'false');
      button.textContent = preset.label;
      if (preset.transparentPadding) {
        button.classList.add('is-output-preset');
      }
      button.addEventListener('click', () => {
        applyPhotoEditorCropPreset(preset.key);
      });
      photoEditorCropPresetList.appendChild(button);
    }

    photoEditorCropPresetList.dataset.initialized = 'true';
  }

  if (
    photoEditorCurveChannelList &&
    !photoEditorCurveChannelList.dataset.initialized
  ) {
    photoEditorCurveChannelList.innerHTML = '';

    for (let index = 0; index < 4; index += 1) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'photo-editor-curve-channel-button';
      button.dataset.photoEditorCurveChannel = '';
      button.setAttribute('aria-pressed', 'false');
      button.addEventListener('click', () => {
        setPhotoEditorCurveChannel(button.dataset.photoEditorCurveChannel);
      });
      photoEditorCurveChannelList.appendChild(button);
    }

    photoEditorCurveChannelList.dataset.initialized = 'true';
  }

  if (photoEditorAdjustmentList && !photoEditorAdjustmentList.dataset.initialized) {
    photoEditorAdjustmentList.innerHTML = '';

    for (const slider of PHOTO_EDIT_SLIDERS) {
      const row = document.createElement('div');
      row.className = 'photo-editor-slider-row';

      const label = document.createElement('div');
      label.className = 'photo-editor-slider-label';

      const labelText = document.createElement('span');
      labelText.textContent = slider.label;
      label.appendChild(labelText);

      const valueText = document.createElement('span');
      valueText.dataset.photoEditorValue = slider.key;
      valueText.textContent = String(slider.defaultValue);
      label.appendChild(valueText);

      const control = document.createElement('div');
      control.className = 'photo-editor-slider-control';

      const input = document.createElement('input');
      input.type = 'range';
      input.min = String(slider.min);
      input.max = String(slider.max);
      input.value = String(slider.defaultValue);
      input.dataset.photoEditorSlider = slider.key;
      input.addEventListener('input', () => {
        updatePhotoEditorSliderValue(slider.key, input.value);
      });
      input.addEventListener('change', () => {
        finishPhotoEditorInteractivePreview();
      });
      control.appendChild(input);

      const resetButton = document.createElement('button');
      resetButton.type = 'button';
      resetButton.className = 'photo-editor-slider-reset';
      resetButton.setAttribute('aria-label', `${slider.label}をリセット`);
      resetButton.title = `${slider.label}をリセット`;
      resetButton.innerHTML =
        '<span class="material-symbols-outlined">restart_alt</span>';
      resetButton.addEventListener('click', () => {
        updatePhotoEditorSliderValue(slider.key, slider.defaultValue);
      });
      control.appendChild(resetButton);

      row.appendChild(label);
      row.appendChild(control);
      photoEditorAdjustmentList.appendChild(row);
    }

    photoEditorAdjustmentList.dataset.initialized = 'true';
  }
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
  syncSettingsDataUi();
  syncSettingsBackgroundUi();
  syncSelectionDependentSettingsUi();
  await refreshSettingsAiSubjectModelStatus({ silent: true });

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
  syncSettingsDataUi();
  syncSettingsUtilityActionsUi();
  syncSettingsMaintenanceUi();
  syncSettingsUninstallUi();
}

function resetCurrentMonthState() {
  setCurrentSelectionValue(null);
  currentPhotos = [];
  allCurrentMonthPhotos = [];
  currentPhotoGroupIndexMap = new Map();
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

function createThumbnailPlaceholder(message = 'サムネイル未生成') {
  const placeholder = document.createElement('div');
  placeholder.className = 'photo-card-image photo-card-image-placeholder';
  placeholder.draggable = false;

  const icon = document.createElement('span');
  icon.className = 'material-symbols-outlined photo-card-placeholder-icon';
  icon.textContent = 'image';

  const text = document.createElement('span');
  text.className = 'photo-card-placeholder-text';
  text.textContent = message;

  placeholder.appendChild(icon);
  placeholder.appendChild(text);

  return placeholder;
}

function createPhotoCard(item, photoIndex = 0) {
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
    image.className = 'photo-card-image is-loading';
    image.alt = item.fileName || 'photo';
    image.loading =
      photoIndex < GALLERY_EAGER_THUMBNAIL_COUNT ? 'eager' : 'lazy';
    image.decoding = 'async';
    image.fetchPriority =
      photoIndex < GALLERY_EAGER_THUMBNAIL_COUNT ? 'high' : 'low';
    image.draggable = false;
    image.addEventListener(
      'load',
      () => {
        image.classList.remove('is-loading');
        image.classList.add('is-loaded');
      },
      { once: true }
    );
    image.addEventListener(
      'error',
      () => {
        const fallback = createThumbnailPlaceholder('サムネイル要再生成');
        card.classList.add('has-thumbnail-error');
        image.replaceWith(fallback);
      },
      { once: true }
    );
    image.src = item.thumbnailUrl;
    visual = image;
  } else {
    visual = createThumbnailPlaceholder();
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
  timeInline.textContent = timeText || '時刻不明';
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

function clearPendingGalleryDateJump() {
  galleryDateJumpRequestId += 1;

  if (galleryDateJumpRenderFrame) {
    cancelAnimationFrame(galleryDateJumpRenderFrame);
    galleryDateJumpRenderFrame = 0;
  }

  if (galleryDateJumpTimer) {
    clearTimeout(galleryDateJumpTimer);
    galleryDateJumpTimer = 0;
  }

  if (galleryJumpAnimationTimer) {
    clearTimeout(galleryJumpAnimationTimer);
    galleryJumpAnimationTimer = 0;
  }

  if (activeGalleryJumpTarget) {
    activeGalleryJumpTarget.classList.remove('is-jump-target');
    activeGalleryJumpTarget = null;
  }
}

function startGalleryDateJumpRequest() {
  galleryDateJumpRequestId += 1;

  if (galleryDateJumpRenderFrame) {
    cancelAnimationFrame(galleryDateJumpRenderFrame);
    galleryDateJumpRenderFrame = 0;
  }

  if (galleryDateJumpTimer) {
    clearTimeout(galleryDateJumpTimer);
    galleryDateJumpTimer = 0;
  }

  if (monthGalleryAppendFrame) {
    cancelAnimationFrame(monthGalleryAppendFrame);
    monthGalleryAppendFrame = 0;
  }

  return galleryDateJumpRequestId;
}

function cancelMonthGalleryScheduledWork() {
  if (monthGalleryLoadCheckTimer) {
    clearTimeout(monthGalleryLoadCheckTimer);
    monthGalleryLoadCheckTimer = null;
  }

  clearPendingGalleryDateJump();

  if (activeGalleryDateSyncFrame) {
    cancelAnimationFrame(activeGalleryDateSyncFrame);
    activeGalleryDateSyncFrame = 0;
  }

  if (monthGalleryAppendFrame) {
    cancelAnimationFrame(monthGalleryAppendFrame);
    monthGalleryAppendFrame = 0;
  }

  isAppendingMonthGalleryBatch = false;
  monthGalleryLoadCheckScheduled = false;
}

function resetMonthGalleryRenderState() {
  cancelMonthGalleryScheduledWork();

  renderedPhotoCount = 0;
  renderedMonthGalleryKey = '';
  renderedGalleryGroupMap = new Map();
  renderedGalleryGroupList = [];
  activeGalleryGroupIndex = -1;
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
  const visibleHeight = Math.max(
    monthGalleryList.clientHeight || window.innerHeight - Math.max(rect.top, 0),
    320
  );
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

function buildGalleryGroupSection(groupDate, startIndex = -1) {
  const section = document.createElement('section');
  section.className = 'gallery-group';
  section.dataset.groupDate = groupDate;

  const dayBox = document.createElement('div');
  dayBox.className = 'day-box';
  dayBox.textContent = groupDate;

  const grid = document.createElement('div');
  grid.className = 'gallery-grid';

  section.appendChild(dayBox);
  section.appendChild(grid);

  return {
    section,
    grid,
    groupDate,
    startIndex,
    endIndex: startIndex >= 0 ? startIndex + 1 : startIndex,
  };
}

function getActiveGalleryGroupDateFromViewport() {
  if (!monthGalleryList || !isMonthSelection(currentSelection)) {
    return '';
  }

  if (renderedGalleryGroupList.length === 0) {
    return '';
  }

  const containerRect = monthGalleryList.getBoundingClientRect();
  const visibleTop = Math.max(containerRect.top, 0);
  const visibleBottom = Math.min(containerRect.bottom, window.innerHeight);
  const visibleHeight = Math.max(0, visibleBottom - visibleTop);
  const anchorY = visibleTop + Math.min(48, Math.max(visibleHeight, 1) * 0.18);
  let bestGroup = null;
  let bestIndex = -1;
  let bestDistance = Number.POSITIVE_INFINITY;

  for (let index = 0; index < renderedGalleryGroupList.length; index += 1) {
    const group = renderedGalleryGroupList[index]?.section;

    if (!group || !group.isConnected || group.hidden) {
      continue;
    }

    const rect = group.getBoundingClientRect();

    if (rect.bottom < visibleTop) {
      continue;
    }

    if (rect.top > visibleBottom) {
      break;
    }

    const distance =
      rect.top <= anchorY && rect.bottom >= anchorY
        ? 0
        : Math.abs(rect.top - anchorY);

    if (distance < bestDistance) {
      bestDistance = distance;
      bestGroup = group;
      bestIndex = index;
    }
  }

  activeGalleryGroupIndex = bestIndex;

  return bestGroup?.dataset?.groupDate || '';
}

function syncActiveGalleryGroupDate() {
  activeGalleryDateSyncFrame = 0;
  setActiveSidebarGroupDate(getActiveGalleryGroupDateFromViewport());
}

function scheduleActiveGalleryDateSync() {
  if (activeGalleryDateSyncFrame) {
    return;
  }

  activeGalleryDateSyncFrame = requestAnimationFrame(syncActiveGalleryGroupDate);
}

function getGalleryScrollTarget() {
  const galleryOverflowY = monthGalleryList
    ? window.getComputedStyle(monthGalleryList).overflowY
    : '';

  if (
    monthGalleryList &&
    galleryOverflowY !== 'visible'
  ) {
    return {
      element: monthGalleryList,
      getCurrent: () => monthGalleryList.scrollTop,
      setCurrent: (value) => {
        monthGalleryList.scrollTop = value;
      },
      scrollTo(top, behavior) {
        monthGalleryList.scrollTo({ top, behavior });
      },
      getTargetTop(targetSection) {
        const containerRect = monthGalleryList.getBoundingClientRect();
        const targetRect = targetSection.getBoundingClientRect();
        return (
          monthGalleryList.scrollTop +
          targetRect.top -
          containerRect.top -
          18
        );
      },
    };
  }

  if (appRoot && appRoot.scrollHeight > appRoot.clientHeight + 1) {
    return {
      element: appRoot,
      getCurrent: () => appRoot.scrollTop,
      setCurrent: (value) => {
        appRoot.scrollTop = value;
      },
      scrollTo(top, behavior) {
        appRoot.scrollTo({ top, behavior });
      },
      getTargetTop(targetSection) {
        const containerRect = appRoot.getBoundingClientRect();
        const targetRect = targetSection.getBoundingClientRect();
        return appRoot.scrollTop + targetRect.top - containerRect.top - 18;
      },
    };
  }

  return {
    element: window,
    getCurrent: () =>
      window.scrollY ||
      document.documentElement.scrollTop ||
      document.body.scrollTop ||
      0,
    setCurrent: (value) => {
      window.scrollTo(0, value);
    },
    scrollTo(top, behavior) {
      window.scrollTo({ top, behavior });
    },
    getTargetTop(targetSection) {
      return (
        window.scrollY +
        targetSection.getBoundingClientRect().top -
        18
      );
    },
  };
}

function jumpGalleryToTarget(scrollTarget, targetTop) {
  if (!scrollTarget) {
    return;
  }

  const normalizedTargetTop = Math.max(0, targetTop);
  scrollTarget.setCurrent(normalizedTargetTop);
  scheduleActiveGalleryDateSync();
}

function scrollGallerySectionIntoView(targetSection) {
  if (!targetSection || !monthGalleryList) {
    return;
  }

  const scrollTarget = getGalleryScrollTarget();
  jumpGalleryToTarget(scrollTarget, scrollTarget.getTargetTop(targetSection));
}

function playGalleryGroupJumpAnimation(targetSection) {
  if (!targetSection) {
    return;
  }

  if (galleryJumpAnimationTimer) {
    clearTimeout(galleryJumpAnimationTimer);
    galleryJumpAnimationTimer = 0;
  }

  if (activeGalleryJumpTarget && activeGalleryJumpTarget !== targetSection) {
    activeGalleryJumpTarget.classList.remove('is-jump-target');
  }

  activeGalleryJumpTarget = targetSection;
  targetSection.classList.remove('is-jump-target');
  // Force a reflow so repeating a nearby date jump restarts the pulse.
  void targetSection.offsetWidth;
  targetSection.classList.add('is-jump-target');

  galleryJumpAnimationTimer = setTimeout(() => {
    targetSection.classList.remove('is-jump-target');
    if (activeGalleryJumpTarget === targetSection) {
      activeGalleryJumpTarget = null;
    }
    galleryJumpAnimationTimer = 0;
    scheduleActiveGalleryDateSync();
  }, 1200);
}

function scheduleGalleryDateJump(targetSection) {
  if (!targetSection) {
    return;
  }

  if (galleryDateJumpTimer) {
    clearTimeout(galleryDateJumpTimer);
    galleryDateJumpTimer = 0;
  }

  // Wait briefly for the appended card batch to enter layout, then jump once.
  // Repeated post-jump corrections read as stutter when users switch dates.
  galleryDateJumpTimer = setTimeout(() => {
    galleryDateJumpTimer = 0;
    scrollGallerySectionIntoView(targetSection);
    setKeyboardFocusToFirstPhotoInGroup(targetSection);
    playGalleryGroupJumpAnimation(targetSection);
  }, 32);
}

function scheduleGalleryDateJumpAfterRender({
  groupDate,
  readyRenderCount,
  backgroundRenderCount,
  requestId,
}) {
  if (galleryDateJumpRenderFrame) {
    cancelAnimationFrame(galleryDateJumpRenderFrame);
    galleryDateJumpRenderFrame = 0;
  }

  const waitForTargetSection = () => {
    galleryDateJumpRenderFrame = 0;

    if (requestId !== galleryDateJumpRequestId) {
      return;
    }

    const targetSection = renderedGalleryGroupMap.get(groupDate)?.section;

    if (targetSection) {
      setActiveSidebarGroupDate(groupDate);
      scheduleGalleryDateJump(targetSection);

      if (backgroundRenderCount > renderedPhotoCount) {
        appendMonthGalleryPhotoBatch(backgroundRenderCount);
      }

      return;
    }

    if (readyRenderCount > renderedPhotoCount && !monthGalleryAppendFrame) {
      appendMonthGalleryPhotoBatch(readyRenderCount);
    }

    if (readyRenderCount <= renderedPhotoCount && !monthGalleryAppendFrame) {
      return;
    }

    galleryDateJumpRenderFrame = requestAnimationFrame(waitForTargetSection);
  };

  galleryDateJumpRenderFrame = requestAnimationFrame(waitForTargetSection);
}

function scrollMonthGalleryToGroupDate(groupDate) {
  const normalizedGroupDate =
    typeof groupDate === 'string' ? groupDate.trim() : '';

  if (!normalizedGroupDate || currentPhotos.length === 0) {
    return false;
  }

  const groupIndex = currentPhotoGroupIndexMap.get(normalizedGroupDate);

  if (!groupIndex) {
    showToast('現在の絞り込みでは、この日の写真は表示されていません');
    return false;
  }

  const readyRenderCount = Math.min(currentPhotos.length, groupIndex.startIndex + 1);
  const targetRenderCount = Math.min(
    currentPhotos.length,
    groupIndex.endIndex + calculateInitialVisiblePhotoCount()
  );
  const jumpRequestId = startGalleryDateJumpRequest();

  if (readyRenderCount > renderedPhotoCount) {
    appendMonthGalleryPhotoBatch(readyRenderCount);
    scheduleGalleryDateJumpAfterRender({
      groupDate: normalizedGroupDate,
      readyRenderCount,
      backgroundRenderCount: targetRenderCount,
      requestId: jumpRequestId,
    });
    return true;
  }

  const groupState = renderedGalleryGroupMap.get(normalizedGroupDate);
  const targetSection = groupState?.section;

  if (!targetSection) {
    return true;
  }

  setActiveSidebarGroupDate(normalizedGroupDate);
  scheduleGalleryDateJump(targetSection);

  if (targetRenderCount > renderedPhotoCount) {
    appendMonthGalleryPhotoBatch(targetRenderCount);
  }

  return true;
}

function appendMonthGalleryPhotoBatch(targetCount) {
  if (!monthGalleryList || targetCount <= renderedPhotoCount) {
    return;
  }

  const normalizedTargetCount = Math.min(targetCount, currentPhotos.length);
  const nextTargetCount = Math.min(
    normalizedTargetCount,
    renderedPhotoCount + GALLERY_MAX_CARDS_PER_APPEND
  );
  const fragment = document.createDocumentFragment();

  for (let index = renderedPhotoCount; index < nextTargetCount; index += 1) {
    const photo = currentPhotos[index];
    const groupDate = getPhotoGroupDate(photo);

    let groupState = renderedGalleryGroupMap.get(groupDate);

    if (!groupState) {
      groupState = buildGalleryGroupSection(groupDate, index);
      renderedGalleryGroupMap.set(groupDate, groupState);
      renderedGalleryGroupList.push(groupState);
      fragment.appendChild(groupState.section);
    } else {
      groupState.endIndex = index + 1;
    }

    groupState.grid.appendChild(createPhotoCard(photo, index));
  }

  renderedPhotoCount = nextTargetCount;
  monthGalleryList.appendChild(fragment);
  syncSelectionUi();
  syncKeyboardFocusedPhotoCard();
  scheduleActiveGalleryDateSync();

  if (renderedPhotoCount < normalizedTargetCount && !monthGalleryAppendFrame) {
    monthGalleryAppendFrame = requestAnimationFrame(() => {
      monthGalleryAppendFrame = 0;
      appendMonthGalleryPhotoBatch(normalizedTargetCount);
      scheduleMonthGalleryLoadCheck({ immediate: true });
    });
  }
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

  const remainingScroll =
    monthGalleryList.scrollHeight -
    monthGalleryList.scrollTop -
    monthGalleryList.clientHeight;

  if (remainingScroll > GALLERY_LOAD_AHEAD_PX) {
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
    monthGalleryList.scrollHeight -
      monthGalleryList.scrollTop -
      monthGalleryList.clientHeight <=
      GALLERY_LOAD_AHEAD_PX
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
  monthGalleryList?.addEventListener('scroll', () => {
    scheduleMonthGalleryLoadCheck();
    scheduleActiveGalleryDateSync();
  }, { passive: true });
  window.addEventListener('scroll', () => {
    scheduleMonthGalleryLoadCheck();
    scheduleActiveGalleryDateSync();
  }, true);
  window.addEventListener('resize', () => {
    scheduleMonthGalleryLoadCheck({ immediate: true });
    scheduleMainHeaderResponsiveLayout();
    scheduleActiveGalleryDateSync();
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
  const renderStartedAt = performance.now();
  cancelMonthGalleryScheduledWork();
  monthGalleryList.innerHTML = '';

  if (!currentSelection) {
    resetMonthGalleryRenderState();
    setActiveSidebarGroupDate('');
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
    setActiveSidebarGroupDate('');
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
  renderedGalleryGroupList = [];
  activeGalleryGroupIndex = -1;

  appendMonthGalleryPhotoBatch(
    Math.min(
      currentPhotos.length,
      Math.max(previousRenderedPhotoCount, calculateInitialVisiblePhotoCount())
    )
  );
  scheduleMonthGalleryLoadCheck({ immediate: true });
  scheduleActiveGalleryDateSync();
  syncKeyboardFocusedPhotoCard();

  logRendererPerformance('renderMonthGallery', performance.now() - renderStartedAt, {
    photoCount: currentPhotos.length,
    initialRenderedCount: renderedPhotoCount,
    resetProgressive: Boolean(resetProgressive),
  });
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

  for (const yearEntry of getOrderedTimelineSidebarData()) {
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
      const monthDayKey = getMonthDayKey(yearEntry.year, monthEntry.month);
      const isActiveMonth =
        isMonthSelection(currentSelection) &&
        currentSelection.year === yearEntry.year &&
        currentSelection.month === monthEntry.month;
      const isDayListOpen = expandedMonthDayKey === monthDayKey;
      const hasMonthDays =
        Array.isArray(monthEntry.days) && monthEntry.days.length > 0;

      if (isActiveMonth) {
        monthButton.classList.add('active');
      }

      if (isDayListOpen) {
        monthButton.classList.add('is-days-open');
      }

      const monthName = document.createElement('span');
      monthName.className = 'month-name';
      monthName.textContent = pad2(monthEntry.month);

      const monthCount = document.createElement('span');
      monthCount.className = 'month-count';
      monthCount.textContent = `${monthEntry.count}枚`;

      const monthMeta = document.createElement('span');
      monthMeta.className = 'month-meta';

      const monthChevron = document.createElement('span');
      monthChevron.className = 'month-day-toggle';
      monthChevron.setAttribute('aria-hidden', 'true');
      monthChevron.textContent = '▾';

      monthMeta.append(monthCount);

      if (hasMonthDays) {
        monthMeta.appendChild(monthChevron);
        monthButton.setAttribute(
          'aria-expanded',
          isDayListOpen ? 'true' : 'false'
        );
        monthButton.setAttribute(
          'aria-label',
          `${yearEntry.year}年${monthEntry.month}月 ${
            isDayListOpen ? '日付一覧を閉じる' : '日付一覧を開く'
          }`
        );
      }

      monthButton.appendChild(monthName);
      monthButton.appendChild(monthMeta);

      monthButton.addEventListener('click', async () => {
        if (isDayListOpen) {
          expandedMonthDayKey = '';

          if (isActiveMonth) {
            renderSidebar();
            return;
          }

          await selectMonth(yearEntry.year, monthEntry.month, {
            expandDays: false,
          });
          return;
        }

        expandedMonthDayKey = monthDayKey;
        await selectMonth(yearEntry.year, monthEntry.month, {
          expandDays: false,
        });
      });

      monthList.appendChild(monthButton);

      if (
        isDayListOpen &&
        hasMonthDays
      ) {
        const dayList = document.createElement('div');
        dayList.className = 'day-list';

        for (const dayEntry of monthEntry.days) {
          const groupDate =
            typeof dayEntry.groupDate === 'string'
              ? dayEntry.groupDate.trim()
              : '';

          if (!groupDate) {
            continue;
          }

          const dayButton = document.createElement('button');
          dayButton.className = 'day-button';
          dayButton.type = 'button';
          dayButton.dataset.groupDate = groupDate;

          const dayName = document.createElement('span');
          dayName.className = 'day-name';
          dayName.textContent = String(dayEntry.day);

          const dayCount = document.createElement('span');
          dayCount.className = 'day-count';
          dayCount.textContent = `${dayEntry.count}枚`;

          dayButton.append(dayName, dayCount);
          dayButton.addEventListener('click', async () => {
            await selectMonthDay(yearEntry.year, monthEntry.month, groupDate);
          });

          dayList.appendChild(dayButton);
        }

        monthList.appendChild(dayList);
      }
    }

    yearBlock.appendChild(yearButton);
    yearBlock.appendChild(monthList);
    sidebarTree.appendChild(yearBlock);
  }

  syncSidebarActiveDayButtons();
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
  const totalStartedAt = performance.now();
  const requestId = ++monthSelectionRequestId;
  clearMonthHeaderAnimationState();
  const monthSwitchOverlay = createMonthSwitchOverlay({
    includeHeader: false,
    viewportFixed: true,
  });
  const fetchStartedAt = performance.now();
  const photosPromise = fetchPhotos();
  let photos;
  let fetchMs = 0;

  try {
    photos = await photosPromise;
    fetchMs = performance.now() - fetchStartedAt;
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

  const renderStartedAt = performance.now();
  setCurrentMonthPhotos(photos);

  syncSelectionLinkedUi({ forceSidebarRender: true });
  stopScrollToTopAnimation();
  scrollGalleryViewToTop({ animated: false });
  renderMonthGallery({ resetProgressive: true });
  playMonthSwitchAnimation({ includeHeader: false });
  requestAnimationFrame(() => {
    fadeOutMonthSwitchOverlay(monthSwitchOverlay);
  });

  const renderMs = performance.now() - renderStartedAt;
  logRendererPerformance('selectPhotoScope', performance.now() - totalStartedAt, {
    mode: nextSelection?.mode || '',
    photoCount: Array.isArray(photos) ? photos.length : 0,
    fetchMs,
    renderMs,
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

async function selectMonth(year, month, { expandDays = false } = {}) {
  if (expandDays) {
    expandedMonthDayKey = getMonthDayKey(year, month);
  }

  await selectPhotoScope(
    () => window.electronAPI.getPhotosByMonth(year, month),
    createMonthSelection(year, month)
  );
}

async function selectMonthDay(year, month, groupDate) {
  const normalizedGroupDate =
    typeof groupDate === 'string' ? groupDate.trim() : '';

  if (!normalizedGroupDate) {
    return;
  }

  expandedMonthDayKey = getMonthDayKey(year, month);

  const isCurrentMonth =
    isMonthSelection(currentSelection) &&
    currentSelection.year === year &&
    currentSelection.month === month;

  if (!isCurrentMonth) {
    await selectMonth(year, month);
  } else {
    renderSidebar();
  }

  setActiveSidebarGroupDate(normalizedGroupDate);
  scrollMonthGalleryToGroupDate(normalizedGroupDate);
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
  } else if (normalizedSelection.mode === 'health') {
    await reloadHealthIssueSelection(normalizedSelection.kind);
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
  syncSettingsDataUi();
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

async function refreshViewAfterDataRestore() {
  closePhotoLabelModal();
  closeWorldNameEditModal();
  closeImageModal();

  trackedFolders = await window.electronAPI.getTrackedFolders();
  clearSelectionState();
  await restoreSidebarAndMonthSelection({ clearWhenEmpty: true });
}

async function runSettingsDataAction({
  isBlocked,
  confirmOptions = null,
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

  if (confirmOptions) {
    const confirmed = await openConfirmModal(confirmOptions);

    if (!confirmed) {
      return null;
    }
  }

  beginForegroundProgressOperation({
    statusMessage: busyStatus,
    progressMessage,
    showProgress: Boolean(progressMessage),
  });
  setSettingsDataSectionOpen(true);
  setSettingsDataStatus(busyStatus, 'busy');

  try {
    const result = await run();

    if (result?.canceled) {
      importStatus.textContent = '処理をキャンセルしました';
      setSettingsDataStatus('');
      return result;
    }

    if (!result?.ok) {
      throw new Error(result?.message || '処理に失敗しました');
    }

    if (typeof onSuccess === 'function') {
      await onSuccess(result);
    }

    if (typeof buildSuccessStatus === 'function') {
      const successStatus = buildSuccessStatus(result);
      importStatus.textContent = successStatus;
      setSettingsDataStatus(successStatus, 'success');
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
      setSettingsDataStatus(errorStatus, 'error');
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
      loadTrackedFolders: true,
      loadOverview: true,
    });
  }
}

async function createAppDataBackupFromSettings() {
  await runSettingsDataAction({
    isBlocked: () => isImporting || !window.electronAPI.createAppDataBackup,
    busyStatus: 'バックアップを作成中...',
    progressMessage: 'アプリデータを書き出しています...',
    run: () => window.electronAPI.createAppDataBackup(),
    buildSuccessStatus: (result) =>
      `バックアップ作成: ${result.fileName || '完了'} / 写真 ${
        result.photoCount || 0
      }件`,
    buildSuccessToast: () => 'バックアップを作成しました',
    buildErrorStatus: (message) => `バックアップ作成に失敗しました: ${message}`,
    buildErrorToast: (message) => `バックアップ作成に失敗しました: ${message}`,
  });
}

function buildHealthCheckStatus(result) {
  if (!result) {
    return '状態チェックに失敗しました';
  }

  if (result.healthy) {
    return `状態チェック: 問題なし / 写真 ${result.totalPhotoCount || 0}件`;
  }

  return [
    `状態チェック: 写真 ${result.totalPhotoCount || 0}件`,
    result.missingOriginalCount > 0
      ? `元画像なし ${result.missingOriginalCount}件`
      : null,
    result.missingThumbnailCount > 0
      ? `サムネイルなし ${result.missingThumbnailCount}件`
      : null,
    result.missingWorldInfoCount > 0
      ? `World情報未取得 ${result.missingWorldInfoCount}件`
      : null,
    result.worldMetadataIssueCount > 0
      ? `Worldメタデータ要確認 ${result.worldMetadataIssueCount}件`
      : null,
  ]
    .filter(Boolean)
    .join(' / ');
}

async function checkAppDataHealthFromSettings() {
  await runSettingsDataAction({
    isBlocked: () => isImporting || !window.electronAPI.checkAppDataHealth,
    busyStatus: '状態チェック中...',
    progressMessage: '登録データの状態を確認しています...',
    run: () => window.electronAPI.checkAppDataHealth(),
    buildSuccessStatus: buildHealthCheckStatus,
    buildSuccessToast: (result) =>
      result.healthy ? '状態チェック: 問題ありません' : '状態チェックが完了しました',
    buildErrorStatus: (message) => `状態チェックに失敗しました: ${message}`,
    buildErrorToast: (message) => `状態チェックに失敗しました: ${message}`,
  });
}

function getHealthIssueViewMeta(issueKind) {
  return (
    HEALTH_ISSUE_VIEW_META[issueKind] ||
    HEALTH_ISSUE_VIEW_META['world-metadata']
  );
}

function renderHealthIssuePhotos(issueKind, photos, { closeSettings = true } = {}) {
  const meta = getHealthIssueViewMeta(issueKind);

  setCurrentSelectionValue(createHealthSelection(issueKind, meta.label));
  setCurrentMonthPhotos(Array.isArray(photos) ? photos : []);
  clearSelectionState();
  syncSelectionLinkedUi({ forceSidebarRender: true });
  stopScrollToTopAnimation();
  scrollGalleryViewToTop({ animated: false });
  renderMonthGallery({ resetProgressive: true });

  if (closeSettings) {
    closeSettingsModal();
  }
}

async function fetchHealthIssuePhotos(issueKind) {
  if (window.electronAPI.getHealthIssuePhotos) {
    return window.electronAPI.getHealthIssuePhotos(issueKind);
  }

  if (
    issueKind === 'world-metadata' &&
    window.electronAPI.getWorldMetadataIssuePhotos
  ) {
    return window.electronAPI.getWorldMetadataIssuePhotos();
  }

  return {
    ok: false,
    message: '状態チェックカテゴリの抽出APIが利用できません',
    photoCount: 0,
    photos: [],
  };
}

async function reloadHealthIssueSelection(issueKind) {
  const result = await fetchHealthIssuePhotos(issueKind);

  if (!result?.ok) {
    throw new Error(result?.message || '状態チェック結果の再読み込みに失敗しました');
  }

  renderHealthIssuePhotos(issueKind, result.photos, { closeSettings: false });
  return result;
}

async function showHealthIssuePhotosFromSettings(issueKind) {
  const meta = getHealthIssueViewMeta(issueKind);

  await runSettingsDataAction({
    isBlocked: () => isImporting || !window.electronAPI.getHealthIssuePhotos,
    busyStatus: meta.busyStatus,
    progressMessage: meta.progressMessage,
    run: () => fetchHealthIssuePhotos(issueKind),
    onSuccess: async (result) => {
      renderHealthIssuePhotos(issueKind, result.photos);
    },
    buildSuccessStatus: (result) =>
      `${meta.successPrefix}: ${result.photoCount || 0}件を表示`,
    buildSuccessToast: (result) =>
      result.photoCount > 0
        ? meta.successToast(result.photoCount)
        : meta.emptyToast,
    buildErrorStatus: (message) => `${meta.errorPrefix}に失敗しました: ${message}`,
    buildErrorToast: (message) => `${meta.errorPrefix}に失敗しました: ${message}`,
  });
}

async function showWorldMetadataIssuePhotosFromSettings() {
  await showHealthIssuePhotosFromSettings('world-metadata');
}

async function regenerateMissingThumbnailsFromSettings() {
  await runSettingsDataAction({
    isBlocked: () =>
      isImporting || !window.electronAPI.regenerateMissingThumbnails,
    confirmOptions: {
      title: '欠損サムネイルを再生成',
      message:
        '状態チェックでサムネイルなしに該当する写真だけを再生成します。元画像が見つからない写真は失敗として記録されます。続行しますか？',
      confirmText: '再生成する',
    },
    busyStatus: '欠損サムネイルを再生成中...',
    progressMessage: 'サムネイルが欠損している写真だけを再生成しています...',
    run: () => window.electronAPI.regenerateMissingThumbnails(),
    onSuccess: async () => {
      if (isHealthSelection(currentSelection)) {
        await reloadHealthIssueSelection(currentSelection.kind);
      } else {
        await selectCurrentSelection();
      }
    },
    buildSuccessStatus: (result) =>
      `欠損サムネイル再生成: ${result.regeneratedCount || 0}件 / ` +
      `対象 ${result.totalCount || 0}件` +
      (result.failedCount > 0 ? ` / 失敗 ${result.failedCount}件` : ''),
    buildSuccessToast: (result) =>
      result.failedCount > 0
        ? `欠損サムネイル再生成: ${result.failedCount}件失敗しました`
        : '欠損サムネイルの再生成が完了しました',
    buildErrorStatus: (message) =>
      `欠損サムネイル再生成に失敗しました: ${message}`,
    buildErrorToast: (message) =>
      `欠損サムネイル再生成に失敗しました: ${message}`,
  });
}

async function refreshWorldMetadataIssuesFromSettings() {
  await runSettingsDataAction({
    isBlocked: () =>
      isImporting || !window.electronAPI.refreshWorldMetadataIssues,
    busyStatus: 'World要確認分を再取得中...',
    progressMessage: 'Worldメタデータ要確認の該当分だけ再取得しています...',
    run: () => window.electronAPI.refreshWorldMetadataIssues(),
    onSuccess: async () => {
      if (
        isHealthSelection(currentSelection) &&
        currentSelection.kind === 'world-metadata'
      ) {
        await reloadHealthIssueSelection('world-metadata');
      }
    },
    buildSuccessStatus: (result) =>
      `World要確認再取得: ${result.queuedCount || 0}件キュー投入 / ` +
      `対象World ${result.targetCount || 0}件 / ` +
      `該当写真 ${result.photoCount || 0}件`,
    buildSuccessToast: (result) =>
      result.queuedCount > 0
        ? `World要確認の再取得を${result.queuedCount}件開始しました`
        : '再取得が必要なWorld要確認はありません',
    buildErrorStatus: (message) => `World要確認再取得に失敗しました: ${message}`,
    buildErrorToast: (message) => `World要確認再取得に失敗しました: ${message}`,
  });
}

async function exportPhotoCatalogFromSettings(format) {
  const normalizedFormat = format === 'json' ? 'json' : 'csv';
  const formatLabel = normalizedFormat.toUpperCase();

  await runSettingsDataAction({
    isBlocked: () => isImporting || !window.electronAPI.exportPhotoCatalog,
    busyStatus: `${formatLabel}をエクスポート中...`,
    progressMessage: '写真一覧を書き出しています...',
    run: () => window.electronAPI.exportPhotoCatalog(normalizedFormat),
    buildSuccessStatus: (result) =>
      `${formatLabel}エクスポート: ${result.fileName || '完了'} / 写真 ${
        result.photoCount || 0
      }件`,
    buildSuccessToast: () => `${formatLabel}をエクスポートしました`,
    buildErrorStatus: (message) =>
      `${formatLabel}エクスポートに失敗しました: ${message}`,
    buildErrorToast: (message) =>
      `${formatLabel}エクスポートに失敗しました: ${message}`,
  });
}

async function restoreAppDataBackupFromSettings() {
  await runSettingsDataAction({
    isBlocked: () => isImporting || !window.electronAPI.restoreAppDataBackup,
    confirmOptions: {
      title: 'バックアップから復元',
      message:
        '現在の登録データ、ラベル、メモ、お気に入り、World情報、更新対象フォルダをバックアップ内容で置き換えます。元画像ファイル自体は削除しません。続行しますか？',
      confirmText: '復元する',
    },
    busyStatus: 'バックアップから復元中...',
    progressMessage: 'アプリデータを復元しています...',
    run: () => window.electronAPI.restoreAppDataBackup(),
    onSuccess: async () => {
      await refreshViewAfterDataRestore();
    },
    buildSuccessStatus: (result) =>
      `復元完了: 写真 ${result.photoCount || result.restoredPhotoCount || 0}件 / ` +
      `フォルダ ${result.trackedFolderCount || 0}件 / ` +
      `ラベル ${result.tagCount || 0}件` +
      (result.automaticBackupFileName
        ? ` / 復元前バックアップ ${result.automaticBackupFileName}`
        : ''),
    buildSuccessToast: () => 'バックアップから復元しました',
    buildErrorStatus: (message) => `バックアップ復元に失敗しました: ${message}`,
    buildErrorToast: (message) => `バックアップ復元に失敗しました: ${message}`,
  });
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
      `ラベル ${result.tagCount || 0}件` +
      (result.automaticBackupFileName
        ? ` / 初期化前バックアップ ${result.automaticBackupFileName}`
        : ''),
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
    renderSidebar();
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

  modalEditPhotoButton?.addEventListener('click', () => {
    openPhotoEditorModal();
  });

  bindSubModalCloseTriggers(
    worldNameEditBackdrop,
    worldNameEditClose,
    closeWorldNameEditModal
  );

  bindSubModalCloseTriggers(
    photoEditorBackdrop,
    photoEditorClose,
    closePhotoEditorModal
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

  photoEditorResetButton?.addEventListener('click', () => {
    resetPhotoEditorAll();
  });

  photoEditorCompareButton?.addEventListener('click', () => {
    togglePhotoEditorComparePreview();
  });

  photoEditorPresetResetButton?.addEventListener('click', () => {
    resetPhotoEditorAdjustments();
  });

  photoEditorAutoStrengthInput?.addEventListener('input', () => {
    updatePhotoEditorAutoEnhanceStrength(photoEditorAutoStrengthInput.value);
  });
  photoEditorAutoStrengthInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorAdjustmentResetButton?.addEventListener('click', () => {
    resetPhotoEditorAdjustments();
  });

  photoEditorAdjustmentTargetSelect?.addEventListener('change', () => {
    updatePhotoEditorAdjustmentTarget(photoEditorAdjustmentTargetSelect.value);
  });

  photoEditorSavePresetButton?.addEventListener('click', () => {
    saveCurrentPhotoEditorPreset();
  });

  photoEditorPresetNameInput?.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      saveCurrentPhotoEditorPreset();
    }
  });

  photoEditorCropResetButton?.addEventListener('click', () => {
    resetPhotoEditorCrop();
  });

  photoEditorTextResetButton?.addEventListener('click', () => {
    resetPhotoEditorTextOverlay();
  });

  photoEditorTextAddButton?.addEventListener('click', () => {
    addPhotoEditorTextOverlay();
  });

  photoEditorTextDeleteButton?.addEventListener('click', () => {
    deletePhotoEditorActiveTextOverlay();
  });

  photoEditorTextContentInput?.addEventListener('input', () => {
    updatePhotoEditorTextOverlay({
      text: photoEditorTextContentInput.value,
      enabled: Boolean(photoEditorTextContentInput.value.trim()),
    });
  });
  photoEditorTextContentInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorTextFontSelect?.addEventListener('change', () => {
    syncPhotoEditorTextFontSelectPreview(photoEditorTextFontSelect.value);
    updatePhotoEditorTextOverlay(
      { fontKey: photoEditorTextFontSelect.value },
      { interactive: false }
    );
  });

  photoEditorTextSizeInput?.addEventListener('input', () => {
    updatePhotoEditorTextOverlay({ size: photoEditorTextSizeInput.value });
  });
  photoEditorTextSizeInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorTextColorInput?.addEventListener('input', () => {
    updatePhotoEditorTextOverlay({ color: photoEditorTextColorInput.value });
  });
  photoEditorTextColorInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorTextWeightSelect?.addEventListener('change', () => {
    updatePhotoEditorTextOverlay(
      { weight: photoEditorTextWeightSelect.value },
      { interactive: false }
    );
  });

  photoEditorTextStrokeTypeSelect?.addEventListener('change', () => {
    updatePhotoEditorTextOverlay(
      { strokeType: photoEditorTextStrokeTypeSelect.value },
      { interactive: false }
    );
  });

  photoEditorTextStrokeWidthInput?.addEventListener('input', () => {
    updatePhotoEditorTextOverlay({
      strokeWidth: photoEditorTextStrokeWidthInput.value,
    });
  });
  photoEditorTextStrokeWidthInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorTextStrokeColorInput?.addEventListener('input', () => {
    updatePhotoEditorTextOverlay({
      strokeColor: photoEditorTextStrokeColorInput.value,
    });
  });
  photoEditorTextStrokeColorInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorTextFillTransparentInput?.addEventListener('click', () => {
    updatePhotoEditorTextOverlay(
      {
        fillTransparent: !getPhotoEditorActiveTextOverlay()?.fillTransparent,
      },
      { interactive: false }
    );
    finishPhotoEditorInteractivePreview();
  });

  photoEditorTextMaskModeSelect?.addEventListener('change', () => {
    updatePhotoEditorTextOverlay(
      { maskMode: photoEditorTextMaskModeSelect.value },
      { interactive: false }
    );
    finishPhotoEditorInteractivePreview();
  });

  photoEditorTextLetterSpacingInput?.addEventListener('input', () => {
    updatePhotoEditorTextOverlay({
      letterSpacing: photoEditorTextLetterSpacingInput.value,
    });
  });
  photoEditorTextLetterSpacingInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorImageOverlayResetButton?.addEventListener('click', () => {
    resetPhotoEditorImageOverlays();
  });

  photoEditorImageOverlayAddButton?.addEventListener('click', async () => {
    await selectPhotoEditorOverlayImages();
  });

  photoEditorImageOverlayOpacityInput?.addEventListener('input', () => {
    updatePhotoEditorActiveImageOverlay({
      opacity:
        clampNumber(photoEditorImageOverlayOpacityInput.value, 0, 100, 100) / 100,
    });
  });

  photoEditorImageOverlayOpacityInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorImageOverlayBlendModeSelect?.addEventListener('change', () => {
    updatePhotoEditorActiveImageOverlay(
      {
        blendMode: photoEditorImageOverlayBlendModeSelect.value,
      },
      { interactive: false }
    );
    finishPhotoEditorInteractivePreview();
  });

  photoEditorImageOverlayMaskModeSelect?.addEventListener('change', () => {
    updatePhotoEditorActiveImageOverlay(
      {
        maskMode: photoEditorImageOverlayMaskModeSelect.value,
      },
      { interactive: false }
    );
    finishPhotoEditorInteractivePreview();
  });

  photoEditorImageOverlayForwardButton?.addEventListener('click', () => {
    movePhotoEditorActiveImageOverlay(1);
  });

  photoEditorImageOverlayBackwardButton?.addEventListener('click', () => {
    movePhotoEditorActiveImageOverlay(-1);
  });

  photoEditorImageOverlayDeleteButton?.addEventListener('click', () => {
    deletePhotoEditorActiveImageOverlay();
  });

  photoEditorSubjectResetButton?.addEventListener('click', () => {
    clearPhotoEditorSubjectMask();
  });

  photoEditorSubjectAutoButton?.addEventListener('click', async () => {
    await generatePhotoEditorAutoSubjectMask();
  });

  photoEditorSubjectStandardDownloadButton?.addEventListener('click', async () => {
    await downloadSettingsAiSubjectModel('withoutbg-snap');
  });

  photoEditorSubjectHighQualityDownloadButton?.addEventListener('click', async () => {
    await downloadSettingsAiSubjectModel('withoutbg-focus');
  });

  photoEditorSubjectHighQualityButton?.addEventListener('click', async () => {
    await generatePhotoEditorHighQualitySubjectMask();
  });

  photoEditorSubjectImportButton?.addEventListener('click', () => {
    photoEditorSubjectFileInput?.click();
  });

  photoEditorSubjectFileInput?.addEventListener('change', () => {
    const file = photoEditorSubjectFileInput.files?.[0];
    loadPhotoEditorSubjectMaskFile(file);
    photoEditorSubjectFileInput.value = '';
  });

  photoEditorSubjectDeleteButton?.addEventListener('click', () => {
    clearPhotoEditorSubjectMask();
  });

  photoEditorSubjectTransparentSaveButton?.addEventListener('click', async () => {
    await savePhotoEditorImage({ transparentSubjectBackground: true });
  });

  photoEditorSubjectShowMaskInput?.addEventListener('click', () => {
    updatePhotoEditorSubjectMask(
      { showOverlay: !photoEditorState?.subjectMask?.showOverlay },
      { interactive: false }
    );
    finishPhotoEditorInteractivePreview();
  });

  photoEditorSubjectInvertInput?.addEventListener('click', () => {
    updatePhotoEditorSubjectMask(
      { invert: !photoEditorState?.subjectMask?.invert },
      { interactive: false }
    );
    finishPhotoEditorInteractivePreview();
  });

  photoEditorExportResetButton?.addEventListener('click', () => {
    resetPhotoEditorExportSettings();
  });

  photoEditorBlurResetButton?.addEventListener('click', () => {
    resetPhotoEditorBlur();
  });

  photoEditorCurveResetButton?.addEventListener('click', () => {
    resetPhotoEditorCurve();
  });

  photoEditorBlurConfirmButton?.addEventListener('click', () => {
    confirmPhotoEditorBlur();
  });

  photoEditorSaveButton?.addEventListener('click', async () => {
    await savePhotoEditorImage();
  });

  photoEditorUndoButton?.addEventListener('click', () => {
    undoPhotoEditorEdit();
  });

  photoEditorRedoButton?.addEventListener('click', () => {
    redoPhotoEditorEdit();
  });

  photoEditorRuleGridButton?.addEventListener('click', () => {
    if (!photoEditorState) {
      return;
    }

    photoEditorState.showRuleOfThirdsGrid = !photoEditorState.showRuleOfThirdsGrid;
    syncPhotoEditorOverlayControls();
    if (!paintPhotoEditorPreviewOverlayOnly()) {
      schedulePhotoEditorRender();
    }
  });

  photoEditorRulerButton?.addEventListener('click', () => {
    if (!photoEditorState) {
      return;
    }

    photoEditorState.showRulers = !photoEditorState.showRulers;
    photoEditorState.snapGuide = null;
    if (!photoEditorState.showRulers) {
      photoEditorState.draftRulerGuide = null;
    }
    syncPhotoEditorOverlayControls();
    if (!paintPhotoEditorPreviewOverlayOnly()) {
      schedulePhotoEditorRender();
    }
  });

  for (const toggle of photoEditorAccordionToggles) {
    toggle.addEventListener('click', () => {
      const key = toggle.dataset.photoEditorAccordionToggle;

      if (!key) {
        return;
      }

      const shouldOpen = toggle.getAttribute('aria-expanded') !== 'true';
      setPhotoEditorAccordionOpen(key, shouldOpen);

      if (key === 'mask' && !shouldOpen) {
        cancelPhotoEditorPendingMask();
      }

      if (key === 'blur') {
        syncPhotoEditorBlurControls();
        schedulePhotoEditorRender();
      }

      if (key === 'curve') {
        syncPhotoEditorCurveControls();
      }

      if (key === 'text') {
        syncPhotoEditorTextControls();
      }

      if (key === 'imageOverlay') {
        syncPhotoEditorImageOverlayControls();
        if (!paintPhotoEditorPreviewOverlayOnly()) {
          schedulePhotoEditorRender();
        }
      }
    });
  }

  [
    photoEditorResetButton,
    photoEditorUndoButton,
    photoEditorRedoButton,
    photoEditorCompareButton,
    photoEditorSaveButton,
    photoEditorPresetResetButton,
    photoEditorAdjustmentResetButton,
    ...photoEditorAccordionToggles,
    photoEditorPresetList,
    photoEditorPresetNameInput,
    photoEditorSavePresetButton,
    photoEditorCropPresetList,
    photoEditorCropResetButton,
    photoEditorCropRotateLeftButton,
    photoEditorCropRotateRightButton,
    photoEditorCropFlipXButton,
    photoEditorCropFlipYButton,
    photoEditorCropZoomInput,
    photoEditorCropTiltInput,
    photoEditorCropXInput,
    photoEditorCropYInput,
    photoEditorTextResetButton,
    photoEditorTextAddButton,
    photoEditorTextDeleteButton,
    photoEditorTextList,
    photoEditorTextContentInput,
    photoEditorTextFontSelect,
    photoEditorTextSizeInput,
    photoEditorTextColorInput,
    photoEditorTextWeightSelect,
    photoEditorTextStrokeTypeSelect,
    photoEditorTextStrokeWidthInput,
    photoEditorTextStrokeColorInput,
    photoEditorTextFillTransparentInput,
    photoEditorTextLetterSpacingInput,
    photoEditorImageOverlayResetButton,
    photoEditorImageOverlayAddButton,
    photoEditorImageOverlayLibrary,
    photoEditorImageOverlayList,
    photoEditorImageOverlayOpacityInput,
    photoEditorImageOverlayBlendModeSelect,
    photoEditorImageOverlayForwardButton,
    photoEditorImageOverlayBackwardButton,
    photoEditorImageOverlayDeleteButton,
    photoEditorExportResetButton,
    photoEditorExportFormatSelect,
    photoEditorExportMaxEdgeSelect,
    photoEditorExportQualityInput,
    photoEditorBlurResetButton,
    photoEditorBlurModeGroup,
    photoEditorBlurAmountInput,
    photoEditorBlurConfirmButton,
    photoEditorCurveResetButton,
    photoEditorCurveModeGroup,
    photoEditorCurveChannelList,
    photoEditorCurveCanvas,
    photoEditorAdjustmentList,
  ]
    .filter(Boolean)
    .forEach((element) => {
      element.addEventListener('pointerdown', cancelPhotoEditorPendingMaskFromControl);
    });

  photoEditorCropZoomInput?.addEventListener('input', () => {
    updatePhotoEditorCrop({ zoom: photoEditorCropZoomInput.value });
  });
  photoEditorCropZoomInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorCropRotateLeftButton?.addEventListener('click', () => {
    rotatePhotoEditorCrop(-90);
  });

  photoEditorCropRotateRightButton?.addEventListener('click', () => {
    rotatePhotoEditorCrop(90);
  });

  photoEditorCropFlipXButton?.addEventListener('click', () => {
    togglePhotoEditorCropFlipX();
  });

  photoEditorCropFlipYButton?.addEventListener('click', () => {
    togglePhotoEditorCropFlipY();
  });

  photoEditorCropTiltInput?.addEventListener('input', () => {
    updatePhotoEditorCrop({ tilt: photoEditorCropTiltInput.value });
  });
  photoEditorCropTiltInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorCropXInput?.addEventListener('input', () => {
    updatePhotoEditorCrop({ offsetX: photoEditorCropXInput.value });
  });
  photoEditorCropXInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorCropYInput?.addEventListener('input', () => {
    updatePhotoEditorCrop({ offsetY: photoEditorCropYInput.value });
  });
  photoEditorCropYInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorExportFormatSelect?.addEventListener('change', () => {
    updatePhotoEditorExportSettings({
      format: photoEditorExportFormatSelect.value,
    });
  });

  photoEditorExportMaxEdgeSelect?.addEventListener('change', () => {
    updatePhotoEditorExportSettings({
      maxEdge: photoEditorExportMaxEdgeSelect.value,
    });
  });

  photoEditorExportQualityInput?.addEventListener('input', () => {
    updatePhotoEditorExportSettings({
      quality: photoEditorExportQualityInput.value,
    });
  });
  photoEditorExportQualityInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorBlurModeGroup?.addEventListener('click', (event) => {
    const modeButton = event.target.closest('[data-photo-editor-blur-mode]');

    if (!modeButton || !photoEditorState) {
      return;
    }

    updatePhotoEditorBlur({
      mode: modeButton.dataset.photoEditorBlurMode || 'full',
    });
  });

  photoEditorBlurAmountInput?.addEventListener('input', () => {
    updatePhotoEditorBlur({ amount: photoEditorBlurAmountInput.value });
  });
  photoEditorBlurAmountInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorCurveModeGroup?.addEventListener('click', (event) => {
    const modeButton = event.target.closest('[data-photo-editor-curve-mode]');

    if (!modeButton || !photoEditorState) {
      return;
    }

    setPhotoEditorCurveMode(modeButton.dataset.photoEditorCurveMode || 'rgb');
  });

  photoEditorMaskToolGroup?.addEventListener('click', (event) => {
    const toolButton = event.target.closest('[data-photo-editor-mask-tool]');

    if (!toolButton || !photoEditorState) {
      return;
    }

    setPhotoEditorMaskTool(toolButton.dataset.photoEditorMaskTool || 'none');
  });

  photoEditorMaskShapeGroup?.addEventListener('click', (event) => {
    const shapeButton = event.target.closest('[data-photo-editor-mask-shape]');

    if (!shapeButton || !photoEditorState) {
      return;
    }

    setPhotoEditorMaskShape(shapeButton.dataset.photoEditorMaskShape || 'rect');
  });

  photoEditorMaskBlurStrengthInput?.addEventListener('input', () => {
    setPhotoEditorBlurStrength(photoEditorMaskBlurStrengthInput.value);
  });
  photoEditorMaskBlurStrengthInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorFillColorInput?.addEventListener('input', () => {
    if (!photoEditorState) {
      return;
    }

    beginPhotoEditorHistoryMutation();
    photoEditorState.fillColor = photoEditorFillColorInput.value || '#111827';
    if (photoEditorState.draftMask?.type === 'fill') {
      updatePhotoEditorDraftMask({ color: photoEditorState.fillColor });
    }
    syncPhotoEditorMaskToolUi();
    schedulePhotoEditorHistoryCommit();
  });
  photoEditorFillColorInput?.addEventListener('change', () => {
    finishPhotoEditorInteractivePreview();
  });

  photoEditorMaskConfirmButton?.addEventListener('click', () => {
    confirmPhotoEditorMask();
  });

  photoEditorMaskUndoButton?.addEventListener('click', () => {
    undoPhotoEditorMask();
  });

  photoEditorMaskClearButton?.addEventListener('click', () => {
    clearPhotoEditorMasks();
  });

  photoEditorCanvas?.addEventListener('pointerdown', beginPhotoEditorMaskDrag);
  photoEditorCanvas?.addEventListener('pointermove', updatePhotoEditorMaskDrag);
  photoEditorCanvas?.addEventListener('pointerup', finishPhotoEditorMaskDrag);
  photoEditorCanvas?.addEventListener('pointercancel', finishPhotoEditorMaskDrag);
  photoEditorCanvas?.addEventListener('wheel', handlePhotoEditorCanvasWheel, {
    passive: false,
  });
  photoEditorCurveCanvas?.addEventListener('pointerdown', beginPhotoEditorCurveDrag);
  photoEditorCurveCanvas?.addEventListener('pointermove', updatePhotoEditorCurveDrag);
  photoEditorCurveCanvas?.addEventListener('pointerup', finishPhotoEditorCurveDrag);
  photoEditorCurveCanvas?.addEventListener(
    'pointercancel',
    finishPhotoEditorCurveDrag
  );
  window.addEventListener('resize', () => {
    if (
      photoEditorState?.sourceImage &&
      photoEditorModal?.classList.contains('hidden') === false
    ) {
      drawPhotoEditorCurveCanvas();
      schedulePhotoEditorRender();
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

  settingsDataToggleButton?.addEventListener('click', () => {
    setSettingsDataSectionOpen(!isSettingsDataSectionOpen);
    syncSettingsDataUi();
  });

  settingsAiModelList?.addEventListener('click', async (event) => {
    const downloadButton = event.target.closest('[data-ai-subject-model-download]');
    const deleteButton = event.target.closest('[data-ai-subject-model-delete]');
    const folderButton = event.target.closest('[data-ai-subject-model-folder]');

    if (downloadButton) {
      await downloadSettingsAiSubjectModel(
        downloadButton.dataset.aiSubjectModelDownload
      );
      return;
    }

    if (folderButton) {
      await openSettingsAiSubjectModelFolder(
        folderButton.dataset.aiSubjectModelFolder
      );
      return;
    }

    if (deleteButton) {
      await deleteSettingsAiSubjectModel(deleteButton.dataset.aiSubjectModelDelete);
    }
  });

  window.electronAPI.onAiSubjectModelDownloadProgress?.((payload) => {
    const model = getPhotoEditorSubjectModel(payload?.modelId);
    const receivedBytes = Number(payload?.receivedBytes) || 0;
    const totalBytes = Number(payload?.totalBytes) || 0;
    const fileName = payload?.fileName ? ` / ${payload.fileName}` : '';
    const receivedText = formatAiSubjectModelSize(receivedBytes);
    const totalText = formatAiSubjectModelSize(totalBytes);

    setSettingsAiModelStatus(
      totalBytes > 0
        ? `${model?.displayName || model?.label || 'AIモデル'}をダウンロード中${fileName}... ${receivedText} / ${totalText}`
        : `${model?.displayName || model?.label || 'AIモデル'}をダウンロード中${fileName}... ${receivedText}`,
      'busy'
    );
  });

  createAppDataBackupButton?.addEventListener('click', async () => {
    await createAppDataBackupFromSettings();
  });

  checkAppDataHealthButton?.addEventListener('click', async () => {
    await checkAppDataHealthFromSettings();
  });

  showMissingOriginalFilesButton?.addEventListener('click', async () => {
    await showHealthIssuePhotosFromSettings('missing-original');
  });

  showMissingThumbnailsButton?.addEventListener('click', async () => {
    await showHealthIssuePhotosFromSettings('missing-thumbnail');
  });

  showMissingWorldInfoButton?.addEventListener('click', async () => {
    await showHealthIssuePhotosFromSettings('missing-world-info');
  });

  showWorldMetadataIssuesButton?.addEventListener('click', async () => {
    await showWorldMetadataIssuePhotosFromSettings();
  });

  regenerateMissingThumbnailsButton?.addEventListener('click', async () => {
    await regenerateMissingThumbnailsFromSettings();
  });

  refreshWorldMetadataIssuesButton?.addEventListener('click', async () => {
    await refreshWorldMetadataIssuesFromSettings();
  });

  restoreAppDataBackupButton?.addEventListener('click', async () => {
    await restoreAppDataBackupFromSettings();
  });

  exportPhotoCatalogCsvButton?.addEventListener('click', async () => {
    await exportPhotoCatalogFromSettings('csv');
  });

  exportPhotoCatalogJsonButton?.addEventListener('click', async () => {
    await exportPhotoCatalogFromSettings('json');
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

  window.addEventListener('worldshot:languagechange', () => {
    syncFontPreferenceForLanguage();
    renderSettingsAiSubjectModels();
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

    if (photoEditorModal && !photoEditorModal.classList.contains('hidden')) {
      closePhotoEditorModal();
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
      photoLabelModal?.classList.contains('hidden') === false ||
      photoEditorModal?.classList.contains('hidden') === false
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
      photoEditorModal?.classList.contains('hidden') === false ||
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
  initializePhotoEditorUi();
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


