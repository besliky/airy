export {
  buildContextMenuItems,
  contextMenuLabels,
  installContextMenu,
  type ContextMenuItem,
  type ContextMenuLabels,
} from './context-menu'
export {
  appMenuLabels,
  editMenuTemplate,
  toggleDevToolsItem,
  viewMenuTemplate,
  windowMenuTemplate,
  type AppMenuLabels,
} from './app-menu'
export { COPILOT_GUIDE_URL, DOCS_README_URL, GITHUB_REPO_URL, openHelpUrl } from './github-menu'
export { atomicWriteFile, looksLikeZip } from './atomic-write'
export {
  ASSET_MANIFEST_FILENAME,
  createAssetLifecycle,
  extractImageSources,
  rewriteImageSources,
  resolveSafeRelativeImagePath,
  type AssetLifecycle,
  type AssetLifecycleScope,
  type AssetRewrite,
  type OwnedAssetManifest,
  type OwnedAssetRecord,
  type PreparedSaveAsAssets,
  type ReconcileResult,
} from './asset-lifecycle'
export {
  saveAsSuggestion,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
} from './dialog-memory'
export { ALL_OPEN_EXTENSIONS, OPEN_EXTENSION_GROUPS } from './open-filters'
export {
  TextRecoveryStore,
  shouldOfferTextRecovery,
  textRecoveryPathFor,
  type RecoveryDecision,
} from './text-recovery'
export {
  recordDialogDir,
  readLastDialogDirs,
  writeLastDialogDir,
  type DialogDirEntry,
} from './dialog-memory'
export {
  DEFAULT_SAVE_DIR_KEY,
  configuredDefaultSaveDir,
  isUsableSaveDir,
  readDefaultSaveDirSetting,
  resolveDefaultSaveDir,
  type PathProvider,
} from './default-save-dir'
export {
  AUTHOR_NAME_KEY,
  AUTHOR_NAME_MAX,
  configuredAuthorName,
  readAuthorNameSetting,
  sanitizeAuthorName,
} from './author-name'
export { installNavigationGuard } from './navigation-guard'
export {
  DROP_OPEN_CHANNEL,
  droppableFilePaths,
  installDropOpenBridge,
  KNOWN_UNSUPPORTED_DOC_RE,
  OPENABLE_DOC_RE,
  partitionDropPayload,
} from './drop-open'
export { safeExternalUrl, type SafeExternalUrlOptions } from './safe-external-url'
export {
  fetchWithSsrfGuard,
  isBlockedAddress,
  isSafeRemoteUrl,
  type FetchWithSsrfGuardOptions,
} from './safe-remote-url'
export { fetchRemoteImage, remoteImageHeaders } from './remote-image'
export { GENERATED_IMAGE_DIR, readGeneratedImage, storeGeneratedImage } from './generated-images'
export {
  buildPrintableHtml,
  printHtmlToPdf,
  sanitizePrintableBody,
  type PrintableHtml,
  type PrintWindow,
} from './print-html-pdf'
export { crashErrorPageUrl, isRecoverableRendererCrash, voidLoad } from './process-safety'
export {
  forgetRendererFileAccess,
  grantRendererDir,
  grantRendererFileAccess,
  grantedRendererDirs,
  isPathInsideDir,
  pathIsInsideAny,
  rendererMayReadPath,
  resetRendererFileGrants,
} from './renderer-file-access'
export {
  forgetWitnessedDrops,
  MAX_ATTACHMENT_ADD_PATHS,
  MAX_ATTACHMENT_PATH_CHARS,
  mayGrantAttachmentRead,
  parseAttachmentPaths,
  recordWitnessedDrops,
  resetWitnessedDrops,
  witnessedDroppedPath,
} from './witnessed-drops'
export { WITNESS_DROP_CHANNEL } from './witness-channel'
