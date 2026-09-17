/**
 * Markdown binding of the shared app-owned asset lifecycle
 * (@airy-office/electron-utils). Sibling .md/.markdown documents may reference
 * the same assets/ directory, so the GC pattern scans those extensions.
 */
import {
  createAssetLifecycle,
  extractImageSources,
  resolveSafeRelativeImagePath,
  rewriteImageSources,
} from '@airy-office/electron-utils'

export { resolveSafeRelativeImagePath }
export const extractMarkdownImageSources = extractImageSources
export const rewriteMarkdownImageSources = rewriteImageSources

export const {
  copyImageIntoOwnedAssets,
  discardPendingOwnedAssets,
  pendingOwnedAssetsForDocument,
  prepareAssetsForSaveAs,
  readOwnedAssetManifest,
  reconcileOwnedAssets,
  renameOwnedAssetDocument,
  resolveSourcePendingAfterSaveAs,
  rollbackPreparedSaveAsAssets,
  writeImageIntoOwnedAssets,
} = createAssetLifecycle({
  appLabel: 'markdown',
  siblingDocumentPattern: /\.(?:md|markdown)$/i,
})
