/**
 * HTML binding of the shared app-owned asset lifecycle
 * (@airy-office/electron-utils). Sibling .html/.htm documents may reference
 * the same assets/ directory, so the GC pattern scans those extensions.
 */
import {
  createAssetLifecycle,
  extractImageSources,
  resolveSafeRelativeImagePath,
  rewriteImageSources,
} from '@airy-office/electron-utils'

export { resolveSafeRelativeImagePath }
export const extractHtmlImageSources = extractImageSources
export const rewriteHtmlImageSources = rewriteImageSources

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
  appLabel: 'html',
  siblingDocumentPattern: /\.html?$/i,
})
