import { friendlyErrorKey, type FriendlyErrorKey } from '../../shared/error-codes'
import type { StringKey } from './locale'
import type { TFunc } from './locale'
import { showToast } from '@airy-office/ui/toast-bus'

/** renderer localization per friendly-error key */
const KEY_TO_STRING: Record<FriendlyErrorKey, StringKey> = {
  enoent: 'errFileNotFound',
  eperm: 'errPermissionDenied',
  ebusy: 'errFileLocked',
  emfile: 'errTooManyFiles',
}

/**
 * Show an operation failure as the styled error toast instead of a raw
 * window.alert: known filesystem errors get the friendly localized message,
 * anything else falls back to the raw error text.
 */
export function showErrorToast(error: unknown, t: TFunc): void {
  const key = friendlyErrorKey(error)
  const text = key ? t(KEY_TO_STRING[key]) : error instanceof Error ? error.message : String(error)
  showToast(text, 'error')
}
