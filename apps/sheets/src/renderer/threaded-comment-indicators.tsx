import { useEffect, useState } from 'react'

import type { UniverRuntime } from './univer-state'
// module-level t: Univer mounts this chrome in its own React root, outside
// the app's I18n context
import { t } from './i18n/locale'
import { formatAddress } from '../domain/cell-address'
import {
  getThread,
  getThreadedCommentsSnapshot,
  subscribeThreadedComments,
  type ThreadedCommentThread,
} from './threaded-comments'

const INDICATOR_KEY_PREFIX = 'airy-thread-indicator'
const INDICATOR_SIZE = 12

/// Invoked when the user clicks a cell indicator; the app opens the comments
/// panel scrolled to that thread.
let openThreadListener: ((thread: ThreadedCommentThread) => void) | null = null

export function setThreadPanelOpener(
  listener: ((thread: ThreadedCommentThread) => void) | null,
): void {
  openThreadListener = listener
}

function ThreadIndicator({
  sheetId,
  threadId,
  fallback,
}: {
  readonly sheetId: string
  readonly threadId: string
  readonly fallback: ThreadedCommentThread
}): React.JSX.Element {
  // Live lookup: the badge re-renders when the store changes (a thread turned
  // resolved shows the muted style immediately).
  const [, setTick] = useState(0)
  useEffect(() => subscribeThreadedComments(() => setTick((value) => value + 1)), [])
  const thread = getThread(sheetId, threadId) ?? fallback
  return (
    <button
      type="button"
      tabIndex={0}
      className={thread.resolved ? 'thread-indicator resolved' : 'thread-indicator'}
      role="button"
      aria-label={`${t('appThreadComments')} ${formatAddress(thread.row, thread.column)}`}
      onClick={() => openThreadListener?.(thread)}
      onKeyDown={(event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault()
          openThreadListener?.(thread)
        }
      }}
    >
      <svg aria-hidden="true" viewBox="0 0 12 12">
        <path d="M0 0h12L0 12Z" />
      </svg>
    </button>
  )
}

interface IndicatorHandle {
  readonly dispose: () => void
}

/// Pins a small triangle indicator (Excel's comment badge) to every threaded
/// cell. Indicators are float DOMs anchored to the thread's cell, so they
/// follow scroll and zoom; the map tracks one handle per sheet:thread pair
/// and only installs/disposes the diff on each store or sheet change.
export function installThreadedCommentIndicators(runtime: UniverRuntime): {
  dispose: () => void
} {
  const handles = new Map<string, IndicatorHandle>()
  let disposed = false

  const clear = (): void => {
    for (const handle of handles.values()) handle.dispose()
    handles.clear()
  }

  const sync = (): void => {
    if (disposed) return
    const desired = new Map<string, ThreadedCommentThread>()
    for (const [sheetId, threads] of getThreadedCommentsSnapshot()) {
      for (const thread of threads) desired.set(`${sheetId}:${thread.id}`, thread)
    }
    for (const [key, handle] of handles) {
      if (!desired.has(key)) {
        handle.dispose()
        handles.delete(key)
      }
    }
    for (const [key, thread] of desired) {
      if (handles.has(key)) continue
      const [sheetId] = splitKey(key)
      const workbook = runtime.univerAPI.getActiveWorkbook()
      const worksheet = workbook?.getSheetBySheetId(sheetId)
      if (!workbook || !worksheet) continue
      const componentKey = `${INDICATOR_KEY_PREFIX}-${handles.size}-${thread.id.slice(-8)}`
      let registered: DisposableLike | null = null
      try {
        registered = runtime.univerAPI.registerComponent(componentKey, () => (
          <ThreadIndicator sheetId={sheetId} threadId={thread.id} fallback={thread} />
        ))
        const floating = worksheet.addFloatDomToRange(
          worksheet.getRange(thread.row, thread.column, 1, 1),
          {
            componentKey,
            allowTransform: false,
            eventPassThrough: false,
          },
          {
            width: INDICATOR_SIZE,
            height: INDICATOR_SIZE,
            marginX: '100%',
            marginY: 0,
          },
          componentKey,
        )
        handles.set(key, {
          dispose: () => {
            floating?.dispose()
            disposeComponent(registered)
          },
        })
      } catch {
        // Indicators are best-effort decoration; the panel still lists the
        // thread even if Univer refuses the float DOM.
        disposeComponent(registered)
      }
    }
  }

  const disposables = [
    subscribeThreadedComments(sync),
    runtime.univerAPI.addEvent(runtime.univerAPI.Event.ActiveSheetChanged, () => sync()),
  ]
  sync()

  return {
    dispose: () => {
      disposed = true
      clear()
      for (const disposable of disposables) disposeComponent(disposable)
    },
  }
}

function splitKey(key: string): [string, string] {
  const separator = key.indexOf(':')
  return separator === -1 ? [key, ''] : [key.slice(0, separator), key.slice(separator + 1)]
}

/// Univer's disposeables changed shape across versions (disposeable object vs
/// plain cleanup function); both are accepted here.
type DisposableLike = { dispose(): void } | (() => void)

function disposeComponent(registered: DisposableLike | null): void {
  if (registered === null) return
  if (typeof registered === 'function') registered()
  else registered.dispose()
}
