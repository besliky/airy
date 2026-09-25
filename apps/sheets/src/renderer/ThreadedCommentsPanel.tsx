import { useEffect, useRef, useState } from 'react'

import { formatAddress } from '../domain/cell-address'
import { useI18n } from './i18n/locale'
import type { UniverRuntime } from './univer-state'
import {
  addThread,
  appendReply,
  deleteThread,
  setThreadResolved,
  subscribeThreadedComments,
  threadsForSheet,
  type ThreadedCommentThread,
} from './threaded-comments'

/// Excel's Comments pane: a floating, non-modal panel listing the active
/// sheet's threads. Authors, timestamps, reply chains, resolve and delete all
/// surface here; the grid carries only the cell indicator.
export function ThreadedCommentsPanel({
  runtime,
  sheetId,
  onClose,
  onNavigate,
  draftThread,
  composerFocusSignal,
}: {
  readonly runtime: UniverRuntime
  readonly sheetId: string | null
  readonly onClose: () => void
  /// Selects the thread's anchor cell in the grid.
  readonly onNavigate: (thread: ThreadedCommentThread) => void
  /// A thread freshly created from the ribbon; its reply input takes focus.
  readonly draftThread: ThreadedCommentThread | null
  /// Bumped to move focus to the new-thread composer (Review › New Comment).
  readonly composerFocusSignal: number
}): React.JSX.Element {
  const { t } = useI18n()
  const [snapshotTick, setSnapshotTick] = useState(0)
  const [replyDrafts, setReplyDrafts] = useState<Map<string, string>>(new Map())
  const draftInputRef = useRef<HTMLInputElement | null>(null)
  const composerInputRef = useRef<HTMLInputElement | null>(null)
  const lastDraftIdRef = useRef<string | null>(null)
  const lastComposerSignalRef = useRef(composerFocusSignal)

  useEffect(
    () =>
      subscribeThreadedComments(() => {
        setSnapshotTick((value) => value + 1)
      }),
    [],
  )
  // The tick re-renders on store changes; the list itself is read fresh each
  // render (small, in-memory).
  void snapshotTick
  const threads = threadsForSheet(sheetId)

  const setReplyDraft = (threadId: string, value: string): void => {
    setReplyDrafts((previous) => {
      const next = new Map(previous)
      next.set(threadId, value)
      return next
    })
  }

  const replyText = (threadId: string): string => (replyDrafts.get(threadId) ?? '').trim()

  const submitReply = (threadId: string): void => {
    const text = replyText(threadId)
    if (text === '' || sheetId === null) return
    appendReply(sheetId, threadId, text)
    setReplyDraft(threadId, '')
  }

  const submitNewThread = (): void => {
    const text = replyText('')
    if (text === '' || sheetId === null) return
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const range = workbook?.getActiveRange()
    const worksheet = workbook?.getActiveSheet()
    if (!workbook || !range || !worksheet || worksheet.getSheetId() !== sheetId) {
      setReplyDraft('', '')
      return
    }
    addThread(sheetId, range.getRow(), range.getColumn(), text)
    setReplyDraft('', '')
  }

  // A ribbon-created thread gets its reply input focused, so typing goes to
  // the new thread without a second click; a plain New Comment focuses the
  // bottom composer instead.
  useEffect(() => {
    if (draftThread === null || draftThread.id === lastDraftIdRef.current) return
    lastDraftIdRef.current = draftThread.id
    draftInputRef.current?.focus()
  }, [draftThread])
  useEffect(() => {
    if (composerFocusSignal === lastComposerSignalRef.current) return
    lastComposerSignalRef.current = composerFocusSignal
    composerInputRef.current?.focus()
  }, [composerFocusSignal])

  return (
    <section
      className="thread-comments"
      role="dialog"
      aria-label={t('appThreadComments')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') {
          event.stopPropagation()
          onClose()
        }
      }}
    >
      <header>
        <span className="thread-title">{t('appThreadComments')}</span>
        <button
          className="thread-close"
          data-tip={t('appClose')}
          aria-label={t('appClose')}
          onClick={onClose}
        >
          ✕
        </button>
      </header>
      {threads.length === 0 ? (
        <p className="thread-empty">{t('appThreadEmpty')}</p>
      ) : (
        <ul className="thread-list">
          {threads.map((thread) => (
            <li
              key={thread.id}
              className={thread.resolved ? 'thread-card resolved' : 'thread-card'}
            >
              <header className="thread-card-head">
                <button
                  type="button"
                  className="thread-anchor"
                  data-tip={t('appThreadGoToCell')}
                  aria-label={`${t('appThreadGoToCell')} ${formatAddress(thread.row, thread.column)}`}
                  onClick={() => onNavigate(thread)}
                >
                  {formatAddress(thread.row, thread.column)}
                </button>
                <button
                  type="button"
                  className="thread-resolve"
                  aria-pressed={thread.resolved}
                  onClick={() => {
                    if (sheetId === null) return
                    setThreadResolved(sheetId, thread.id, !thread.resolved)
                  }}
                >
                  {thread.resolved ? t('appThreadReopen') : t('appThreadResolve')}
                </button>
                <button
                  type="button"
                  className="thread-delete"
                  aria-label={t('appDeleteLabel')}
                  onClick={() => {
                    if (sheetId === null) return
                    deleteThread(sheetId, thread.id)
                  }}
                >
                  ✕
                </button>
              </header>
              <div className="thread-author">
                {thread.messages[0]?.author || t('appThreadUnknownAuthor')}
                <span className="thread-time">
                  {formatThreadTime(thread.messages[0]?.createdAt)}
                </span>
              </div>
              <p className="thread-text">{thread.messages[0]?.text}</p>
              {thread.messages.length > 1 && (
                <ul className="thread-replies">
                  {thread.messages.slice(1).map((message) => (
                    <li key={message.id} className="thread-reply">
                      <span className="thread-author">
                        {message.author || t('appThreadUnknownAuthor')}
                        <span className="thread-time">{formatThreadTime(message.createdAt)}</span>
                      </span>
                      <p className="thread-text">{message.text}</p>
                    </li>
                  ))}
                </ul>
              )}
              <div className="thread-compose">
                <input
                  type="text"
                  ref={thread.id === draftThread?.id ? draftInputRef : undefined}
                  value={replyDrafts.get(thread.id) ?? ''}
                  placeholder={t('appThreadReplyPlaceholder')}
                  aria-label={t('appThreadReplyPlaceholder')}
                  onChange={(event) => setReplyDraft(thread.id, event.target.value)}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter') {
                      event.preventDefault()
                      submitReply(thread.id)
                    }
                  }}
                />
                <button
                  type="button"
                  className="thread-post"
                  disabled={replyText(thread.id) === ''}
                  onClick={() => submitReply(thread.id)}
                >
                  {t('appThreadPost')}
                </button>
              </div>
            </li>
          ))}
        </ul>
      )}
      <div className="thread-compose new-thread">
        <input
          type="text"
          ref={composerInputRef}
          value={replyDrafts.get('') ?? ''}
          placeholder={t('appThreadNewPlaceholder')}
          aria-label={t('appThreadNewPlaceholder')}
          onChange={(event) => setReplyDraft('', event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault()
              submitNewThread()
            }
          }}
        />
        <button
          type="button"
          className="thread-post"
          disabled={replyText('') === ''}
          onClick={submitNewThread}
        >
          {t('appThreadPost')}
        </button>
      </div>
    </section>
  )
}

/// Timestamps are document data rendered as chrome: shown in the viewer's
/// locale and timezone, like Excel's Comments pane.
function formatThreadTime(iso: string | undefined): string {
  if (iso === undefined || iso === '') return ''
  const date = new Date(iso)
  if (Number.isNaN(date.getTime())) return iso
  return date.toLocaleString(undefined, {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  })
}
