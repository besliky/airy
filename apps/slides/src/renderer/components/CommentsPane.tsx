/**
 * Comments pane (right side) — current page's comment threads + create/reply/
 * resolve/delete. The source of truth lives in the main process (the pptx
 * comments + commentsExtended parts); this only displays and sends back intents.
 * The list is owned by App (refreshed after page switch/add/delete), keeping
 * the ribbon badge in sync.
 */
import React, { useEffect, useRef, useState } from 'react'
import type { SlideComment } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { IconComment, IconSidebarCollapse } from './icons'

interface Props {
  slideIndex: number
  comments: SlideComment[]
  /** Focus the input when it changes (the "new comment" entry increments it per click) */
  focusNonce?: number
  onAdd: (text: string) => void
  onReply: (text: string, parent: SlideComment) => void
  /** resolve/reopen a set of comments (typically a whole thread) as one undo step */
  onResolve: (refs: Array<{ authorId: number; idx: number }>, done: boolean) => void
  onDelete: (c: SlideComment) => void
  onCollapse: () => void
}

/** dt (ISO) → local short time; returned as-is on parse failure. */
function fmtDt(dt: string, locale: string): string {
  const d = new Date(dt)
  if (Number.isNaN(d.getTime())) return dt
  return d.toLocaleString(locale, {
    month: 'numeric',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  })
}

const sameRef = (a: SlideComment, b: SlideComment) =>
  a.parentId?.authorId === b.authorId && a.parentId?.idx === b.idx

/** group the flat list into threads: parents in order, replies nested under their parent */
export function commentThreads(
  comments: SlideComment[],
): Array<{ head: SlideComment; replies: SlideComment[] }> {
  const threads: Array<{ head: SlideComment; replies: SlideComment[] }> = []
  for (const c of comments) {
    if (c.parentId) continue
    threads.push({ head: c, replies: comments.filter((r) => r !== c && sameRef(r, c)) })
  }
  return threads
}

function CommentCard({
  c,
  onDelete,
  compact,
}: {
  c: SlideComment
  onDelete: (c: SlideComment) => void
  compact?: boolean
}) {
  const { t, dateLocale } = useI18n()
  return (
    <div className={`comment-card${compact ? ' comment-reply' : ''}`}>
      <div className="comment-head">
        <span className="comment-avatar">{(c.initials || c.author).slice(0, 2)}</span>
        <span className="comment-meta">
          <span className="comment-author">{c.author}</span>
          {c.dt && <span className="comment-time">{fmtDt(c.dt, dateLocale)}</span>}
        </span>
        <button
          className="comment-del"
          data-tip={t('paneCommentsDelete')}
          aria-label={t('paneCommentsDelete')}
          onClick={() => onDelete(c)}
        >
          ✕
        </button>
      </div>
      <div className="comment-text">{c.text}</div>
    </div>
  )
}

export function CommentsPane({
  slideIndex,
  comments,
  focusNonce,
  onAdd,
  onReply,
  onResolve,
  onDelete,
  onCollapse,
}: Props) {
  const { t } = useI18n()
  const [draft, setDraft] = useState('')
  const inputRef = useRef<HTMLTextAreaElement>(null)
  /** thread (head ref) whose inline reply composer is open */
  const [replyTo, setReplyTo] = useState<SlideComment | null>(null)
  const [replyDraft, setReplyDraft] = useState('')
  /** resolved threads render collapsed; expanded ones override while open */
  const [expanded, setExpanded] = useState<Set<string>>(new Set())
  const replyRef = useRef<HTMLTextAreaElement>(null)

  useEffect(() => {
    if (focusNonce) inputRef.current?.focus()
  }, [focusNonce])

  useEffect(() => {
    if (replyTo) replyRef.current?.focus()
  }, [replyTo])

  const submit = () => {
    const text = draft.trim()
    if (!text) return
    onAdd(text)
    setDraft('')
  }

  const submitReply = () => {
    const text = replyDraft.trim()
    if (!text || !replyTo) return
    onReply(text, replyTo)
    setReplyDraft('')
    setReplyTo(null)
  }

  const threads = commentThreads(comments)

  return (
    <aside className="comments-pane">
      <div className="ai-panel-header">
        <span className="ai-panel-title">{t('paneCommentsTitle', { n: slideIndex + 1 })}</span>
        <div className="ai-panel-header-actions">
          <button
            className="ai-header-btn"
            onClick={onCollapse}
            data-tip={t('paneCommentsCollapse')}
            aria-label={t('paneCommentsCollapse')}
          >
            <IconSidebarCollapse size={15} />
          </button>
        </div>
      </div>

      <div className="comments-list">
        {comments.length === 0 && (
          <div className="comments-empty">
            <IconComment size={22} />
            <p>{t('paneCommentsEmpty')}</p>
            <p className="comments-empty-sub">{t('paneCommentsEmptySub')}</p>
          </div>
        )}
        {threads.map(({ head, replies }) => {
          const key = `${head.authorId}-${head.idx}`
          const resolved = head.resolved || replies.some((r) => r.resolved)
          const collapsed = resolved && !expanded.has(key)
          const threadRefs = [
            { authorId: head.authorId, idx: head.idx },
            ...replies.map((r) => ({ authorId: r.authorId, idx: r.idx })),
          ]
          return (
            <div key={key} className={`comment-thread${resolved ? ' resolved' : ''}`}>
              {collapsed ? (
                <button
                  className="comment-thread-collapsed"
                  onClick={() => setExpanded((prev) => new Set(prev).add(key))}
                >
                  <span className="comment-thread-check" aria-hidden="true">
                    ✓
                  </span>
                  <span className="comment-thread-summary">
                    {head.text}
                    {replies.length > 0 &&
                      ` · ${t('paneCommentsReplies', { n: String(replies.length) })}`}
                  </span>
                </button>
              ) : (
                <>
                  <CommentCard c={head} onDelete={onDelete} />
                  {replies.map((r) => (
                    <CommentCard key={`${r.authorId}-${r.idx}`} c={r} onDelete={onDelete} compact />
                  ))}
                  <div className="comment-thread-footer">
                    <button
                      className="comment-action"
                      onClick={() => {
                        setReplyTo(head)
                        setReplyDraft('')
                      }}
                    >
                      {t('paneCommentsReply')}
                    </button>
                    <button
                      className="comment-action"
                      onClick={() => {
                        onResolve(threadRefs, !resolved)
                        // reopening expands the thread again
                        if (resolved) {
                          setExpanded((prev) => {
                            const next = new Set(prev)
                            next.delete(key)
                            return next
                          })
                        }
                      }}
                    >
                      {resolved ? t('paneCommentsReopen') : t('paneCommentsResolve')}
                    </button>
                    {resolved && (
                      <span className="comment-resolved-badge">{t('paneCommentsResolved')}</span>
                    )}
                  </div>
                  {replyTo === head && (
                    <div className="comment-reply-box">
                      <textarea
                        ref={replyRef}
                        rows={2}
                        value={replyDraft}
                        placeholder={t('paneCommentsReplyPlaceholder')}
                        onChange={(e) => setReplyDraft(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
                            e.preventDefault()
                            submitReply()
                          }
                          if (e.key === 'Escape') {
                            e.preventDefault()
                            setReplyTo(null)
                          }
                        }}
                      />
                      <div className="comment-reply-actions">
                        <button className="btn-ghost" onClick={() => setReplyTo(null)}>
                          {t('paneCancel')}
                        </button>
                        <button
                          className="btn-primary"
                          disabled={!replyDraft.trim()}
                          onClick={submitReply}
                        >
                          {t('paneCommentsPost')}
                        </button>
                      </div>
                    </div>
                  )}
                </>
              )}
            </div>
          )
        })}
      </div>

      <div className="comment-new">
        <textarea
          ref={inputRef}
          rows={3}
          value={draft}
          placeholder={t('paneCommentsPlaceholder')}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
              e.preventDefault()
              submit()
            }
          }}
        />
        <button className="btn-primary" disabled={!draft.trim()} onClick={submit}>
          {t('paneCommentsPost')}
        </button>
      </div>
    </aside>
  )
}
