import { existsSync } from 'node:fs'
import { basename, join } from 'node:path'

/** Basename of a generated file with path separators, invalid file-name
    characters and control bytes flattened (control characters are
    intentionally rejected from generated file names). */
function sanitizeGeneratedFileName(name: string): string {
  return (
    basename(name)
      // eslint-disable-next-line no-control-regex
      .replace(/[/\\:*?"<>|\u0000-\u001f]/g, '_')
      .trim()
  )
}

/** Pick a safe, unused PDF path inside the configured Airy save directory. */
export function uniqueGeneratedPdfPath(
  dir: string,
  suggestedName: string,
  pathExists: (path: string) => boolean = existsSync,
): string {
  let fileName = sanitizeGeneratedFileName(String(suggestedName || 'merged.pdf'))
  if (!fileName || fileName === '.' || fileName === '..') fileName = 'merged.pdf'
  if (!/\.pdf$/i.test(fileName)) fileName += '.pdf'

  const stem = fileName.slice(0, -4)
  let candidate = join(dir, fileName)
  for (let i = 2; pathExists(candidate); i++) candidate = join(dir, `${stem}-${i}.pdf`)
  return candidate
}

/** Pick a safe, unused path inside the configured Airy save directory for an
    extracted embedded file (PDF portfolio attachment, UX-1734): the
    attachment's own name stays recognizable, a plausible extension (1-8
    alphanumeric characters) is preserved, anything absurd is dropped. */
export function uniqueGeneratedAttachmentPath(
  dir: string,
  suggestedName: string,
  pathExists: (path: string) => boolean = existsSync,
): string {
  const clean = sanitizeGeneratedFileName(String(suggestedName || ''))
  const dot = clean.lastIndexOf('.')
  const hasExt =
    dot > 0 && dot < clean.length - 1 && /^[A-Za-z0-9]{1,8}$/.test(clean.slice(dot + 1))
  const ext = hasExt ? clean.slice(dot) : ''
  const stem = (hasExt ? clean.slice(0, dot) : clean) || 'attachment'
  let candidate = join(dir, stem + ext)
  for (let i = 2; pathExists(candidate); i++) candidate = join(dir, `${stem}-${i}${ext}`)
  return candidate
}
