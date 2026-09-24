// open_document routing (S6 "xlsx + legacy"): extension-based dispatch onto
// the session kinds, per ADR-9 — native formats open directly, legacy/ODF
// formats import through conversion into the native model:
//
//   .docx            -> DocxSession (native)
//   .xlsx/.xlsm      -> XlsxSession (sidecar, native)
//   .xls/.ods        -> XlsxSession via sidecar convert_workbook (calamine;
//                       styles lost -> warning)
//   .doc             -> soffice -> temp .docx -> DocxSession with origin;
//                       without LibreOffice -> read-only TextSession
//                       (word-extractor through @airy-office/file-parse)
//   .odt             -> soffice -> temp .docx -> DocxSession with origin;
//                       without LibreOffice -> actionable error
//   .md/.markdown    -> MarkdownSession (native text session, line-based
//                       editing; PAR-003)
//   .html/.htm       -> HtmlSession (native text session, line-based editing
//                       with a parse5 structure summary; PAR-004)
//   .pptx            -> SlidesSession (native pptx-engine model; PAR-001 —
//                       legacy .ppt/.odp are refused with a conversion hint:
//                       soffice pptx round-trips lose too much to promise
//                       byte fidelity)
//   .pdf             -> read-only TextSession (pdfjs text extraction through
//                       @airy-office/file-parse; PAR-002 — pages are joined
//                       with blank lines, editing is not supported headlessly)
import { mkdtemp, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { extname, join } from 'node:path'

import { docToText, pdfToText } from '@airy-office/file-parse'

import { DocxSession, type SessionOrigin } from '../docx/session.js'
import { resolveConfined } from '../docx/paths.js'
import { HtmlSession } from '../html/session.js'
import { MarkdownSession } from '../markdown/session.js'
import {
  classifyOleContent,
  encryptedOfficeRefusal,
  OLE_SNIFF_MAX_BYTES,
  oleAsZipRefusal,
  readHead,
} from './ole.js'
import { SlidesSession } from '../slides/session.js'
import { assertWithinOpenCap } from '../sessions/size-fence.js'
import { TextSession } from '../sessions/text.js'
import { XlsxSession } from '../xlsx/session.js'
import { convertViaSoffice, findSoffice, SOFFICE_FILTERS, sofficeMissingError } from './soffice.js'

export const SUPPORTED_OPEN_EXTENSIONS = [
  'docx',
  'xlsx',
  'xlsm',
  'xls',
  'ods',
  'doc',
  'odt',
  'md',
  'markdown',
  'html',
  'htm',
  'pptx',
  'pdf',
] as const

export type OpenedDocument =
  DocxSession | XlsxSession | TextSession | MarkdownSession | HtmlSession | SlidesSession

export function extensionOf(path: string): string {
  return extname(path).replace('.', '').toLowerCase()
}

async function statOrNull(path: string): Promise<{ mtimeMs: number; size: number } | null> {
  try {
    const info = await stat(path)
    return { mtimeMs: info.mtimeMs, size: info.size }
  } catch {
    return null
  }
}

/**
 * Refuse a password-protected file before its engine (sidecar, pptx engine,
 * LibreOffice, word-extractor) sees garbage: an encrypted Office document is
 * an OLE2 container, not a plain package (BUG-1504). Head sniff (64 KiB) —
 * best-effort on this path; the docx session re-checks with the full bytes.
 * Legacy CFB documents without an encryption signal (a plain .doc/.xls) pass.
 */
async function refuseEncryptedContainer(path: string, format: string): Promise<void> {
  const head = await readHead(path, OLE_SNIFF_MAX_BYTES)
  const ole = classifyOleContent(head)
  if (ole === 'encrypted-ooxml' || ole === 'encrypted-legacy') {
    throw encryptedOfficeRefusal(path, format)
  }
}

const DOC_TEXT_FALLBACK_WARNING =
  'Legacy .doc opened read-only: text extraction without formatting structure. ' +
  'Full editing (converted .docx session) requires LibreOffice.'

const PDF_TEXT_WARNING =
  'PDF opened read-only: text extraction via pdfjs (pages are separated by blank lines; ' +
  'scanned/image-only pages extract no text). Editing PDFs headlessly is not supported.'

const LEGACY_PRESENTATION_HINT =
  'Legacy presentation formats are not supported yet. Convert the deck to .pptx first, ' +
  'e.g. "soffice --convert-to pptx file.ppt", then open the .pptx.'

/** Open a document of any supported format inside the workspace root. */
export async function openDocument(rawPath: string, root?: string): Promise<OpenedDocument> {
  // confinement is checked up front for every format so a path outside the
  // root always reports the confinement error, whatever its extension
  const path = resolveConfined(rawPath, root)
  const ext = extensionOf(rawPath)
  switch (ext) {
    case 'docx':
      return DocxSession.open(rawPath, root)
    case 'xlsx':
    case 'xlsm':
    case 'xls':
    case 'ods': {
      // one head sniff serves both container refusals
      const ole = classifyOleContent(await readHead(path, OLE_SNIFF_MAX_BYTES))
      if (ole === 'encrypted-ooxml' || ole === 'encrypted-legacy') {
        throw encryptedOfficeRefusal(path, ext)
      }
      // UX-1691: a plain (unencrypted) OLE2 container saved as .xlsx/.xlsm is
      // a renamed legacy workbook — refuse naming the class instead of the
      // sidecar's raw zip parse failure ("invalid Zip archive: Could not find
      // EOCD"). A plain OLE .xls is the legitimate conversion route and .ods
      // is out of scope here, so only the OOXML extensions refuse.
      if ((ext === 'xlsx' || ext === 'xlsm') && ole === 'plain') {
        throw oleAsZipRefusal(path, ext)
      }
      return XlsxSession.open(rawPath, root)
    }
    case 'doc': {
      const tool = await findSoffice()
      if (!tool) {
        return TextSession.open(rawPath, root, {
          format: 'doc',
          warning: DOC_TEXT_FALLBACK_WARNING,
          extract: (bytes) => {
            // word-extractor dies with "memory outside buffer bounds" on an
            // encrypted container — name password protection instead (BUG-1504)
            const ole = classifyOleContent(bytes)
            if (ole === 'encrypted-ooxml' || ole === 'encrypted-legacy') {
              throw encryptedOfficeRefusal(path, 'doc')
            }
            return docToText(bytes)
          },
        })
      }
      return openConvertedWordDocument(rawPath, root, tool, 'doc')
    }
    case 'odt': {
      const tool = await findSoffice()
      if (!tool) {
        throw sofficeMissingError(
          'Opening .odt requires LibreOffice (it is converted to .docx for editing).',
        )
      }
      return openConvertedWordDocument(rawPath, root, tool, 'odt')
    }
    case 'md':
    case 'markdown':
      return MarkdownSession.open(rawPath, root)
    case 'html':
    case 'htm':
      return HtmlSession.open(rawPath, root)
    case 'pptx':
      await refuseEncryptedContainer(path, ext)
      return SlidesSession.open(rawPath, root)
    case 'pdf':
      return TextSession.open(rawPath, root, {
        format: 'pdf',
        warning: PDF_TEXT_WARNING,
        extract: (bytes) => pdfToText(bytes),
      })
    case 'ppt':
    case 'odp':
    case 'pps':
    case 'pot':
      throw new Error(LEGACY_PRESENTATION_HINT)
    default:
      throw new Error(
        `Unsupported file type ".${ext || '(none)'}". Supported extensions: ` +
          SUPPORTED_OPEN_EXTENSIONS.map((supported) => `.${supported}`).join(', ') +
          '.',
      )
  }
}

/** .doc/.odt -> LibreOffice -> temp .docx -> docx session remembering the origin. */
async function openConvertedWordDocument(
  rawPath: string,
  root: string | undefined,
  tool: { binaryPath: string },
  format: 'doc' | 'odt',
): Promise<DocxSession> {
  const path = resolveConfined(rawPath, root)
  const stamp = await statOrNull(path)
  if (!stamp) {
    throw new Error(`Cannot read "${path}": file does not exist.`)
  }
  // SEC-1302: the raw cap runs on the ORIGINAL, before soffice is spawned —
  // the DocxSession fence below only sees the server-controlled temp .docx,
  // so without this an oversized hostile file would spend host memory/CPU in
  // the conversion subprocess first.
  assertWithinOpenCap(path, stamp.size, format)
  // BUG-1504: headless soffice cannot decrypt either — an encrypted container
  // would surface as an opaque conversion failure; name password protection.
  await refuseEncryptedContainer(path, format)
  const tempDir = await mkdtemp(join(tmpdir(), 'airy-import-'))
  try {
    const converted = await convertViaSoffice(tool, path, {
      filter: SOFFICE_FILTERS.docx,
      extension: 'docx',
      outDir: tempDir,
    })
    const origin: SessionOrigin = { path, format, stamp, tempDir }
    return await DocxSession.open(converted, root, origin)
  } catch (e) {
    await rm(tempDir, { recursive: true, force: true })
    throw new Error(
      `Cannot convert "${path}" to .docx: ${e instanceof Error ? e.message : String(e)}`,
      { cause: e },
    )
  }
}
