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
import { mkdtemp, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { extname, join } from 'node:path'

import { docToText } from '@airy-office/file-parse'

import { DocxSession, type SessionOrigin } from '../docx/session.js'
import { resolveConfined } from '../docx/paths.js'
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
] as const

export type OpenedDocument = DocxSession | XlsxSession | TextSession

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

const DOC_TEXT_FALLBACK_WARNING =
  'Legacy .doc opened read-only: text extraction without formatting structure. ' +
  'Full editing (converted .docx session) requires LibreOffice.'

/** Open a document of any supported format inside the workspace root. */
export async function openDocument(rawPath: string, root?: string): Promise<OpenedDocument> {
  // confinement is checked up front for every format so a path outside the
  // root always reports the confinement error, whatever its extension
  resolveConfined(rawPath, root)
  const ext = extensionOf(rawPath)
  switch (ext) {
    case 'docx':
      return DocxSession.open(rawPath, root)
    case 'xlsx':
    case 'xlsm':
    case 'xls':
    case 'ods':
      return XlsxSession.open(rawPath, root)
    case 'doc': {
      const tool = await findSoffice()
      if (!tool) {
        return TextSession.open(rawPath, root, {
          format: 'doc',
          warning: DOC_TEXT_FALLBACK_WARNING,
          extract: (bytes) => docToText(bytes),
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
