/// Threaded-comment writer: replaces a worksheet's full threaded-comment set —
/// the `xl/threadedComments/threadedCommentN.xml` part, the workbook-level
/// `xl/persons/person.xml` author directory, the worksheet rels, the workbook
/// rels, and the [Content_Types].xml entries. This is the modern Excel comment
/// model ([MS-XLSX] section 2.3.7): one part per sheet, replies linked to
/// their root through parentId, and resolution tracked by the root's `done`
/// flag. Legacy notes (see xlsx-notes.ts) are a separate, coexisting part.
/// Modern Excel also writes a legacy shadow comment per thread; that shadow is
/// optional for readers, so the minimal valid set written here is the
/// threadedComments part plus the persons part.

export class ThreadedCommentEditError extends Error {}

/// One resolved message as stored on a thread (save-side model).
export interface ThreadedMessage {
  readonly id: string
  readonly personId: string
  readonly author: string
  readonly text: string
  /// ISO-8601 creation timestamp, written verbatim into the `dT` attribute.
  readonly createdAt: string
  readonly userId?: string | undefined
  readonly providerId?: string | undefined
}

/// One thread: the root message's anchor cell plus its reply chain (root
/// first). Resolution is a thread-level flag in the UI model and lands as
/// `done="1"` on the root message.
export interface ThreadedThread {
  readonly id: string
  readonly row: number
  readonly column: number
  readonly resolved: boolean
  readonly messages: readonly ThreadedMessage[]
}

interface MutableThreadedPackage {
  paths(): Promise<readonly string[]>
  has(path: string): Promise<boolean>
  readText(path: string): Promise<string>
  write(path: string, content: string): void
  add(path: string, content: string): void
  remove(path: string): void
}

const THREADED_COMMENTS_REL_TYPE =
  'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment'
const PERSONS_REL_TYPE = 'http://schemas.microsoft.com/office/2017/10/relationships/person'
const THREADED_COMMENTS_CONTENT_TYPE = 'application/vnd.ms-excel.threadedComments+xml'
const PERSONS_CONTENT_TYPE = 'application/vnd.ms-excel.person+xml'
const CONTENT_TYPES_PATH = '[Content_Types].xml'
const WORKBOOK_RELS_PATH = 'xl/_rels/workbook.xml.rels'
const PERSONS_PART_PATH = 'xl/persons/person.xml'

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function columnName(index: number): string {
  let label = ''
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1) {
    label = String.fromCharCode(65 + (i % 26)) + label
  }
  return label
}

function cellRef(row: number, column: number): string {
  return `${columnName(column)}${row + 1}`
}

function worksheetRelsPath(worksheetPath: string): string {
  return worksheetPath.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels')
}

function relTarget(relsXml: string, type: string): string | null {
  const pattern = new RegExp(`<Relationship\\b[^>]*Type="${type}"[^>]*/?>`)
  const found = pattern.exec(relsXml)
  if (!found) return null
  const target = / Target="([^"]*)"/.exec(found[0])
  return target?.[1] ?? null
}

/// "../threadedComments/threadedComment1.xml" → package path.
function resolveRelTarget(basePath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const base = basePath.split('/').slice(0, -1)
  for (const part of target.split('/')) {
    if (part === '..') base.pop()
    else if (part !== '.') base.push(part)
  }
  return base.join('/')
}

function nextFreeRid(relsXml: string): string {
  const ids = [...relsXml.matchAll(/ Id="rId(\d+)"/g)].map((match) => Number(match[1]))
  return `rId${ids.length === 0 ? 1 : Math.max(...ids) + 1}`
}

async function nextFreePath(
  pkg: MutableThreadedPackage,
  template: (index: number) => string,
): Promise<string> {
  for (let index = 1; index < 10_000; index += 1) {
    const candidate = template(index)
    if (!(await pkg.has(candidate))) return candidate
  }
  throw new ThreadedCommentEditError('No free part name for the threaded comments part.')
}

function appendRel(relsXml: string, id: string, type: string, target: string): string {
  return relsXml.replace(
    '</Relationships>',
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`,
  )
}

function removeRel(relsXml: string, type: string): string {
  return relsXml.replace(new RegExp(`<Relationship\\b[^>]*Type="${type}"[^>]*/?>`), '')
}

function ensureContentTypeOverride(contentTypes: string, partName: string, type: string): string {
  if (contentTypes.includes(`PartName="/${partName}"`)) return contentTypes
  return contentTypes.replace(
    '</Types>',
    `<Override PartName="/${partName}" ContentType="${type}"/></Types>`,
  )
}

const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '</Relationships>'

interface PersonRecord {
  readonly id: string
  readonly displayName: string
  readonly userId?: string | undefined
  readonly providerId?: string | undefined
}

/// Serializes one sheet's threads. Roots carry ref/done, replies carry
/// parentId; attributes follow the CT_ThreadedComment schema order.
export function buildThreadedCommentsXml(threads: readonly ThreadedThread[]): string {
  const comments = threads
    .map((thread) =>
      thread.messages
        .map((message, index) => {
          const attributes = [`dT="${escapeXml(message.createdAt)}"`]
          if (index === 0) {
            attributes.push(`ref="${cellRef(thread.row, thread.column)}"`)
          }
          attributes.push(`personId="${escapeXml(message.personId)}"`)
          attributes.push(`id="${escapeXml(message.id)}"`)
          if (index > 0) {
            attributes.push(`parentId="${escapeXml(thread.messages[0]!.id)}"`)
          }
          if (index === 0 && thread.resolved) {
            attributes.push('done="1"')
          }
          return (
            `<threadedComment ${attributes.join(' ')}>` +
            `<text>${escapeXml(message.text)}</text></threadedComment>`
          )
        })
        .join(''),
    )
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<threadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"' +
    ' xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `${comments}</threadedComments>`
  )
}

/// Serializes the workbook-level author directory. persons without a
/// displayName still need an entry: Excel resolves unknown persons to the raw
/// id, but an absent entry makes the thread unattributed.
export function buildPersonsXml(persons: Iterable<PersonRecord>): string {
  const people = [...persons]
    .sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0))
    .map((person) => {
      const attributes = [`id="${escapeXml(person.id)}"`]
      attributes.push(`displayName="${escapeXml(person.displayName)}"`)
      if (person.userId !== undefined) attributes.push(`userId="${escapeXml(person.userId)}"`)
      attributes.push(`providerId="${escapeXml(person.providerId ?? 'None')}"`)
      return `<person ${attributes.join(' ')}/>`
    })
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<persons xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
    `${people}</persons>`
  )
}

/// Reads the existing persons directory so a rewrite preserves entries this
/// save does not know about (threads on sheets left untouched).
function parseExistingPersons(xml: string): PersonRecord[] {
  const people: PersonRecord[] = []
  for (const match of xml.matchAll(/<person\b[^>]*\/>/g)) {
    const element = match[0]
    const id = /\bid="([^"]*)"/.exec(element)?.[1]
    if (id === undefined) continue
    const decode = (raw: string | undefined): string | undefined =>
      raw
        ?.replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&amp;/g, '&')
    const userId = decode(/\suserId="([^"]*)"/.exec(element)?.[1])
    const providerId = decode(/\sproviderId="([^"]*)"/.exec(element)?.[1])
    people.push({
      id,
      displayName: decode(/displayName="([^"]*)"/.exec(element)?.[1]) ?? '',
      ...(userId === undefined ? {} : { userId }),
      ...(providerId === undefined ? {} : { providerId }),
    })
  }
  return people
}

/// Replaces the package's whole threaded-comment set (empty list removes it).
/// `entries` must cover every sheet whose threads changed; `allWorksheetPaths`
/// drives the orphaned-persons cleanup, since a persons part is only removable
/// when no worksheet references a threadedComments part anymore.
export async function applyThreadedCommentStates(
  pkg: MutableThreadedPackage,
  entries: readonly {
    readonly worksheetPath: string
    readonly threads: readonly ThreadedThread[]
  }[],
  allWorksheetPaths: readonly string[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (entries.length === 0) return

  // Persons mentioned by this save (existing entries are preserved below).
  const mentioned = new Map<string, PersonRecord>()
  for (const entry of entries) {
    for (const thread of entry.threads) {
      for (const message of thread.messages) {
        if (!mentioned.has(message.personId)) {
          mentioned.set(message.personId, {
            id: message.personId,
            displayName: message.author,
            ...(message.userId === undefined ? {} : { userId: message.userId }),
            ...(message.providerId === undefined ? {} : { providerId: message.providerId }),
          })
        }
      }
    }
  }

  for (const entry of entries) {
    const relsPath = worksheetRelsPath(entry.worksheetPath)
    const hasRels = await pkg.has(relsPath)
    let relsXml = hasRels ? await pkg.readText(relsPath) : EMPTY_RELS
    const existingTarget = relTarget(relsXml, THREADED_COMMENTS_REL_TYPE)
    const existingPath =
      existingTarget === null ? null : resolveRelTarget(entry.worksheetPath, existingTarget)
    let relsChanged = false

    if (entry.threads.length === 0) {
      if (existingPath !== null) {
        pkg.remove(existingPath)
        touchedEntries.add(existingPath)
        relsXml = removeRel(relsXml, THREADED_COMMENTS_REL_TYPE)
        relsChanged = true
        const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
        const stripped = contentTypes.replace(
          new RegExp(
            `<Override\\b[^>]*PartName="/${existingPath.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`,
          ),
          '',
        )
        if (stripped !== contentTypes) {
          pkg.write(CONTENT_TYPES_PATH, stripped)
          touchedEntries.add(CONTENT_TYPES_PATH)
        }
      }
    } else {
      const xml = buildThreadedCommentsXml(entry.threads)
      let partPath = existingPath
      if (partPath === null) {
        partPath = await nextFreePath(
          pkg,
          (index) => `xl/threadedComments/threadedComment${index}.xml`,
        )
        const rid = nextFreeRid(relsXml)
        relsXml = appendRel(
          relsXml,
          rid,
          THREADED_COMMENTS_REL_TYPE,
          `../threadedComments/${partPath.split('/').pop()}`,
        )
        relsChanged = true
        pkg.add(partPath, xml)
      } else {
        pkg.write(partPath, xml)
      }
      touchedEntries.add(partPath)
      const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
      const updated = ensureContentTypeOverride(
        contentTypes,
        partPath,
        THREADED_COMMENTS_CONTENT_TYPE,
      )
      if (updated !== contentTypes) {
        pkg.write(CONTENT_TYPES_PATH, updated)
        touchedEntries.add(CONTENT_TYPES_PATH)
      }
    }

    if (relsChanged) {
      if (hasRels) pkg.write(relsPath, relsXml)
      else pkg.add(relsPath, relsXml)
      touchedEntries.add(relsPath)
    }
  }

  // Workbook-level persons directory. Existing entries survive so threads on
  // untouched sheets keep their authors; this save's mentions win on id clash.
  const existingPersonsPath = WORKBOOK_RELS_PATH
  const hasWorkbookRels = await pkg.has(existingPersonsPath)
  let workbookRelsXml = hasWorkbookRels ? await pkg.readText(existingPersonsPath) : EMPTY_RELS
  const existingPersonsTarget = relTarget(workbookRelsXml, PERSONS_REL_TYPE)
  const existingPersonsPart =
    existingPersonsTarget === null
      ? null
      : resolveRelTarget('xl/workbook.xml', existingPersonsTarget)
  const hasAnyThreadedParts = await packageHasThreadedParts(pkg, allWorksheetPaths)
  let workbookRelsChanged = false

  if (hasAnyThreadedParts) {
    const merged = new Map<string, PersonRecord>()
    if (existingPersonsPart !== null && (await pkg.has(existingPersonsPart))) {
      for (const person of parseExistingPersons(await pkg.readText(existingPersonsPart))) {
        merged.set(person.id, person)
      }
    }
    for (const [id, person] of mentioned) merged.set(id, person)
    const personsXml = buildPersonsXml(merged.values())
    let personsPath = existingPersonsPart
    if (personsPath === null) {
      personsPath = PERSONS_PART_PATH
      workbookRelsXml = appendRel(
        workbookRelsXml,
        nextFreeRid(workbookRelsXml),
        PERSONS_REL_TYPE,
        'persons/person.xml',
      )
      workbookRelsChanged = true
      pkg.add(personsPath, personsXml)
    } else {
      pkg.write(personsPath, personsXml)
    }
    touchedEntries.add(personsPath)
    const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
    const updated = ensureContentTypeOverride(contentTypes, personsPath, PERSONS_CONTENT_TYPE)
    if (updated !== contentTypes) {
      pkg.write(CONTENT_TYPES_PATH, updated)
      touchedEntries.add(CONTENT_TYPES_PATH)
    }
  } else if (existingPersonsPart !== null && (await pkg.has(existingPersonsPart))) {
    // No threaded comments remain anywhere: the persons directory would be an
    // unreferenced author list, so it goes with them.
    pkg.remove(existingPersonsPart)
    touchedEntries.add(existingPersonsPart)
    workbookRelsXml = removeRel(workbookRelsXml, PERSONS_REL_TYPE)
    workbookRelsChanged = true
    const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
    const stripped = contentTypes.replace(
      new RegExp(
        `<Override\\b[^>]*PartName="/${existingPersonsPart.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`,
      ),
      '',
    )
    if (stripped !== contentTypes) {
      pkg.write(CONTENT_TYPES_PATH, stripped)
      touchedEntries.add(CONTENT_TYPES_PATH)
    }
  }

  if (workbookRelsChanged) {
    if (hasWorkbookRels) pkg.write(WORKBOOK_RELS_PATH, workbookRelsXml)
    else pkg.add(WORKBOOK_RELS_PATH, workbookRelsXml)
    touchedEntries.add(WORKBOOK_RELS_PATH)
  }
}

/// True when any worksheet's rels still target an existing threadedComments
/// part. Entries processed above already reflect their post-write state.
async function packageHasThreadedParts(
  pkg: MutableThreadedPackage,
  worksheetPaths: readonly string[],
): Promise<boolean> {
  for (const worksheetPath of worksheetPaths) {
    const relsPath = worksheetRelsPath(worksheetPath)
    if (!(await pkg.has(relsPath))) continue
    const target = relTarget(await pkg.readText(relsPath), THREADED_COMMENTS_REL_TYPE)
    if (target === null) continue
    if (await pkg.has(resolveRelTarget(worksheetPath, target))) return true
  }
  return false
}
