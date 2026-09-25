import type { LineNumberSettings, ParsedDoc, SectionInfo, SectionSettings, DocGrid } from './types'

/** US Letter, portrait, 1-inch margins */
export const DEFAULT_SECTION: SectionSettings = {
  pageWidth: 12240,
  pageHeight: 15840,
  orientation: 'portrait',
  marginTop: 1440,
  marginRight: 1440,
  marginBottom: 1440,
  marginLeft: 1440,
  pageBorder: false,
  columns: 1,
  headerDist: 720,
  footerDist: 720,
}

/** Vertical alignment of page content (sectPr w:vAlign); top/default returns undefined */
function vAlignOf(xml: string): 'center' | 'both' | 'bottom' | undefined {
  const v = /<w:vAlign w:val="(center|both|bottom)"\s*\/>/.exec(xml)?.[1]
  return v as 'center' | 'both' | 'bottom' | undefined
}

function intAttr(tag: string, name: string, fallback: number): number {
  const m = new RegExp(`${name}="(-?\\d+)"`).exec(tag)
  const v = m ? parseInt(m[1], 10) : NaN
  return Number.isFinite(v) ? v : fallback
}

/**
 * True when w:pgBorders declares at least one visible side. Word writes explicit
 * "no border" as side elements with w:val="none" (or "nil"), so the mere presence
 * of w:pgBorders must not be treated as a drawn border.
 */
function hasVisiblePageBorder(xml: string): boolean {
  const pgBorders = /<w:pgBorders[^>]*\/>|<w:pgBorders[\s\S]*?<\/w:pgBorders>/.exec(xml)?.[0]
  if (!pgBorders) return false
  const sides = pgBorders.match(/<w:(?:top|left|bottom|right)\b[^>]*\/?>/g) ?? []
  return sides.some((side) => {
    const val = /w:val="([^"]*)"/.exec(side)?.[1]
    return val !== undefined && val !== 'none' && val !== 'nil'
  })
}

/** Styling details of a visible w:pgBorders box (undefined when no visible side). */
function pageBorderPropsOf(xml: string): SectionSettings['pageBorderProps'] {
  const pgBorders = /<w:pgBorders[^>]*\/>|<w:pgBorders[\s\S]*?<\/w:pgBorders>/.exec(xml)?.[0]
  if (!pgBorders || !hasVisiblePageBorder(xml)) return undefined
  const display = /w:display="(firstPage|notFirstPage)"/.exec(pgBorders)?.[1] as
    'firstPage' | 'notFirstPage' | undefined
  const offsetFrom = /w:offsetFrom="(page|text)"/.exec(pgBorders)?.[1] as
    'page' | 'text' | undefined
  let spacePt = 0
  let szEighths = 0
  let color: string | undefined
  const sides: NonNullable<SectionSettings['pageBorderProps']>['sides'] = {}
  for (const side of pgBorders.match(/<w:(?:top|left|bottom|right)\b[^>]*\/?>/g) ?? []) {
    const val = /w:val="([^"]*)"/.exec(side)?.[1]
    if (!val || val === 'none' || val === 'nil') continue
    const name = /<w:(top|left|bottom|right)\b/.exec(side)![1] as keyof typeof sides
    const sideColor = /w:color="([0-9A-Fa-f]{6})"/.exec(side)?.[1]
    sides[name] = {
      val,
      widthPt: intAttr(side, 'w:sz', 0) / 8,
      spacePt: intAttr(side, 'w:space', 0),
      ...(sideColor ? { color: sideColor } : {}),
    }
    spacePt = Math.max(spacePt, intAttr(side, 'w:space', 0))
    szEighths = Math.max(szEighths, intAttr(side, 'w:sz', 0))
    color ??= sideColor
  }
  return {
    ...(display ? { display } : {}),
    ...(offsetFrom ? { offsetFrom } : {}),
    spacePt,
    widthPt: szEighths / 8,
    ...(color ? { color } : {}),
    sides,
  }
}

/** Page setup from one w:sectPr XML slice. */
export function sectionSettingsFromXml(xml: string): SectionSettings {
  const pgSz = /<w:pgSz[^>]*\/?>/.exec(xml)?.[0] ?? ''
  const pgMar = /<w:pgMar[^>]*\/?>/.exec(xml)?.[0] ?? ''

  // Parse docGrid (w:docGrid)
  let docGrid: DocGrid | undefined
  const docGridTag = /<w:docGrid[^>]*\/?>/.exec(xml)?.[0]
  if (docGridTag) {
    const typeMatch = /w:type="([^"]+)"/.exec(docGridTag)
    const linePitchMatch = /w:linePitch="(\d+)"/.exec(docGridTag)
    const charSpaceMatch = /w:charSpace="(-?\d+)"/.exec(docGridTag)
    const gridType = (typeMatch?.[1] ?? 'default') as DocGrid['type']
    const validTypes: DocGrid['type'][] = ['default', 'lines', 'linesAndChars', 'snapToChars']
    docGrid = {
      type: validTypes.includes(gridType) ? gridType : 'default',
      ...(linePitchMatch ? { linePitch: parseInt(linePitchMatch[1], 10) } : {}),
      ...(charSpaceMatch ? { charSpace: parseInt(charSpaceMatch[1], 10) } : {}),
    }
  }

  // explicit unequal column widths (w:cols > w:col children); w:sep only
  // exists in the expanded form, so both read the paired element
  const colsElement = /<w:cols[^>]*>[\s\S]*?<\/w:cols>/.exec(xml)?.[0]
  const colWidths = (colsElement?.match(/<w:col [^>]*w:w="\d+"[^>]*\/>/g) ?? [])
    .map((tag) => intAttr(tag, 'w:w', 0))
    .filter((w) => w > 0)
  const columnSep = colsElement !== undefined && /<w:sep\b/.test(colsElement)

  const pageBorderProps = pageBorderPropsOf(xml)
  const lnNum = lineNumbersOf(xml)

  return {
    pageWidth: intAttr(pgSz, 'w:w', DEFAULT_SECTION.pageWidth),
    pageHeight: intAttr(pgSz, 'w:h', DEFAULT_SECTION.pageHeight),
    orientation: pgSz.includes('w:orient="landscape"') ? 'landscape' : 'portrait',
    marginTop: intAttr(pgMar, 'w:top', DEFAULT_SECTION.marginTop),
    marginRight: intAttr(pgMar, 'w:right', DEFAULT_SECTION.marginRight),
    marginBottom: intAttr(pgMar, 'w:bottom', DEFAULT_SECTION.marginBottom),
    marginLeft: intAttr(pgMar, 'w:left', DEFAULT_SECTION.marginLeft),
    headerDist: intAttr(pgMar, 'w:header', 720),
    footerDist: intAttr(pgMar, 'w:footer', 720),
    ...(vAlignOf(xml) ? { vAlign: vAlignOf(xml) } : {}),
    pageBorder: hasVisiblePageBorder(xml),
    ...(pageBorderProps ? { pageBorderProps } : {}),
    columns: intAttr(/<w:cols[^>]*\/?>/.exec(xml)?.[0] ?? '', 'w:num', 1),
    colSpace: intAttr(/<w:cols[^>]*\/?>/.exec(xml)?.[0] ?? '', 'w:space', 720),
    ...(colWidths.length >= 2 ? { colWidths } : {}),
    ...(columnSep ? { columnSep: true } : {}),
    ...(/<w:bidi\s*\/>/.test(xml) ? { bidi: true } : {}),
    ...(docGrid ? { docGrid } : {}),
    ...(textDirectionOf(xml) ? { textDirection: textDirectionOf(xml) } : {}),
    ...(lnNum ? { lineNumbers: lnNum } : {}),
  }
}

function textDirectionOf(xml: string): string | undefined {
  const val = /<w:textDirection[^>]*w:val="([^"]+)"/.exec(xml)?.[1]
  return val && val !== 'lrTb' ? val : undefined
}

/**
 * Line numbering (sectPr w:lnNumType): countBy/start/restart/distance, all
 * attributes optional (Word defaults: count 1, start 1, restart newPage).
 */
function lineNumbersOf(xml: string): LineNumberSettings | undefined {
  const tag = /<w:lnNumType[^>]*\/?>/.exec(xml)?.[0]
  if (!tag) return undefined
  const num = (name: string): number | undefined => {
    const m = new RegExp(`${name}="(\\d+)"`).exec(tag)
    return m ? parseInt(m[1], 10) : undefined
  }
  const restart = /w:restart="(newPage|newSection|continuous)"/.exec(tag)?.[1] as
    'newPage' | 'newSection' | 'continuous' | undefined
  const countBy = num('w:count')
  const start = num('w:start')
  const distance = num('w:distance')
  return {
    ...(countBy !== undefined ? { countBy } : {}),
    ...(start !== undefined ? { start } : {}),
    ...(restart ? { restart } : {}),
    ...(distance !== undefined ? { distance } : {}),
  }
}

/** Read page setup from the trailing (hidden) w:sectPr. */
export function readSectionSettings(parsed: ParsedDoc): SectionSettings {
  const sectBlock = parsed.blocks.find((b) => b.hidden && b.originalXml?.includes('<w:sectPr'))
  return sectionSettingsFromXml(sectBlock?.originalXml ?? '')
}

const SECT_PR_RE = /<w:sectPr[^>]*\/>|<w:sectPr[\s\S]*?<\/w:sectPr>/

/** On/off flag element (w:titlePg, w:evenAndOddHeaders, …) present and not switched
 *  off: matches self-closing and paired forms, w:val="0|false|off" counts as off. */
export function xmlFlagOn(xml: string, tag: string): boolean {
  for (const m of xml.matchAll(new RegExp(`<${tag}(?=[\\s/>])[^>]*>`, 'g'))) {
    const val = /w:val="([^"]*)"/.exec(m[0])?.[1]
    if (val === undefined || !/^(?:0|false|off)$/.test(val)) return true
  }
  return false
}

function hfRefs(
  xml: string,
  kind: 'header' | 'footer',
): Partial<Record<'default' | 'first' | 'even', string>> {
  const refs: Partial<Record<'default' | 'first' | 'even', string>> = {}
  for (const ref of xml.match(new RegExp(`<w:${kind}Reference[^>]*/>`, 'g')) ?? []) {
    const type = /w:type="(default|first|even)"/.exec(ref)?.[1] ?? 'default'
    const rId = /r:id="([^"]+)"/.exec(ref)?.[1]
    if (rId) refs[type as 'default' | 'first' | 'even'] = rId
  }
  return refs
}

function sectionFromSectPr(
  sectPrXml: string,
  firstBlockIndex: number,
  lastBlockIndex: number,
): SectionInfo {
  const type = /<w:type[^>]*w:val="(nextPage|continuous|evenPage|oddPage|nextColumn)"/.exec(
    sectPrXml,
  )?.[1]
  const pgNumStart = /<w:pgNumType[^>]*w:start="(\d+)"/.exec(sectPrXml)?.[1]
  const pgNumFmt = /<w:pgNumType[^>]*w:fmt="([^"]+)"/.exec(sectPrXml)?.[1]
  return {
    settings: sectionSettingsFromXml(sectPrXml),
    startType: (type as SectionInfo['startType']) ?? 'nextPage',
    firstBlockIndex,
    lastBlockIndex,
    sectPrXml,
    titlePg: xmlFlagOn(sectPrXml, 'w:titlePg'),
    ...(pgNumStart !== undefined ? { pageNumberStart: parseInt(pgNumStart, 10) } : {}),
    ...(pgNumFmt !== undefined ? { pageNumberFmt: pgNumFmt } : {}),
    headerRefs: hfRefs(sectPrXml, 'header'),
    footerRefs: hfRefs(sectPrXml, 'footer'),
  }
}

/**
 * Rewrite the sectPr page numbering w:pgNumType (fmt = number format, start = starting
 * page number; undefined fields are omitted, and the tag is removed when both are unset).
 * Schema order: pgNumType comes after pgMar/pgBorders and before cols/docGrid.
 */
export function applyPageNumType(
  sectPrXml: string,
  fmt: string | undefined,
  start: number | undefined,
): string {
  let xml = sectPrXml.replace(/<w:pgNumType[^>]*\/>/, '')
  if (fmt === undefined && start === undefined) return xml
  const tag = `<w:pgNumType${fmt !== undefined ? ` w:fmt="${fmt}"` : ''}${start !== undefined ? ` w:start="${start}"` : ''}/>`
  if (/<w:cols[\s/>]/.test(xml)) xml = xml.replace(/(<w:cols[\s/>])/, `${tag}$1`)
  else if (/<w:docGrid/.test(xml)) xml = xml.replace(/(<w:docGrid)/, `${tag}$1`)
  else xml = xml.replace(/<\/w:sectPr>/, `${tag}</w:sectPr>`)
  return xml
}

/**
 * Enumerate every section in document order. A paragraph-level w:sectPr
 * (section-break paragraph) closes its section; the trailing hidden sectPr closes the last.
 * Block ranges use docxIndex (body child order), boundaries inclusive.
 *
 * A document whose body-level trailing sectPr is missing (the OOXML schema makes
 * it optional; hand-crafted/repaired files omit it) leaves the blocks after the
 * last paragraph-level sectPr without an owning section. Word treats the final
 * section's missing properties as defaults (portrait Letter) with header/footer
 * references inherited from the previous section, so the same implicit final
 * section is modeled here — the tail renders as its own page instead of being
 * absorbed into the previous section (BUG-1708).
 */
export function readSections(parsed: ParsedDoc): SectionInfo[] {
  const sections: SectionInfo[] = []
  let first = 0
  let last = 0
  for (const block of parsed.blocks) {
    if (block.docxIndex != null) last = Math.max(last, block.docxIndex)
    const xml = block.originalXml ?? ''
    if (block.docxIndex == null || !xml.includes('<w:sectPr')) continue
    const sectPrXml = SECT_PR_RE.exec(xml)?.[0]
    if (!sectPrXml) continue
    sections.push(sectionFromSectPr(sectPrXml, first, block.docxIndex))
    first = block.docxIndex + 1
  }
  if (sections.length === 0) {
    sections.push(sectionFromSectPr('', 0, last))
    return sections
  }
  const lastSection = sections[sections.length - 1]
  if (lastSection.lastBlockIndex < last) {
    // empty sectPr slice: default settings, nextPage start, and no own hf refs
    // (absence = linked to previous, resolved by inheritance downstream)
    sections.push(sectionFromSectPr('', lastSection.lastBlockIndex + 1, last))
  }
  return sections
}

/** Rewrite pgSz / pgMar inside a w:sectPr XML slice, preserving everything else. */
export function applySectionSettings(sectPrXml: string, settings: SectionSettings): string {
  const orient = settings.orientation === 'landscape' ? ' w:orient="landscape"' : ''
  const pgSz = `<w:pgSz w:w="${settings.pageWidth}" w:h="${settings.pageHeight}"${orient}/>`
  let xml = sectPrXml
  if (/<w:pgSz[^>]*\/?>/.test(xml)) {
    xml = xml.replace(/<w:pgSz[^>]*\/?>/, pgSz)
  } else {
    xml = xml.replace(/(<w:sectPr[^>]*>)/, `$1${pgSz}`)
  }
  const replaceMarAttr = (tag: string, name: string, value: number): string => {
    if (new RegExp(`${name}="`).test(tag)) {
      return tag.replace(new RegExp(`${name}="-?\\d+"`), `${name}="${value}"`)
    }
    return tag.replace(/\/>$/, ` ${name}="${value}"/>`)
  }
  const marMatch = /<w:pgMar[^>]*\/>/.exec(xml)
  if (marMatch) {
    let tag = marMatch[0]
    tag = replaceMarAttr(tag, 'w:top', settings.marginTop)
    tag = replaceMarAttr(tag, 'w:right', settings.marginRight)
    tag = replaceMarAttr(tag, 'w:bottom', settings.marginBottom)
    tag = replaceMarAttr(tag, 'w:left', settings.marginLeft)
    if (settings.headerDist !== undefined)
      tag = replaceMarAttr(tag, 'w:header', settings.headerDist)
    if (settings.footerDist !== undefined)
      tag = replaceMarAttr(tag, 'w:footer', settings.footerDist)
    xml = xml.replace(marMatch[0], tag)
  } else {
    const pgMar =
      `<w:pgMar w:top="${settings.marginTop}" w:right="${settings.marginRight}"` +
      ` w:bottom="${settings.marginBottom}" w:left="${settings.marginLeft}"` +
      ' w:header="708" w:footer="708" w:gutter="0"/>'
    xml = xml.replace(/<\/w:sectPr>/, `${pgMar}</w:sectPr>`)
  }

  // page border: single-line box measured from the page edge (pgBorders sits
  // between pgMar and cols in CT_SectPr)
  xml = xml.replace(/<w:pgBorders[^>]*\/>|<w:pgBorders[\s\S]*?<\/w:pgBorders>/, '')
  if (settings.pageBorder) {
    const side = (name: string) =>
      `<w:${name} w:val="single" w:sz="4" w:space="24" w:color="auto"/>`
    const pgBorders =
      '<w:pgBorders w:offsetFrom="page">' +
      `${side('top')}${side('left')}${side('bottom')}${side('right')}` +
      '</w:pgBorders>'
    xml = xml.replace(/(<w:pgMar[^>]*\/>)/, `$1${pgBorders}`)
  }

  // line numbering (w:lnNumType between pgBorders and pgNumType in CT_SectPr):
  // undefined = numbering off, the tag is dropped; attrs are omitted when unset.
  // Insert before the first present later-schema element, else after the last
  // present earlier-schema one (pgBorders > paperSrc > pgMar; applySectionSettings
  // guarantees pgSz/pgMar exist by this point).
  xml = xml.replace(/<w:lnNumType[^>]*\/>/, '')
  if (settings.lineNumbers) {
    const ln = settings.lineNumbers
    const tag =
      '<w:lnNumType' +
      (ln.countBy !== undefined ? ` w:count="${ln.countBy}"` : '') +
      (ln.start !== undefined ? ` w:start="${ln.start}"` : '') +
      (ln.restart !== undefined ? ` w:restart="${ln.restart}"` : '') +
      (ln.distance !== undefined ? ` w:distance="${ln.distance}"` : '') +
      '/>'
    const later = [
      'pgNumType',
      'cols',
      'formProt',
      'vAlign',
      'noEndnote',
      'titlePg',
      'textDirection',
      'bidi',
      'rtlGutter',
      'docGrid',
      'printerSettings',
    ].find((name) => new RegExp(`<w:${name}[\\s/>]`).test(xml))
    if (later) {
      xml = xml.replace(new RegExp(`(<w:${later}[\\s/>])`), `${tag}$1`)
    } else {
      const anchor =
        /<w:pgBorders[^>]*\/>|<w:pgBorders[\s\S]*?<\/w:pgBorders>/.exec(xml)?.[0] ??
        /<w:paperSrc[^>]*\/?>/.exec(xml)?.[0] ??
        /<w:pgMar[^>]*\/>/.exec(xml)![0]
      xml = xml.replace(anchor, `${anchor}${tag}`)
    }
  }

  // columns (w:cols may be self-closing or carry explicit <w:col> children);
  // w:sep ("line between columns") only exists in the expanded form
  const colsMatch = /<w:cols[^>]*\/>|<w:cols[^>]*>[\s\S]*?<\/w:cols>/.exec(xml)
  const numAttr = settings.columns > 1 ? ` w:num="${settings.columns}"` : ''
  const wantSep = settings.columnSep === true
  const hasSep = colsMatch !== null && /<w:sep\b/.test(colsMatch[0])
  const colsAnchor = (tag: string): string => {
    const anchor = /(<w:pgBorders[\s\S]*?<\/w:pgBorders>|<w:pgMar[^>]*\/>)/.exec(xml)
    if (anchor) return xml.replace(anchor[0], `${anchor[0]}${tag}`)
    return xml.replace(/<\/w:sectPr>/, `${tag}</w:sectPr>`)
  }
  if (
    settings.colWidths !== undefined &&
    settings.columns > 1 &&
    settings.colWidths.length === settings.columns
  ) {
    // explicit unequal widths: rebuild the element (opt-in via colWidths) —
    // unless the document already carries exactly these values (round-trip;
    // the separator state participates in the unchanged check)
    const currentWidths = (colsMatch?.[0].match(/<w:col [^>]*w:w="\d+"[^>]*\/>/g) ?? []).map((t) =>
      intAttr(t, 'w:w', 0),
    )
    const unchanged =
      colsMatch !== null &&
      hasSep === wantSep &&
      intAttr(colsMatch[0], 'w:num', 1) === settings.columns &&
      currentWidths.length === settings.colWidths.length &&
      currentWidths.every((w, i) => w === settings.colWidths![i])
    if (!unchanged) {
      const space = settings.colSpace ?? 720
      const children = settings.colWidths
        .map((w, i) =>
          i < settings.colWidths!.length - 1
            ? `<w:col w:w="${w}" w:space="${space}"/>`
            : `<w:col w:w="${w}"/>`,
        )
        .join('')
      const sep = wantSep ? '<w:sep/>' : ''
      const tag = `<w:cols${numAttr} w:space="${space}" w:equalWidth="0">${sep}${children}</w:cols>`
      xml = colsMatch ? xml.replace(colsMatch[0], tag) : colsAnchor(tag)
    }
  } else if (colsMatch) {
    const openTag = /^<w:cols[^>]*>/.exec(colsMatch[0])?.[0] ?? colsMatch[0]
    const selfClosing = colsMatch[0].endsWith('/>')
    const currentNum = / w:num="(\d+)"/.exec(openTag)?.[1] ?? '1'
    // w:col children survive an expansion/collapse triggered by w:sep, but
    // only while the count is unchanged — a different count invalidates them
    const existingChildren =
      currentNum === String(settings.columns)
        ? (colsMatch[0].match(/<w:col\b[^>]*\/?>/g) ?? []).join('')
        : ''
    if (selfClosing || currentNum !== String(settings.columns) || hasSep !== wantSep) {
      // explicit per-column widths only stay valid while the count is unchanged
      let tag = openTag.replace(/ w:num="\d+"/, '').replace(/\/?>$/, '/>')
      if (numAttr) tag = tag.replace(/^<w:cols/, `<w:cols${numAttr}`)
      // honor an explicitly different column gap (720 = OOXML default); equal
      // values leave the tag byte-identical for round-trip safety
      if (settings.colSpace !== undefined && settings.colSpace !== intAttr(tag, 'w:space', 720)) {
        tag = / w:space="\d+"/.test(tag)
          ? tag.replace(/ w:space="\d+"/, ` w:space="${settings.colSpace}"`)
          : tag.replace(/\/>$/, ` w:space="${settings.colSpace}"/>`)
      }
      if (wantSep || existingChildren) {
        // expanded form: w:sep first, then any w:col children (CT_Columns order)
        tag = tag.replace(/\/>$/, `>${wantSep ? '<w:sep/>' : ''}${existingChildren}</w:cols>`)
      }
      xml = xml.replace(colsMatch[0], tag)
    }
  } else if (numAttr) {
    const space = settings.colSpace ?? 425
    xml = colsAnchor(
      wantSep
        ? `<w:cols${numAttr} w:space="${space}"><w:sep/></w:cols>`
        : `<w:cols${numAttr} w:space="${space}"/>`,
    )
  }

  // section direction (w:bidi, after cols in CT_SectPr): undefined = keep the
  // document's tag untouched; true/false = ensure present/absent
  if (settings.bidi !== undefined) {
    const hasBidi = /<w:bidi\s*\/>/.test(xml)
    if (settings.bidi && !hasBidi) {
      if (/<w:docGrid/.test(xml)) xml = xml.replace(/(<w:docGrid)/, '<w:bidi/>$1')
      else xml = xml.replace(/<\/w:sectPr>/, '<w:bidi/></w:sectPr>')
    } else if (!settings.bidi && hasBidi) {
      xml = xml.replace(/<w:bidi\s*\/>/, '')
    }
  }
  return xml
}

/**
 * Rewrite the sectPr start type w:type (how this section starts relative to the previous
 * one). nextPage is the default and is written by removing the tag; w:type comes before
 * pgSz in CT_SectPr.
 */
export function applySectionStartType(
  sectPrXml: string,
  type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn',
): string {
  let xml = sectPrXml.replace(/<w:type[^>]*\/>/, '')
  if (type === 'nextPage') return xml
  const tag = `<w:type w:val="${type}"/>`
  if (/<w:pgSz/.test(xml)) xml = xml.replace(/(<w:pgSz)/, `${tag}$1`)
  else xml = xml.replace(/(<w:sectPr[^>]*>)/, `$1${tag}`)
  return xml
}

/** Read the page color (w:background) from document.xml; null when unset. */
export function readPageColor(parsed: ParsedDoc): string | null {
  const m = /<w:background[^>]*w:color="([0-9A-Fa-f]{6})"/.exec(parsed.internal.documentXml)
  return m ? m[1].toUpperCase() : null
}
