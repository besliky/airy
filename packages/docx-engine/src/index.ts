export * from './types'
export { deobfuscateOdttf, isSfnt, parseFontTable, readEmbeddedFonts } from './font-table'
export { decodeEntities } from './parse-xml-text'
export { parseDocx, styleRunFormat, type ParseExtras } from './parse'
export {
  saveDocx,
  findChartWorkbookPath,
  readDocxPartBase64,
  type SaveBlock,
  type SaveOptions,
  type StyleUpsert,
  type ParsedDocFull,
} from './patch'
export {
  TABLE_HEADER_FILL,
  applyImageWrap,
  applyImageZOrder,
  buildAnchoredTextboxParagraphXml,
  buildShapeParagraphXml,
  buildTextboxParagraphXml,
  buildWordArtParagraphXml,
  type AnchoredTextboxOptions,
  type TextboxContentParagraph,
  bookmarkIdOf,
  generateCaptionXml,
  generateIndexFieldXml,
  generateParagraphXml,
  generateTableModelXml,
  generateTableXml,
  generateTocFieldXml,
  buildTocInstruction,
  tocEntryHidesPage,
  mergePPrFormat,
  setPPrChange,
  stripPPrChange,
  patchFieldParagraphXml,
  patchImageParagraphXml,
  patchMathTokens,
  patchTableCellTexts,
  patchTextboxHeights,
  patchTextboxParas,
  patchTextboxSizes,
  patchShapeStyles,
  type ShapeStylePatch,
  patchDrawingExtent,
  patchDrawingDocPr,
  type DrawingAltPatch,
  patchTableAltText,
  type TableAltPatch,
  shadowEffectLstXml,
  buildLineParagraphXml,
  LINE_KINDS,
  type TextboxSizePatch,
  type CellParaPatch,
  type CellTextsPatch,
  type FieldTextPatch,
  type GenerateContext,
  type ImagePatch,
  type TextboxParaPatch,
  type TextboxParasPatchSet,
  type TableGenOptions,
  type TocEntry,
  type TocFieldOptions,
} from './generate'
export { parseTocInstruction } from './parse-fields'
export {
  buildChartPartXml,
  buildChartWorkbookXlsxBase64,
  patchChartWorkbookXlsxBase64,
  parseChartPartXml,
  patchChartPartXml,
  lumHex,
  CHART_WORKBOOK_REL_TYPE,
  type ChartPatch,
  type ChartSeriesPatch,
} from './chart'
export {
  latexToOmml,
  mathParagraphXml,
  mathTokensOf,
  ommlFragmentsOf,
  ommlToLatex,
  ommlToMathML,
} from './math'
export { scanBody, type BodyElement, type BodyScan } from './scan'
export { CorruptXmlError } from './xml-utils'
export { buildDiagramDisplay, diagramDefaultExtentEmu, freshDiagramGuid } from './smartart-diagram'
export { SMARTART_PRESET_PARTS } from './smartart-vendor'
export {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  type BlankDocxOptions,
  type CustomNumberingLevel,
} from './blank'
export {
  DEFAULT_SECTION,
  applySectionSettings,
  applyPageNumType,
  applySectionStartType,
  readPageColor,
  readSections,
  readSectionSettings,
  sectionSettingsFromXml,
} from './section'
export { nextNoteId, parseNotesXml, type NoteKind } from './notes'
export { readWatermarkText } from './watermark'
export {
  INK_NAME_PREFIX,
  anchoredInkRunXml,
  findInkRuns,
  injectInkRunsIntoParagraph,
  stripInkRuns,
} from './ink'
export { bibliographyLine, citationText, parseSourcesXml } from './sources'
export { readThemeColors, readThemeFonts } from './theme'
export { hashProtectionPassword, verifyProtectionPassword } from './protection'
export { decodeSymbolChar, decodeSymbolText, isSymbolFont, toSymbolPua } from './symbol-fonts'
export {
  bulletMarkerScale,
  computeListMarkerInfos,
  computeListMarkers,
  customEnumItems,
  formatNumber,
  markerTabAdvance,
  type ListItemRef,
  type ListMarkerInfo,
} from './list-markers'
export {
  FORMULA_DIV_ZERO_ERROR,
  FORMULA_EMPTY_ERROR,
  FORMULA_SYNTAX_ERROR,
  collectDirectionOperands,
  evaluateFormulaInGrid,
  formatNumericPicture,
  parseCellNumber,
  parseFormulaInstruction,
  proposeTableFormula,
  type FormulaDirection,
  type FormulaGrid,
  type FormulaGridAnchors,
  type FormulaGridTexts,
} from './table-formulas'
