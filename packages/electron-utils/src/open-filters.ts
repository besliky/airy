/**
 * Extension groups for the suite-wide File > Open dialog: every editor menu
 * offers the same per-type filters plus an "all supported" combined filter,
 * mirroring the Home screen's browse dialog. The extension lists live here
 * (single source of truth); each app localizes the filter names from its own
 * dictionary. Legacy .doc/.ppt binaries are included so they are selectable
 * and surface the explicit "not supported" dialog instead of being grayed
 * out.
 */
export const OPEN_EXTENSION_GROUPS = {
  word: ['docx', 'doc'],
  excel: ['xlsx', 'xlsm', 'xls', 'csv'],
  ppt: ['pptx', 'ppt'],
  pdf: ['pdf'],
  markdown: ['md', 'markdown'],
  html: ['html', 'htm'],
} as const

export const ALL_OPEN_EXTENSIONS: readonly string[] = [
  ...new Set(Object.values(OPEN_EXTENSION_GROUPS).flat()),
]
