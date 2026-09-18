import type { zh } from './zh'

export const en = {
  layoutHyphenation: 'Hyphenation',
  layoutHyphNone: 'None',
  layoutHyphManual: 'Manual',
  layoutHyphAutomatic: 'Automatic',
  layoutHyphManualDesc: 'Insert optional hyphens where you want breaks',
  layoutHyphManualHint: 'Manual hyphenation: press {keys} to insert an optional hyphen',
  layoutHyphSet: 'Automatic hyphenation on',
  layoutHyphUnset: 'Automatic hyphenation off',
  layoutSoftHyphen: 'Optional hyphen',
  layoutColsDialogTitle: 'Columns',
  layoutColsMore: 'More Columns…',
  layoutColOne: 'One',
  layoutColTwo: 'Two',
  layoutColThree: 'Three',
  layoutColLeft: 'Left',
  layoutColRight: 'Right',
  layoutColSpacing: 'Spacing',
  layoutColWidth: 'Width',
  layoutColWidth1: 'Column 1 width',
  layoutLineBetween: 'Line between columns',
} satisfies Record<keyof typeof zh, string>
