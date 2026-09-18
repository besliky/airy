// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { th as app } from '../app/th'
import { th as ribbon } from '../ribbon/th'
import { th as references } from '../references/th'
import { tableStrings } from '../strings-table'
import { th as editor } from '../editor/th'
import { th as review } from '../review/th'
import { th as ai } from '../ai/th'
import { th as layout } from '../layout/th'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.th,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
