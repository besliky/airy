// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { he as app } from '../app/he'
import { he as ribbon } from '../ribbon/he'
import { he as references } from '../references/he'
import { tableStrings } from '../strings-table'
import { he as editor } from '../editor/he'
import { he as review } from '../review/he'
import { he as ai } from '../ai/he'
import { he as layout } from '../layout/he'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.he,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
