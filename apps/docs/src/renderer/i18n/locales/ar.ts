// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { ar as app } from '../app/ar'
import { ar as ribbon } from '../ribbon/ar'
import { ar as references } from '../references/ar'
import { tableStrings } from '../strings-table'
import { ar as editor } from '../editor/ar'
import { ar as review } from '../review/ar'
import { ar as ai } from '../ai/ar'
import { ar as layout } from '../layout/ar'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.ar,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
