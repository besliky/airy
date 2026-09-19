// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { pl as app } from '../app/pl'
import { pl as ribbon } from '../ribbon/pl'
import { pl as references } from '../references/pl'
import { tableStrings } from '../strings-table'
import { pl as editor } from '../editor/pl'
import { pl as review } from '../review/pl'
import { pl as ai } from '../ai/pl'
import { pl as layout } from '../layout/pl'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.pl,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
