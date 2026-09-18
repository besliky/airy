// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { cs as app } from '../app/cs'
import { cs as ribbon } from '../ribbon/cs'
import { cs as references } from '../references/cs'
import { tableStrings } from '../strings-table'
import { cs as editor } from '../editor/cs'
import { cs as review } from '../review/cs'
import { cs as ai } from '../ai/cs'
import { cs as layout } from '../layout/cs'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.cs,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
