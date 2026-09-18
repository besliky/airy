// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { id as app } from '../app/id'
import { id as ribbon } from '../ribbon/id'
import { id as references } from '../references/id'
import { tableStrings } from '../strings-table'
import { id as editor } from '../editor/id'
import { id as review } from '../review/id'
import { id as ai } from '../ai/id'
import { id as layout } from '../layout/id'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.id,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
