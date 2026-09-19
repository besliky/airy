// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { it as app } from '../app/it'
import { it as ribbon } from '../ribbon/it'
import { it as references } from '../references/it'
import { tableStrings } from '../strings-table'
import { it as editor } from '../editor/it'
import { it as review } from '../review/it'
import { it as ai } from '../ai/it'
import { it as layout } from '../layout/it'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.it,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
