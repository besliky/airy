// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { hi as app } from '../app/hi'
import { hi as ribbon } from '../ribbon/hi'
import { hi as references } from '../references/hi'
import { tableStrings } from '../strings-table'
import { hi as editor } from '../editor/hi'
import { hi as review } from '../review/hi'
import { hi as ai } from '../ai/hi'
import { hi as layout } from '../layout/hi'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.hi,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
