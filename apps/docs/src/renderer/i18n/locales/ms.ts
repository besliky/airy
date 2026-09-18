// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { ms as app } from '../app/ms'
import { ms as ribbon } from '../ribbon/ms'
import { ms as references } from '../references/ms'
import { tableStrings } from '../strings-table'
import { ms as editor } from '../editor/ms'
import { ms as review } from '../review/ms'
import { ms as ai } from '../ai/ms'
import { ms as layout } from '../layout/ms'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.ms,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
