// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { pt as app } from '../app/pt'
import { pt as ribbon } from '../ribbon/pt'
import { pt as references } from '../references/pt'
import { tableStrings } from '../strings-table'
import { pt as editor } from '../editor/pt'
import { pt as review } from '../review/pt'
import { pt as ai } from '../ai/pt'
import { pt as layout } from '../layout/pt'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.pt,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
