// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { en as app } from '../app/en'
import { en as ribbon } from '../ribbon/en'
import { en as references } from '../references/en'
import { tableStrings } from '../strings-table'
import { en as editor } from '../editor/en'
import { en as review } from '../review/en'
import { en as ai } from '../ai/en'
import { en as layout } from '../layout/en'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.en,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
