// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { zh as app } from '../app/zh'
import { zh as ribbon } from '../ribbon/zh'
import { zh as references } from '../references/zh'
import { tableStrings } from '../strings-table'
import { zh as editor } from '../editor/zh'
import { zh as review } from '../review/zh'
import { zh as ai } from '../ai/zh'
import { zh as layout } from '../layout/zh'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.zh,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
