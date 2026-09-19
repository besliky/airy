// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { ru as app } from '../app/ru'
import { ru as ribbon } from '../ribbon/ru'
import { ru as references } from '../references/ru'
import { tableStrings } from '../strings-table'
import { ru as editor } from '../editor/ru'
import { ru as review } from '../review/ru'
import { ru as ai } from '../ai/ru'
import { ru as layout } from '../layout/ru'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.ru,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
