// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { ja as app } from '../app/ja'
import { ja as ribbon } from '../ribbon/ja'
import { ja as references } from '../references/ja'
import { tableStrings } from '../strings-table'
import { ja as editor } from '../editor/ja'
import { ja as review } from '../review/ja'
import { ja as ai } from '../ai/ja'
import { ja as layout } from '../layout/ja'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.ja,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
