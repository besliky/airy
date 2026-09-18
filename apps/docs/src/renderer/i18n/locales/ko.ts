// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { ko as app } from '../app/ko'
import { ko as ribbon } from '../ribbon/ko'
import { ko as references } from '../references/ko'
import { tableStrings } from '../strings-table'
import { ko as editor } from '../editor/ko'
import { ko as review } from '../review/ko'
import { ko as ai } from '../ai/ko'
import { ko as layout } from '../layout/ko'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.ko,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
