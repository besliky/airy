// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { zhTW as app } from '../app/zh-TW'
import { zhTW as ribbon } from '../ribbon/zh-TW'
import { zhTW as references } from '../references/zh-TW'
import { tableStrings } from '../strings-table'
import { zhTW as editor } from '../editor/zh-TW'
import { zhTW as review } from '../review/zh-TW'
import { zhTW as ai } from '../ai/zh-TW'
import { zhTW as layout } from '../layout/zh-TW'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings['zh-TW'],
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
