// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { zhTW as app } from '../app/zh-TW'
import { zhTW as dialogs } from '../dialogs/zh-TW'
import { zhTW as ai } from '../ai/zh-TW'

export default { ...app, ...dialogs, ...ai }
