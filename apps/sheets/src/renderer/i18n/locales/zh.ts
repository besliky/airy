// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { zh as app } from '../app/zh'
import { zh as dialogs } from '../dialogs/zh'
import { zh as ai } from '../ai/zh'

export default { ...app, ...dialogs, ...ai }
