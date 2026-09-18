// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { ja as app } from '../app/ja'
import { ja as dialogs } from '../dialogs/ja'
import { ja as ai } from '../ai/ja'

export default { ...app, ...dialogs, ...ai }
