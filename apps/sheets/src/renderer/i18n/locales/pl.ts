// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { pl as app } from '../app/pl'
import { pl as dialogs } from '../dialogs/pl'
import { pl as ai } from '../ai/pl'

export default { ...app, ...dialogs, ...ai }
