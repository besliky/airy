// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { it as app } from '../app/it'
import { it as dialogs } from '../dialogs/it'
import { it as ai } from '../ai/it'

export default { ...app, ...dialogs, ...ai }
