// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { cs as app } from '../app/cs'
import { cs as dialogs } from '../dialogs/cs'
import { cs as ai } from '../ai/cs'

export default { ...app, ...dialogs, ...ai }
