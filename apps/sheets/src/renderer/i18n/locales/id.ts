// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { id as app } from '../app/id'
import { id as dialogs } from '../dialogs/id'
import { id as ai } from '../ai/id'

export default { ...app, ...dialogs, ...ai }
