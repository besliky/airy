// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { pt as app } from '../app/pt'
import { pt as dialogs } from '../dialogs/pt'
import { pt as ai } from '../ai/pt'

export default { ...app, ...dialogs, ...ai }
