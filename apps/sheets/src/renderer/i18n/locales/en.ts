// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { en as app } from '../app/en'
import { en as dialogs } from '../dialogs/en'
import { en as ai } from '../ai/en'

export default { ...app, ...dialogs, ...ai }
