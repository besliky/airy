// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { he as app } from '../app/he'
import { he as dialogs } from '../dialogs/he'
import { he as ai } from '../ai/he'

export default { ...app, ...dialogs, ...ai }
