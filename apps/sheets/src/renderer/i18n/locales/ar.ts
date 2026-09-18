// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { ar as app } from '../app/ar'
import { ar as dialogs } from '../dialogs/ar'
import { ar as ai } from '../ai/ar'

export default { ...app, ...dialogs, ...ai }
