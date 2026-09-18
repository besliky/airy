// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { hi as app } from '../app/hi'
import { hi as dialogs } from '../dialogs/hi'
import { hi as ai } from '../ai/hi'

export default { ...app, ...dialogs, ...ai }
