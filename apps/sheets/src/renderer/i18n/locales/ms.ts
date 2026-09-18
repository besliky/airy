// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { ms as app } from '../app/ms'
import { ms as dialogs } from '../dialogs/ms'
import { ms as ai } from '../ai/ms'

export default { ...app, ...dialogs, ...ai }
