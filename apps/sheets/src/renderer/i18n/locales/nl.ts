// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { nl as app } from '../app/nl'
import { nl as dialogs } from '../dialogs/nl'
import { nl as ai } from '../ai/nl'

export default { ...app, ...dialogs, ...ai }
