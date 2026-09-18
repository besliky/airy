// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { th as app } from '../app/th'
import { th as dialogs } from '../dialogs/th'
import { th as ai } from '../ai/th'

export default { ...app, ...dialogs, ...ai }
