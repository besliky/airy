// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { de as app } from '../app/de'
import { de as dialogs } from '../dialogs/de'
import { de as ai } from '../ai/de'

export default { ...app, ...dialogs, ...ai }
