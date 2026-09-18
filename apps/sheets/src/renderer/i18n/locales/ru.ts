// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { ru as app } from '../app/ru'
import { ru as dialogs } from '../dialogs/ru'
import { ru as ai } from '../ai/ru'

export default { ...app, ...dialogs, ...ai }
