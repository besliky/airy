// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { ko as app } from '../app/ko'
import { ko as dialogs } from '../dialogs/ko'
import { ko as ai } from '../ai/ko'

export default { ...app, ...dialogs, ...ai }
