// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { fr as app } from '../app/fr'
import { fr as dialogs } from '../dialogs/fr'
import { fr as ai } from '../ai/fr'

export default { ...app, ...dialogs, ...ai }
