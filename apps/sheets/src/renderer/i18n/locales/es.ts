// PERF-901: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { es as app } from '../app/es'
import { es as dialogs } from '../dialogs/es'
import { es as ai } from '../ai/es'

export default { ...app, ...dialogs, ...ai }
