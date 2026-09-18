// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { es as app } from '../app/es'
import { es as ribbon } from '../ribbon/es'
import { es as references } from '../references/es'
import { tableStrings } from '../strings-table'
import { es as editor } from '../editor/es'
import { es as review } from '../review/es'
import { es as ai } from '../ai/es'
import { es as layout } from '../layout/es'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.es,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
