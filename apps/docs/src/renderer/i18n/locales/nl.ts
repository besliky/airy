// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { nl as app } from '../app/nl'
import { nl as ribbon } from '../ribbon/nl'
import { nl as references } from '../references/nl'
import { tableStrings } from '../strings-table'
import { nl as editor } from '../editor/nl'
import { nl as review } from '../review/nl'
import { nl as ai } from '../ai/nl'
import { nl as layout } from '../layout/nl'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.nl,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
