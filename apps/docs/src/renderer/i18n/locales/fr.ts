// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { fr as app } from '../app/fr'
import { fr as ribbon } from '../ribbon/fr'
import { fr as references } from '../references/fr'
import { tableStrings } from '../strings-table'
import { fr as editor } from '../editor/fr'
import { fr as review } from '../review/fr'
import { fr as ai } from '../ai/fr'
import { fr as layout } from '../layout/fr'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.fr,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
