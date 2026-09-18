// PERF-904: per-locale dictionary chunk, fetched on demand by loadStrings()
// (../strings.ts) before the first render or a language switch commits.
import { de as app } from '../app/de'
import { de as ribbon } from '../ribbon/de'
import { de as references } from '../references/de'
import { tableStrings } from '../strings-table'
import { de as editor } from '../editor/de'
import { de as review } from '../review/de'
import { de as ai } from '../ai/de'
import { de as layout } from '../layout/de'

export default {
  ...app,
  ...ribbon,
  ...references,
  ...tableStrings.de,
  ...editor,
  ...review,
  ...ai,
  ...layout,
}
