/**
 * The Insert Function catalog, derived from the live Univer registry.
 *
 * The engine registers hundreds of executors (builtins plus the app's own
 * CELL / RATE / MINIFS / … overrides) — the same source of truth that
 * decides #NAME? on recalc (function-registry-probe). The dialog therefore
 * must not browse a hand-written list: names come from the registry's
 * executors, categories from Univer's FUNCTION_NAMES_* enums (each builtin
 * belongs to exactly one), and descriptions/signatures from the merged
 * locale's sheets-formula.functionList entries. The curated 61 entries stay
 * as an override layer: their hand-written localized descriptions and
 * exact syntax win over the derived data.
 */
import {
  FUNCTION_NAMES_ARRAY,
  FUNCTION_NAMES_COMPATIBILITY,
  FUNCTION_NAMES_CUBE,
  FUNCTION_NAMES_DATABASE,
  FUNCTION_NAMES_DATE,
  FUNCTION_NAMES_ENGINEERING,
  FUNCTION_NAMES_FINANCIAL,
  FUNCTION_NAMES_INFORMATION,
  FUNCTION_NAMES_LOGICAL,
  FUNCTION_NAMES_LOOKUP,
  FUNCTION_NAMES_MATH,
  FUNCTION_NAMES_STATISTICAL,
  FUNCTION_NAMES_TEXT,
  FUNCTION_NAMES_WEB,
  type IFunctionService,
} from '@univerjs/engine-formula'

import type { StringKey } from './i18n/locale'

export interface CatalogFunction {
  readonly name: string
  /// Stable English category id; displayed through the locale shards.
  readonly category: string
  /// Curated syntax, or `NAME(param, …)` derived from the locale's
  /// parameter order. Empty when nothing is known beyond the name.
  readonly syntax: string
  /// Curated localized-description key; wins over `description`.
  readonly descKey?: StringKey | undefined
  /// Registry/locale English description (the dialog's fallback text).
  readonly description?: string | undefined
}

export interface CuratedFunction {
  readonly name: string
  readonly category: string
  readonly syntax: string
  readonly descKey?: StringKey | undefined
}

/// Stable category ids. The curated entries and the locale shards use the
/// same ids; ids without a shard key fall back to the raw English name in
/// the dropdown (none today — every id has a key).
const CATEGORY_BY_ENUM: ReadonlyArray<readonly [Record<string, string>, string]> = [
  [FUNCTION_NAMES_MATH, 'Math'],
  [FUNCTION_NAMES_STATISTICAL, 'Statistical'],
  [FUNCTION_NAMES_LOGICAL, 'Logical'],
  [FUNCTION_NAMES_LOOKUP, 'Lookup'],
  [FUNCTION_NAMES_DATE, 'Date & Time'],
  [FUNCTION_NAMES_TEXT, 'Text'],
  [FUNCTION_NAMES_FINANCIAL, 'Financial'],
  [FUNCTION_NAMES_DATABASE, 'Database'],
  [FUNCTION_NAMES_ENGINEERING, 'Engineering'],
  [FUNCTION_NAMES_INFORMATION, 'Information'],
  [FUNCTION_NAMES_COMPATIBILITY, 'Compatibility'],
  [FUNCTION_NAMES_CUBE, 'Cube'],
  [FUNCTION_NAMES_ARRAY, 'Array'],
  [FUNCTION_NAMES_WEB, 'Web'],
]

export const FALLBACK_CATEGORY = 'More'

function categoryNameOf(name: string): string {
  for (const [names, category] of CATEGORY_BY_ENUM) {
    if (Object.values(names).includes(name)) return category
  }
  return FALLBACK_CATEGORY
}

/// The merged Univer locale tree (the object createUniver receives).
export interface FunctionListLocale {
  readonly 'sheets-formula'?: {
    readonly functionList?: Readonly<
      Record<
        string,
        {
          readonly description?: string
          readonly abstract?: string
          readonly functionParameter?: Readonly<Record<string, { readonly name?: string }>>
        }
      >
    >
  }
}

function derivedSyntax(name: string, entry: { functionParameter?: unknown } | undefined): string {
  const parameters = entry?.functionParameter
  if (parameters == null || typeof parameters !== 'object') return ''
  const names = Object.values(parameters as Record<string, { name?: string }>)
    .map((parameter) => parameter?.name)
    .filter((value): value is string => typeof value === 'string' && value.length > 0)
  if (names.length === 0) return ''
  return `${name}(${names.join(', ')})`
}

/// Builds the catalog from the live registry. Falls back to the curated
/// list alone when the engine has not registered its executors yet (the
/// dialog opens before the plugin batch lands only in tests).
export function buildFunctionCatalog(
  functionService: Pick<IFunctionService, 'getExecutors'> | null,
  locale: FunctionListLocale | null,
  curated: readonly CuratedFunction[],
): CatalogFunction[] {
  const curatedByName = new Map(curated.map((spec) => [spec.name, spec]))
  const functionList = locale?.['sheets-formula']?.functionList
  const entries: CatalogFunction[] = []
  const seen = new Set<string>()
  const push = (name: string): void => {
    if (name === '' || seen.has(name)) return
    seen.add(name)
    const override = curatedByName.get(name)
    if (override) {
      entries.push({
        name,
        category: override.category,
        syntax: override.syntax,
        descKey: override.descKey,
      })
      return
    }
    const listed = functionList?.[name]
    entries.push({
      name,
      category: categoryNameOf(name),
      syntax: derivedSyntax(name, listed),
      description: listed?.abstract || listed?.description || undefined,
    })
  }
  const executors = functionService?.getExecutors()
  if (executors && executors.size > 0) {
    for (const name of executors.keys()) {
      push(typeof name === 'string' ? name.toUpperCase() : name.toString())
    }
  } else {
    for (const spec of curated) push(spec.name)
  }
  entries.sort((left, right) => left.name.localeCompare(right.name, 'en'))
  return entries
}

/// Distinct category ids present in the catalog, in dropdown order: the
/// curated order first, then any derived-only categories alphabetically.
export function catalogCategories(
  catalog: readonly CatalogFunction[],
  curatedOrder: readonly string[],
): string[] {
  const present = new Set(catalog.map((entry) => entry.category))
  const ordered = curatedOrder.filter((category) => present.has(category))
  const extra = [...present]
    .filter((category) => !ordered.includes(category))
    .sort((left, right) => left.localeCompare(right, 'en'))
  return [...ordered, ...extra]
}
