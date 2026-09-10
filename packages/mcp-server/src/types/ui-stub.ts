// Type-only stub for @genoffice/ui, which the sheets gateway's shared
// desktop-api references (type import of AiPanelPrefs). The MCP server never
// runs UI code — esbuild strips the type import entirely — but tsc follows it
// and would otherwise drag the React/DOM type surface into this Node package.
// The stub keeps the typecheck hermetic; see tsconfig "paths".
export type AiPanelPrefs = Record<string, unknown>
