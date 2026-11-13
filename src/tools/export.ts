import { MEASUREMENT_RE, TOOL_DEFINITIONS } from "./catalog.js";
import { toInputSchema } from "./schema.js";

/**
 * JSON Schema patterns carry no flags, so a case-insensitive regex loses its `i` on conversion.
 * Python's `re` (used by the jsonschema validator) accepts an inline `(?i)` prefix instead.
 */
function restoreCaseInsensitivePatterns(node: unknown): unknown {
  if (Array.isArray(node)) {
    return node.map(restoreCaseInsensitivePatterns);
  }
  if (node === null || typeof node !== "object") {
    return node;
  }
  const entries = Object.entries(node as Record<string, unknown>).map(([key, value]) => {
    if (key === "pattern" && value === MEASUREMENT_RE.source && MEASUREMENT_RE.flags.includes("i")) {
      return [key, `(?i)${value}`];
    }
    return [key, restoreCaseInsensitivePatterns(value)];
  });
  return Object.fromEntries(entries);
}

/** Where the Python side reads the catalog from, relative to the repository root. */
export const PYTHON_CATALOG_PATH = "python/generated/tool_catalog.json";

export interface ExportedTool {
  name: string;
  description: string;
  category: string;
  mutating: boolean;
  /**
   * JSON Schema (draft 7) of the tool's input, as advertised over MCP. Refinements zod cannot
   * express in JSON Schema (absolute paths, "shape_name or shape_id") are enforced by the
   * TypeScript layer and the engines, not by this schema.
   */
  input_schema: Record<string, unknown>;
}

export function exportToolCatalog(): ExportedTool[] {
  return TOOL_DEFINITIONS.map((definition) => ({
    name: definition.name,
    description: definition.description,
    category: definition.category,
    mutating: definition.mutating,
    input_schema: restoreCaseInsensitivePatterns(toInputSchema(definition.schema)) as Record<string, unknown>,
  }));
}

/** File content written to PYTHON_CATALOG_PATH; stable formatting keeps diffs reviewable. */
export function renderToolCatalog(): string {
  const header = {
    $comment: "Generated from src/tools/catalog.ts by `npm run gen:catalog`. Do not edit by hand.",
    tools: exportToolCatalog(),
  };
  return `${JSON.stringify(header, null, 2)}\n`;
}
