import { z } from "zod";

// Repeated on every mutating tool, so the description is kept to one short sentence.
const INCLUDE_STATE_SCHEMA = {
  type: "boolean",
  description: "Also return the per-slide presentation_state, not just slide_count. Default false.",
};

/**
 * Convert a tool's zod schema into the JSON Schema advertised over MCP.
 * Uses zod 4's native converter; `io: "input"` keeps fields with defaults optional for callers.
 */
export function toInputSchema(schema: z.ZodTypeAny, options: { includeState?: boolean } = {}): Record<string, unknown> {
  const jsonSchema = z.toJSONSchema(schema, {
    target: "draft-7",
    io: "input",
    unrepresentable: "any",
    // Drop converter noise that costs tokens in every tools/list without informing the model:
    // the UUID regex (format already says it) and JS safe-integer bounds.
    override: ({ jsonSchema: node }) => {
      if (node.format === "uuid") {
        delete node.pattern;
      }
      if (node.type === "integer") {
        if (node.maximum === Number.MAX_SAFE_INTEGER) {
          delete node.maximum;
        }
        if (node.minimum === Number.MIN_SAFE_INTEGER) {
          delete node.minimum;
        }
      }
    },
  }) as Record<string, unknown>;

  delete jsonSchema.$schema;

  if (jsonSchema.type !== "object") {
    throw new Error("Tool input schema must be a JSON Schema object.");
  }

  if (options.includeState) {
    // Handled by the server for every mutating tool, so it is added here rather than in each schema.
    jsonSchema.properties = { ...(jsonSchema.properties as Record<string, unknown>), include_state: INCLUDE_STATE_SCHEMA };
  }

  return jsonSchema;
}
