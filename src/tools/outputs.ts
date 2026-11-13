import { z } from "zod";

/**
 * Output schemas for the read tools whose result shape is the same in both engines.
 * A tool listed here advertises `outputSchema` and returns `structuredContent`.
 *
 * Objects are loose on purpose: engines may add fields, and a client validating the result
 * against the schema must not reject it for that.
 */

const slideSummary = z.looseObject({
  index: z.number().int().describe("1-based slide index."),
  title: z.string(),
  layout: z.string().describe("Name of the slide's layout."),
  shape_count: z.number().int(),
});

const placeholder = z.looseObject({
  name: z.string(),
  idx: z.number().int(),
  type: z.string(),
});

export const OUTPUT_SCHEMAS: Record<string, z.ZodTypeAny> = {
  pptx_list_open_presentations: z.looseObject({
    presentations: z.array(
      z.looseObject({
        presentation_id: z.string(),
        original_path: z.string().describe("File the session was opened from; empty for a new deck."),
        engine: z.string(),
        dirty: z.boolean().describe("True when there are changes not yet written by pptx_save_presentation."),
        opened_at: z.string(),
      }),
    ),
  }),
  pptx_get_presentation_state: z.looseObject({
    presentation_id: z.string(),
    slide_count: z.number().int(),
    slides: z.array(slideSummary),
  }),
  pptx_get_layouts: z.looseObject({
    presentation_id: z.string(),
    layouts: z.array(
      z.looseObject({
        index: z.number().int(),
        name: z.string().describe("Pass as layout_name to pptx_add_slide."),
        placeholder_types: z.array(z.string()),
      }),
    ),
  }),
  pptx_get_layout_detail: z.looseObject({
    presentation_id: z.string(),
    layout_name: z.string(),
    placeholder_count: z.number().int(),
    placeholders: z.array(placeholder),
  }),
  pptx_get_placeholders: z.looseObject({
    presentation_id: z.string(),
    slide_index: z.number().int(),
    placeholders: z.array(placeholder),
  }),
};

export function toOutputSchema(schema: z.ZodTypeAny): Record<string, unknown> {
  const jsonSchema = z.toJSONSchema(schema, { target: "draft-7", io: "output", unrepresentable: "any" }) as Record<
    string,
    unknown
  >;
  delete jsonSchema.$schema;
  return jsonSchema;
}
