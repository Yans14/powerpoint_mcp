import path from "node:path";
import { z } from "zod";

const MEASUREMENT_RE = /^\d+(?:\.\d+)?(?:in|pt|cm|px)$/i;

const presentationIdSchema = z
  .string()
  .uuid()
  .describe("Presentation session UUID returned by pptx_create_presentation or pptx_open_presentation.");

const slideIndexSchema = z
  .number()
  .int()
  .min(1)
  .describe("1-based slide index. Verify current indices with pptx_get_presentation_state before mutating operations.");

const absolutePathSchema = z
  .string()
  .min(1)
  .refine((value) => path.isAbsolute(value), {
    message: "Path must be absolute.",
  });

const measurementSchema = z
  .union([
    z.number().describe("Raw EMU integer (advanced use)."),
    z
      .string()
      .regex(MEASUREMENT_RE, 'Measurement string like "2in", "24pt", "5cm", or "96px".'),
  ])
  .describe("Human-friendly measurement converted internally to EMU.");

const colorHexSchema = z
  .string()
  .regex(/^#?[0-9a-fA-F]{6}$/)
  .describe("Hex color string in RRGGBB or #RRGGBB format.");

export interface ToolDefinition {
  name: string;
  description: string;
  schema: z.AnyZodObject;
  mutating: boolean;
}

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "pptx_get_engine_info",
    description: "Return active engine mode and runtime details.",
    schema: z.object({}).strict(),
    mutating: false,
  },
  {
    name: "pptx_create_presentation",
    description:
      "Create a new in-memory session backed by a temporary working copy. Original files are never modified until pptx_save_presentation.",
    schema: z
      .object({
        width: measurementSchema.optional(),
        height: measurementSchema.optional(),
        template_path: absolutePathSchema
          .optional()
          .describe("Optional absolute .potx/.pptx template path to seed the presentation."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_open_presentation",
    description: "Open an existing .pptx into a safe temp working copy and return a new presentation_id.",
    schema: z
      .object({
        file_path: absolutePathSchema.describe("Absolute path to an existing .pptx file."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_save_presentation",
    description: "Persist the current working copy to an explicit absolute output path.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        output_path: absolutePathSchema.describe("Absolute path for saved .pptx output."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_close_presentation",
    description: "Close a presentation session and release temporary resources.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_list_open_presentations",
    description: "List all open presentation sessions and associated file paths.",
    schema: z.object({}).strict(),
    mutating: false,
  },
  {
    name: "pptx_get_presentation_state",
    description: "Return authoritative slide index, titles, layouts, and shape counts for the session.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_get_layouts",
    description:
      "Discover available layout names and placeholder types. Always call this before pptx_add_slide and never guess layout names.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_get_layout_detail",
    description: "Return full placeholder inventory for a named layout.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        layout_name: z
          .string()
          .min(1)
          .describe("Exact layout name returned by pptx_get_layouts. Case-sensitive."),
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_get_masters",
    description: "List slide masters available in the presentation.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_get_theme",
    description: "Return theme colors and major/minor font scheme.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_get_slide",
    description: "Return complete slide shape tree, placeholders, and notes text.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_add_slide",
    description:
      "Add a slide by exact layout name. Layout names must come from pptx_get_layouts. Returns updated presentation_state.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        layout_name: z
          .string()
          .min(1)
          .describe("Exact layout name returned by pptx_get_layouts. Never guess."),
        position: z.number().int().min(1).optional().describe("Optional 1-based insertion position."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_duplicate_slide",
    description: "Duplicate an existing slide and optionally place it at a target position.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        source_index: slideIndexSchema,
        target_position: z.number().int().min(1).optional(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_delete_slide",
    description:
      "Delete a slide by 1-based index. Indices above the deleted index will shift down in returned presentation_state.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_reorder_slides",
    description:
      "Reorder slides using the full desired order list of current indices, e.g. [2,1,3]. Length must equal current slide_count.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        new_order: z.array(z.number().int().min(1)).min(1),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_move_slide",
    description: "Move one slide from from_index to to_index.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        from_index: slideIndexSchema,
        to_index: slideIndexSchema,
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_set_slide_background",
    description: "Set slide background using solid color, image, or simple gradient fallback.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        color_hex: colorHexSchema.optional(),
        image_path: absolutePathSchema.optional(),
        gradient_start_color_hex: colorHexSchema.optional(),
        gradient_end_color_hex: colorHexSchema.optional(),
        gradient_angle_deg: z.number().min(0).max(360).optional(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_get_slide_snapshot",
    description:
      "Return Base64 JPEG preview for visual verification. In OOXML mode this requires LibreOffice and pdftoppm.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        width_px: z.number().int().min(128).max(4096).optional(),
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_get_placeholders",
    description: "List placeholders with names, indices, types, and geometry for a slide.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_set_placeholder_text",
    description:
      "Primary text injection tool. Targets placeholders by exact name to preserve layout/master inheritance formatting.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        placeholder_name: z.string().min(1),
        text_content: z.string(),
        alignment: z.enum(["left", "center", "right", "justify"]).optional(),
        font_name: z.string().optional(),
        font_size_pt: z.number().positive().optional(),
        bold: z.boolean().optional(),
        italic: z.boolean().optional(),
        underline: z.boolean().optional(),
        color_hex: colorHexSchema.optional(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_set_placeholder_image",
    description: "Populate a picture/content placeholder with an image from absolute local path.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        placeholder_name: z.string().min(1),
        image_path: absolutePathSchema,
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_clear_placeholder",
    description: "Clear content of a placeholder by name.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        placeholder_name: z.string().min(1),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_get_placeholder_text",
    description: "Read structured text content and formatting from a placeholder.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        placeholder_name: z.string().min(1),
      })
      .strict(),
    mutating: false,
  },
];

export const TOOL_MAP = new Map(TOOL_DEFINITIONS.map((tool) => [tool.name, tool]));
