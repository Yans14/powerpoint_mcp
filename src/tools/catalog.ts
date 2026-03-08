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
  schema: z.ZodTypeAny;
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

  // --- Phase 3A: Rich text & content reading ---

  {
    name: "pptx_set_placeholder_rich_text",
    description:
      "Write multi-paragraph rich text with per-run formatting into a placeholder. Use this instead of pptx_set_placeholder_text when you need multiple paragraphs, bullet levels, or mixed formatting within the same placeholder.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        placeholder_name: z.string().min(1),
        paragraphs: z
          .array(
            z.object({
              text: z.string().optional().describe("Plain text shorthand (used when runs is omitted)."),
              runs: z
                .array(
                  z.object({
                    text: z.string(),
                    bold: z.boolean().optional(),
                    italic: z.boolean().optional(),
                    underline: z.boolean().optional(),
                    font_name: z.string().optional(),
                    font_size_pt: z.number().positive().optional(),
                    color_hex: colorHexSchema.optional(),
                  }),
                )
                .optional()
                .describe("Rich runs with per-run formatting. If omitted, 'text' field is used as plain text."),
              alignment: z.enum(["left", "center", "right", "justify"]).optional(),
              level: z.number().int().min(0).max(8).optional().describe("Bullet indent level (0 = no indent)."),
            }),
          )
          .min(1),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_add_text_box",
    description:
      "Add a free-form text box shape at the specified position. Supports simple text or rich multi-paragraph content.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        left: measurementSchema,
        top: measurementSchema,
        width: measurementSchema,
        height: measurementSchema,
        text_content: z.string().optional().describe("Simple text content (used when paragraphs is omitted)."),
        paragraphs: z
          .array(
            z.object({
              text: z.string().optional(),
              runs: z
                .array(
                  z.object({
                    text: z.string(),
                    bold: z.boolean().optional(),
                    italic: z.boolean().optional(),
                    underline: z.boolean().optional(),
                    font_name: z.string().optional(),
                    font_size_pt: z.number().positive().optional(),
                    color_hex: colorHexSchema.optional(),
                  }),
                )
                .optional(),
              alignment: z.enum(["left", "center", "right", "justify"]).optional(),
            }),
          )
          .optional()
          .describe("Rich paragraphs (overrides text_content if provided)."),
        alignment: z.enum(["left", "center", "right", "justify"]).optional(),
        font_name: z.string().optional(),
        font_size_pt: z.number().positive().optional(),
        bold: z.boolean().optional(),
        color_hex: colorHexSchema.optional(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_get_slide_text",
    description:
      "Extract ALL text from a slide — placeholders, text boxes, tables (every cell), and group shape children. Returns structured text items with shape info, content type, and full paragraph/run formatting. Essential for understanding slide content before transferring it.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_get_shape_details",
    description:
      "Get detailed info about any shape by name or ID — geometry, text with formatting, table content, picture info, or group children.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional().describe("Shape name (from pptx_get_slide or pptx_get_slide_text)."),
        shape_id: z.number().int().optional().describe("Shape ID (alternative to shape_name)."),
      })
      .strict(),
    mutating: false,
  },

  // --- Phase 3B: Table support ---

  {
    name: "pptx_add_table",
    description: "Create a table shape with the specified rows, columns, and position. Optionally populate with initial data.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        rows: z.number().int().min(1),
        cols: z.number().int().min(1),
        left: measurementSchema,
        top: measurementSchema,
        width: measurementSchema,
        height: measurementSchema,
        data: z
          .array(z.array(z.string()))
          .optional()
          .describe("Optional 2D array of cell text values [row][col]."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_get_table",
    description: "Read all cell content and formatting from a table shape.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional().describe("Table shape name."),
        shape_id: z.number().int().optional().describe("Table shape ID (alternative to shape_name)."),
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_set_table_cell",
    description: "Write text and formatting to a single table cell (0-based row/col).",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        row: z.number().int().min(0),
        col: z.number().int().min(0),
        text: z.string(),
        bold: z.boolean().optional(),
        italic: z.boolean().optional(),
        font_name: z.string().optional(),
        font_size_pt: z.number().positive().optional(),
        color_hex: colorHexSchema.optional(),
        fill_hex: colorHexSchema.optional().describe("Cell background fill color."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_set_table_data",
    description:
      "Batch-write entire table content in one call. Each cell can be a plain string or an object with text + formatting.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        data: z
          .array(
            z.array(
              z.union([
                z.string(),
                z.object({
                  text: z.string(),
                  bold: z.boolean().optional(),
                  italic: z.boolean().optional(),
                  font_name: z.string().optional(),
                  font_size_pt: z.number().positive().optional(),
                  color_hex: colorHexSchema.optional(),
                  fill_hex: colorHexSchema.optional(),
                }),
              ]),
            ),
          )
          .min(1),
      })
      .strict(),
    mutating: true,
  },

  // --- Phase 3C: Shapes, notes, extras ---

  {
    name: "pptx_add_shape",
    description:
      "Add an auto-shape (rectangle, oval, arrow, etc.) at the specified position with optional fill, outline, and text.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_type: z
          .string()
          .min(1)
          .describe(
            "Shape type: rectangle, rounded_rectangle, oval, diamond, triangle, right_arrow, left_arrow, up_arrow, down_arrow, pentagon, hexagon, chevron, star_5_point, line_inverse, cross, frame, rectangular_callout, rounded_rectangular_callout, cloud_callout, cloud.",
          ),
        left: measurementSchema,
        top: measurementSchema,
        width: measurementSchema,
        height: measurementSchema,
        fill_hex: colorHexSchema.optional(),
        line_hex: colorHexSchema.optional(),
        text: z.string().optional().describe("Optional text content inside the shape."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_delete_shape",
    description: "Remove a shape from a slide by name or ID.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_set_slide_notes",
    description: "Write or replace the speaker notes text for a slide.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        notes_text: z.string(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_set_shape_text",
    description:
      "Write rich text content on any non-placeholder shape (text boxes, auto-shapes). Supports multi-paragraph with per-run formatting.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        text_content: z.string().optional().describe("Simple text (used when paragraphs is omitted)."),
        paragraphs: z
          .array(
            z.object({
              text: z.string().optional(),
              runs: z
                .array(
                  z.object({
                    text: z.string(),
                    bold: z.boolean().optional(),
                    italic: z.boolean().optional(),
                    underline: z.boolean().optional(),
                    font_name: z.string().optional(),
                    font_size_pt: z.number().positive().optional(),
                    color_hex: colorHexSchema.optional(),
                  }),
                )
                .optional(),
              alignment: z.enum(["left", "center", "right", "justify"]).optional(),
              level: z.number().int().min(0).max(8).optional(),
            }),
          )
          .optional(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_get_slide_xml",
    description: "Return the raw OOXML content of a slide for debugging and inspection.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
      })
      .strict(),
    mutating: false,
  },

  // --- Phase 4: Flexibility tools ---

  {
    name: "pptx_set_shape_properties",
    description: "Modify the properties (position, size, rotation, fill, outline, name) of an existing shape.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        left: measurementSchema.optional(),
        top: measurementSchema.optional(),
        width: measurementSchema.optional(),
        height: measurementSchema.optional(),
        rotation: z.number().optional().describe("Rotation in degrees."),
        fill_hex: z.union([colorHexSchema, z.literal("none")]).optional().describe("Background fill color, or 'none' for transparent."),
        line_hex: z.union([colorHexSchema, z.literal("none")]).optional().describe("Outline color, or 'none' for no outline."),
        line_width_pt: z.number().positive().optional().describe("Outline width in points."),
        name: z.string().optional().describe("New name for the shape."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_clone_shape",
    description: "Create an exact duplicate of a shape, optionally placing it on a different slide or with an offset.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        target_slide_index: slideIndexSchema.optional().describe("Slide to place the clone on (defaults to same slide)."),
        offset_left: measurementSchema.optional().describe("Offset from original X position (e.g., '0.5in')."),
        offset_top: measurementSchema.optional().describe("Offset from original Y position (e.g., '0.5in')."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_group_shapes",
    description: "Group multiple shapes together into a single group shape. The shapes must be on the same slide.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_names: z.array(z.string()).optional(),
        shape_ids: z.array(z.number().int()).optional(),
        group_name: z.string().optional().describe("Name for the new group shape."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_ungroup_shapes",
    description: "Ungroup a group shape, replacing it with its individual child shapes.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_set_shape_z_order",
    description: "Change the visual layering (z-order) of a shape (front, back, forward, backward).",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        action: z.enum(["front", "back", "forward", "backward"]).describe("front/back moves to extremes; forward/backward moves 1 level."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_add_image",
    description: "Add a standalone image to a slide. Supports JPG, PNG, GIF, BMP, etc.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        image_path: z.string().describe("Absolute file path to the image."),
        left: measurementSchema.optional(),
        top: measurementSchema.optional(),
        width: measurementSchema.optional(),
        height: measurementSchema.optional(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_add_line",
    description: "Add a freeform straight line/connector between two points.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        begin_x: measurementSchema,
        begin_y: measurementSchema,
        end_x: measurementSchema,
        end_y: measurementSchema,
        color_hex: colorHexSchema.optional(),
        width_pt: z.number().positive().optional().describe("Line width in points."),
        dash_style: z.string().optional().describe("E.g., 'dash', 'sysDash', 'sysDot', 'lgDash'."),
        line_name: z.string().optional(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_find_replace_text",
    description: "Find and replace text case-sensitively or insensitively across entire presentation or specific slides. Scans text boxes, tables, and grouped shapes.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        find_text: z.string().min(1),
        replace_text: z.string(),
        case_sensitive: z.boolean().optional().describe("Defaults to true."),
        slide_indices: z.array(slideIndexSchema).optional().describe("Limit to specific slides. Defaults to all slides."),
      })
      .strict(),
    mutating: true,
  },
];

export const TOOL_MAP = new Map(TOOL_DEFINITIONS.map((tool) => [tool.name, tool]));
