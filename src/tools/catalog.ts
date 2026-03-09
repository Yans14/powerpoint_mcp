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

/** Reusable shape identifier fields + refinement (at least one of shape_name/shape_id required). */
const shapeIdentifierSchema = z
  .object({
    shape_name: z.string().optional().describe("Shape name (from pptx_get_slide or pptx_get_slide_shapes)."),
    shape_id: z.number().int().optional().describe("Shape ID (alternative to shape_name)."),
  })
  .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
    message: "Either shape_name or shape_id is required.",
  });

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
        italic: z.boolean().optional(),
        underline: z.boolean().optional(),
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
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
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
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
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
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
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
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
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
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
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
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
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
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
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
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
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
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
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
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
    mutating: true,
  },
  {
    name: "pptx_add_image",
    description: "Add a standalone image to a slide. Supports JPG, PNG, GIF, BMP, etc.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        image_path: absolutePathSchema.describe("Absolute file path to the image."),
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

  // --- Phase 5: Chart tools ---

  {
    name: "pptx_add_chart",
    description:
      "Add a chart to a slide. For category-based charts (bar, column, line, pie, area) provide categories and series with values. For XY/scatter and bubble charts provide series with data_points.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        chart_type: z
          .enum([
            "column_clustered",
            "column_stacked",
            "column_stacked_100",
            "bar_clustered",
            "bar_stacked",
            "bar_stacked_100",
            "line",
            "line_markers",
            "line_stacked",
            "pie",
            "pie_exploded",
            "doughnut",
            "area",
            "area_stacked",
            "area_stacked_100",
            "xy_scatter",
            "xy_scatter_lines",
            "xy_scatter_smooth",
            "bubble",
            "radar",
            "stock_hlc",
            "stock_ohlc",
            "three_d_column",
            "three_d_bar_clustered",
            "three_d_pie",
            "three_d_line",
          ])
          .describe("Chart type."),
        left: measurementSchema.optional(),
        top: measurementSchema.optional(),
        width: measurementSchema.optional(),
        height: measurementSchema.optional(),
        categories: z.array(z.string()).optional().describe("Category labels (for bar/column/line/pie/area charts)."),
        series: z
          .array(
            z.object({
              name: z.string().optional().describe("Series display name."),
              values: z.array(z.number()).optional().describe("Data values (for category charts)."),
              data_points: z
                .array(
                  z.object({
                    x: z.number(),
                    y: z.number(),
                    size: z.number().optional().describe("Bubble size (bubble charts only)."),
                  }),
                )
                .optional()
                .describe("Data points (for XY scatter or bubble charts)."),
            }),
          )
          .min(1),
        has_legend: z.boolean().optional(),
        legend_position: z.enum(["bottom", "corner", "left", "right", "top"]).optional(),
        has_data_labels: z.boolean().optional(),
        data_label_number_format: z.string().optional().describe("E.g., '0%', '#,##0', '0.0'."),
        chart_style: z.number().int().min(1).max(48).optional().describe("Built-in chart style number."),
        title: z.string().optional().describe("Chart title text."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_get_chart_data",
    description: "Read the chart type, categories, series data, legend/label info, and title from an existing chart shape.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
      })
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
    mutating: false,
  },
  {
    name: "pptx_update_chart_data",
    description:
      "Replace the data behind an existing chart while preserving its visual formatting. Provide new categories (optional) and series.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        categories: z.array(z.string()).optional().describe("New category labels (keeps existing if omitted)."),
        series: z
          .array(
            z.object({
              name: z.string().optional(),
              values: z.array(z.number()).optional(),
              data_points: z
                .array(z.object({ x: z.number(), y: z.number(), size: z.number().optional().describe("Bubble size (bubble charts only).") }))
                .optional()
                .describe("For XY scatter or bubble charts."),
            }),
          )
          .min(1),
      })
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
    mutating: true,
  },
  {
    name: "pptx_set_chart_style",
    description: "Modify chart visual properties: legend, data labels, title, and built-in style number.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        has_legend: z.boolean().optional(),
        legend_position: z.enum(["bottom", "corner", "left", "right", "top"]).optional(),
        legend_in_layout: z.boolean().optional(),
        has_data_labels: z.boolean().optional(),
        data_label_number_format: z.string().optional(),
        data_label_position: z
          .enum(["center", "inside_end", "outside_end", "inside_base", "above", "below", "left", "right", "best_fit"])
          .optional(),
        chart_style: z.number().int().min(1).max(48).optional(),
        title: z.string().nullable().optional().describe("Chart title. Set to null or '' to remove."),
      })
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
    mutating: true,
  },

  // --- Phase 6: Agent workflow tools ---

  {
    name: "pptx_copy_shape_between_decks",
    description:
      "Copy a shape (including images, tables, charts) from one open presentation to another. Handles image relationship transfer automatically.",
    schema: z
      .object({
        source_presentation_id: presentationIdSchema,
        source_slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        target_presentation_id: presentationIdSchema,
        target_slide_index: slideIndexSchema,
        offset_left: measurementSchema.optional().describe("Offset from original X position."),
        offset_top: measurementSchema.optional().describe("Offset from original Y position."),
      })
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
    mutating: true,
  },
  {
    name: "pptx_get_slide_shapes",
    description:
      "Get a lightweight listing of all shapes on a slide: name, id, type, position, size, and flags (placeholder, text_frame, table, chart). Much faster than get_slide_text for shape discovery.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_set_table_cell_merge",
    description: "Merge a rectangular range of cells in a table. Uses 0-based row/col indices. The merged cell retains the top-left cell's content.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        start_row: z.number().int().min(0),
        start_col: z.number().int().min(0),
        end_row: z.number().int().min(0),
        end_col: z.number().int().min(0),
      })
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      })
      .refine((d) => d.end_row >= d.start_row && d.end_col >= d.start_col, {
        message: "end_row must be >= start_row and end_col must be >= start_col.",
      }),
    mutating: true,
  },

  // --- Phase 7: Formatting & Fidelity tools ---

  {
    name: "pptx_set_paragraph_spacing",
    description: "Control line spacing and spacing before/after paragraphs.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        paragraph_index: z.number().int().min(0),
        line_spacing: z.number().optional().describe("Line spacing in points (e.g., 14.0)."),
        space_before: z.number().optional().describe("Spacing before paragraph in points."),
        space_after: z.number().optional().describe("Spacing after paragraph in points."),
      })
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
    mutating: true,
  },
  {
    name: "pptx_set_text_box_properties",
    description: "Control margins, word wrap, auto-fit, and vertical alignment for text frames.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        margin_left: measurementSchema.optional(),
        margin_top: measurementSchema.optional(),
        margin_right: measurementSchema.optional(),
        margin_bottom: measurementSchema.optional(),
        word_wrap: z.boolean().optional(),
        auto_size: z.enum(["none", "shape_to_fit_text", "text_to_fit_shape"]).optional(),
        vertical_alignment: z.enum(["top", "middle", "bottom"]).optional(),
      })
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
    mutating: true,
  },
  {
    name: "pptx_set_table_style",
    description: "Apply table styles (banding, header rows) and optional style IDs.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        first_row: z.boolean().optional(),
        last_row: z.boolean().optional(),
        first_col: z.boolean().optional(),
        last_col: z.boolean().optional(),
        banded_rows: z.boolean().optional(),
        banded_cols: z.boolean().optional(),
        style_id: z.string().optional().describe("Internal PPTX table style GUID if known."),
      })
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
    mutating: true,
  },
  {
    name: "pptx_set_shape_fill_gradient",
    description: "Apply linear gradients to shape backgrounds. (Radial not natively supported by python-pptx write API).",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        shape_name: z.string().optional(),
        shape_id: z.number().int().optional(),
        angle: z.number().optional().describe("Linear gradient angle in degrees (0-359). 0 is left-to-right."),
        stops: z
          .array(
            z.object({
              position: z.number().min(0).max(1).describe("Position from 0.0 to 1.0"),
              color_hex: z.string().regex(/^[0-9A-Fa-f]{6}$/, "Must be 6-character hex without #"),
            })
          )
          .optional()
          .describe("Array of gradient stops."),
      })
      .strict()
      .refine((d) => d.shape_name !== undefined || d.shape_id !== undefined, {
        message: "Either shape_name or shape_id is required.",
      }),
    mutating: true,
  },
  {
    name: "pptx_add_connector",
    description: "Add a connecting line that explicitly snaps to other shapes.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_index: slideIndexSchema,
        begin_shape_name: z.string().optional(),
        begin_shape_id: z.number().int().optional(),
        begin_connection_site: z.number().int().optional().describe("Usually 0-3 for top/bottom/left/right defaults."),
        end_shape_name: z.string().optional(),
        end_shape_id: z.number().int().optional(),
        end_connection_site: z.number().int().optional().describe("Usually 0-3 for top/bottom/left/right defaults."),
        connector_type: z.enum(["straight", "elbow", "curve"]).optional().default("straight"),
        color_hex: z.string().regex(/^[0-9A-Fa-f]{6}$/).optional(),
        width_pt: z.number().optional(),
      })
      .strict(),
    mutating: true,
  },

  // --- Phase 8: Agent Orchestrator tools ---

  {
    name: "pptx_agent_start",
    description:
      "Start an AI agent task on a presentation with a natural language query. Analyzes the presentation, generates clarifying questions, and returns a task_id. Requires LLM configuration via environment variables (PPTX_LLM_PROVIDER + ANTHROPIC_API_KEY or OPENAI_API_KEY).",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        query: z
          .string()
          .min(1)
          .describe(
            "Natural language description of the desired transformation, e.g. 'Convert from Acme Corp branding to my company template'.",
          ),
        skip_questions: z
          .boolean()
          .optional()
          .describe("If true, skip clarifying questions and go straight to plan generation. Default false."),
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_agent_respond",
    description:
      "Respond to the agent's clarifying questions to refine the execution plan. Call after pptx_agent_start returns state='clarifying'.",
    schema: z
      .object({
        task_id: z.string().uuid().describe("Task ID returned by pptx_agent_start."),
        answers: z
          .array(
            z.object({
              question_id: z.string().min(1).describe("question_id from the questions array."),
              answer: z.string().min(1).describe("Answer text or letter choice (e.g. 'A', 'B', or full text)."),
            }),
          )
          .min(1),
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_agent_execute",
    description:
      "Execute the approved agent plan step by step. Must set confirm=true. Creates a rollback snapshot automatically. Call pptx_agent_rollback to undo if needed.",
    schema: z
      .object({
        task_id: z.string().uuid(),
        confirm: z.literal(true).describe("Must be exactly true to proceed with execution."),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_agent_status",
    description: "Get current state, progress, and execution log of an agent task.",
    schema: z
      .object({
        task_id: z.string().uuid(),
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_agent_rollback",
    description: "Restore the presentation to its pre-execution state. Only valid after pptx_agent_execute has run.",
    schema: z
      .object({
        task_id: z.string().uuid(),
      })
      .strict(),
    mutating: true,
  },
  {
    name: "pptx_agent_cancel",
    description: "Cancel an agent task and release its resources (snapshot file, secondary presentations opened by the plan).",
    schema: z
      .object({
        task_id: z.string().uuid(),
      })
      .strict(),
    mutating: true,
  },

  // --- Phase 8: Checker tools ---

  {
    name: "pptx_check_positions",
    description:
      "Check all shapes for overlaps, out-of-bounds positions, and near-alignment. Returns per-issue details with severity.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_indices: z
          .array(slideIndexSchema)
          .optional()
          .describe("Slides to check. Defaults to all slides."),
        check_overlaps: z.boolean().optional().describe("Check for overlapping shapes. Default true."),
        check_bounds: z
          .boolean()
          .optional()
          .describe("Check for shapes extending outside slide edges. Default true."),
        check_alignment: z
          .boolean()
          .optional()
          .describe("Flag shapes that are 'almost' aligned (within tolerance). Default false."),
        tolerance_px: z
          .number()
          .int()
          .min(0)
          .max(100)
          .optional()
          .describe("Pixel tolerance for overlap and alignment checks. Default 5."),
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_check_visual_consistency",
    description: "Audit font families, font sizes, and colors across all slides for visual inconsistencies.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        slide_indices: z.array(slideIndexSchema).optional(),
        check_fonts: z.boolean().optional().describe("Check font family consistency. Default true."),
        check_colors: z.boolean().optional().describe("Check color consistency and off-brand colors. Default true."),
        check_sizes: z.boolean().optional().describe("Check for unusual font size outliers. Default true."),
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_check_content",
    description:
      "Find empty placeholders, default/template placeholder text, and slides missing expected content types.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        check_empty_placeholders: z
          .boolean()
          .optional()
          .describe("Flag placeholders with no text or image. Default true."),
        check_default_text: z
          .boolean()
          .optional()
          .describe("Flag placeholders still containing default template text like 'Click to add title'. Default true."),
        slide_indices: z.array(slideIndexSchema).optional(),
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_check_template_conformance",
    description:
      "Compare a presentation against a reference template, checking theme colors, font schemes, and layout usage. Returns a conformance_score from 0.0 to 1.0.",
    schema: z
      .object({
        presentation_id: presentationIdSchema,
        template_path: absolutePathSchema.describe("Absolute path to .pptx or .potx reference template."),
        check_theme: z.boolean().optional().describe("Compare theme color schemes. Default true."),
        check_fonts: z.boolean().optional().describe("Compare major/minor font schemes. Default true."),
        check_layouts: z
          .boolean()
          .optional()
          .describe("Check that slide layouts exist in the template. Default true."),
      })
      .strict(),
    mutating: false,
  },
  {
    name: "pptx_diff_presentations",
    description:
      "Generate a structured before/after diff between two open presentations. Compares slide counts, added/removed slides, and per-shape text changes on modified slides.",
    schema: z
      .object({
        presentation_id_a: presentationIdSchema.describe("The 'before' presentation ID."),
        presentation_id_b: presentationIdSchema.describe("The 'after' presentation ID."),
        deep_diff: z
          .boolean()
          .optional()
          .describe("If true, include per-shape text content changes. Default true."),
      })
      .strict(),
    mutating: false,
  },
];

export const TOOL_MAP = new Map(TOOL_DEFINITIONS.map((tool) => [tool.name, tool]));
