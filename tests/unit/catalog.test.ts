import { describe, expect, it } from "vitest";

import { TOOL_DEFINITIONS } from "../../src/tools/catalog.js";

describe("tool catalog", () => {
  it("contains only unique names", () => {
    const names = TOOL_DEFINITIONS.map((definition) => definition.name);
    expect(new Set(names).size).toBe(names.length);
  });

  it("contains all required phase 1-3 tools", () => {
    const names = new Set(TOOL_DEFINITIONS.map((definition) => definition.name));
    [
      "pptx_get_engine_info",
      "pptx_create_presentation",
      "pptx_open_presentation",
      "pptx_save_presentation",
      "pptx_close_presentation",
      "pptx_list_open_presentations",
      "pptx_get_presentation_state",
      "pptx_get_layouts",
      "pptx_get_layout_detail",
      "pptx_get_masters",
      "pptx_get_theme",
      "pptx_get_slide",
      "pptx_add_slide",
      "pptx_duplicate_slide",
      "pptx_delete_slide",
      "pptx_reorder_slides",
      "pptx_move_slide",
      "pptx_set_slide_background",
      "pptx_get_slide_snapshot",
      "pptx_get_placeholders",
      "pptx_set_placeholder_text",
      "pptx_set_placeholder_image",
      "pptx_clear_placeholder",
      "pptx_get_placeholder_text",
      // Phase 3A
      "pptx_set_placeholder_rich_text",
      "pptx_add_text_box",
      "pptx_get_slide_text",
      "pptx_get_shape_details",
      // Phase 3B
      "pptx_add_table",
      "pptx_get_table",
      "pptx_set_table_cell",
      "pptx_set_table_data",
      // Phase 3C
      "pptx_add_shape",
      "pptx_delete_shape",
      "pptx_set_slide_notes",
      "pptx_set_shape_text",
      "pptx_get_slide_xml",
      // Phase 4
      "pptx_set_shape_properties",
      "pptx_clone_shape",
      "pptx_group_shapes",
      "pptx_ungroup_shapes",
      "pptx_set_shape_z_order",
      "pptx_add_image",
      "pptx_add_line",
      "pptx_find_replace_text",
    ].forEach((tool) => expect(names.has(tool)).toBe(true));
  });
});
