import { describe, expect, it } from "vitest";

import { TOOL_DEFINITIONS } from "../../src/tools/catalog.js";

describe("tool catalog", () => {
  it("contains only unique names", () => {
    const names = TOOL_DEFINITIONS.map((definition) => definition.name);
    expect(new Set(names).size).toBe(names.length);
  });

  it("contains all required phase 1-2 tools", () => {
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
    ].forEach((tool) => expect(names.has(tool)).toBe(true));
  });
});
