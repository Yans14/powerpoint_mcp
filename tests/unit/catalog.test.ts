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
      // Phase 5
      "pptx_add_chart",
      "pptx_get_chart_data",
      "pptx_update_chart_data",
      "pptx_set_chart_style",
      // Phase 6
      "pptx_copy_shape_between_decks",
      "pptx_get_slide_shapes",
      "pptx_set_table_cell_merge",
      // Phase 7
      "pptx_set_paragraph_spacing",
      "pptx_set_text_box_properties",
      "pptx_set_table_style",
      "pptx_set_shape_fill_gradient",
      "pptx_add_connector",
    ].forEach((tool) => expect(names.has(tool)).toBe(true));
  });
});

describe("agent and checker tools in catalog", () => {
  const names = new Set(TOOL_DEFINITIONS.map((d) => d.name));

  it("contains all Phase 8 agent tools", () => {
    const agentTools = [
      "pptx_agent_start",
      "pptx_agent_respond",
      "pptx_agent_execute",
      "pptx_agent_status",
      "pptx_agent_rollback",
      "pptx_agent_cancel",
    ];
    for (const tool of agentTools) {
      expect(names.has(tool), `Missing tool: ${tool}`).toBe(true);
    }
  });

  it("contains all Phase 8 checker tools", () => {
    const checkerTools = [
      "pptx_check_positions",
      "pptx_check_visual_consistency",
      "pptx_check_content",
      "pptx_check_template_conformance",
      "pptx_diff_presentations",
    ];
    for (const tool of checkerTools) {
      expect(names.has(tool), `Missing tool: ${tool}`).toBe(true);
    }
  });

  it("pptx_agent_execute is marked mutating=true", () => {
    const def = TOOL_DEFINITIONS.find((d) => d.name === "pptx_agent_execute");
    expect(def?.mutating).toBe(true);
  });

  it("pptx_agent_start is marked mutating=false", () => {
    const def = TOOL_DEFINITIONS.find((d) => d.name === "pptx_agent_start");
    expect(def?.mutating).toBe(false);
  });

  it("checker tools are all mutating=false", () => {
    const checkerTools = [
      "pptx_check_positions",
      "pptx_check_visual_consistency",
      "pptx_check_content",
      "pptx_check_template_conformance",
      "pptx_diff_presentations",
    ];
    for (const tool of checkerTools) {
      const def = TOOL_DEFINITIONS.find((d) => d.name === tool);
      expect(def?.mutating, `${tool} should be non-mutating`).toBe(false);
    }
  });

  it("all names are still unique", () => {
    const allNames = TOOL_DEFINITIONS.map((d) => d.name);
    expect(new Set(allNames).size).toBe(allNames.length);
  });
});
