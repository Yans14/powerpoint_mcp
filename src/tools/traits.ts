import type { ToolDefinition } from "./catalog.js";

/**
 * Behavioural traits of tools, declared explicitly rather than guessed from tool names.
 * A tool absent from a set does not have that trait.
 */

/** Removes content, discards unsaved work, or overwrites a file on disk. */
const DESTRUCTIVE_TOOLS = new Set([
  "pptx_save_presentation", // overwrites output_path if it exists
  "pptx_close_presentation", // discards unsaved changes
  "pptx_delete_slide",
  "pptx_delete_shape",
  "pptx_clear_placeholder",
  "pptx_agent_rollback",
  "pptx_agent_cancel",
  "pptx_batch", // may contain any of the above
]);

/** Mutating tools for which repeating the same call leaves the presentation unchanged. */
const IDEMPOTENT_MUTATIONS = new Set([
  "pptx_save_presentation",
  "pptx_set_slide_background",
  "pptx_set_placeholder_text",
  "pptx_set_placeholder_rich_text",
  "pptx_clear_placeholder",
  "pptx_set_table_cell",
  "pptx_set_table_data",
  "pptx_set_table_style",
  "pptx_set_slide_notes",
  "pptx_set_shape_text",
  "pptx_set_shape_properties",
  "pptx_set_shape_fill_gradient",
  "pptx_set_paragraph_spacing",
  "pptx_set_text_box_properties",
  "pptx_update_chart_data",
  "pptx_set_chart_style",
]);

/** Reach outside this machine: these call an LLM provider's API. */
const OPEN_WORLD_TOOLS = new Set(["pptx_agent_start", "pptx_agent_respond"]);

/** Not read-only in the catalog's sense, yet each call creates state or spends LLM calls. */
const NON_REPEATABLE_READS = new Set(["pptx_agent_start", "pptx_agent_respond"]);

/** The `core` profile: enough to build and edit a deck end to end with about half the tools. */
const CORE_TOOLS = new Set([
  "pptx_create_presentation",
  "pptx_open_presentation",
  "pptx_save_presentation",
  "pptx_close_presentation",
  "pptx_list_open_presentations",
  "pptx_get_presentation_state",
  "pptx_get_layouts",
  "pptx_get_layout_detail",
  "pptx_get_slide",
  "pptx_get_slide_text",
  "pptx_get_slide_shapes",
  "pptx_get_slide_snapshot",
  "pptx_add_slide",
  "pptx_duplicate_slide",
  "pptx_delete_slide",
  "pptx_move_slide",
  "pptx_get_placeholders",
  "pptx_set_placeholder_text",
  "pptx_set_placeholder_image",
  "pptx_add_text_box",
  "pptx_add_image",
  "pptx_add_shape",
  "pptx_add_table",
  "pptx_set_table_data",
  "pptx_add_chart",
  "pptx_set_shape_text",
  "pptx_set_shape_properties",
  "pptx_delete_shape",
  "pptx_set_slide_notes",
  "pptx_find_replace_text",
  "pptx_batch",
]);

export const TRAIT_SETS = {
  destructive: DESTRUCTIVE_TOOLS,
  idempotentMutation: IDEMPOTENT_MUTATIONS,
  openWorld: OPEN_WORLD_TOOLS,
  nonRepeatableRead: NON_REPEATABLE_READS,
  core: CORE_TOOLS,
} as const;

export type ToolProfile = "full" | "core";

export interface ToolAnnotations {
  title: string;
  readOnlyHint: boolean;
  destructiveHint: boolean;
  idempotentHint: boolean;
  openWorldHint: boolean;
}

/** "pptx_set_shape_text" -> "Set Shape Text". */
export function toolTitle(name: string): string {
  return name
    .replace(/^pptx_/, "")
    .split("_")
    .map((word) => (word === "xml" ? "XML" : word.charAt(0).toUpperCase() + word.slice(1)))
    .join(" ");
}

export function toolAnnotations(definition: ToolDefinition): ToolAnnotations {
  const { name, mutating } = definition;
  return {
    title: toolTitle(name),
    readOnlyHint: !mutating,
    destructiveHint: DESTRUCTIVE_TOOLS.has(name),
    idempotentHint: mutating ? IDEMPOTENT_MUTATIONS.has(name) : !NON_REPEATABLE_READS.has(name),
    openWorldHint: OPEN_WORLD_TOOLS.has(name),
  };
}

/** Mutating tools whose result carries a presentation state, and so accept `include_state`. */
export function acceptsIncludeState(definition: ToolDefinition): boolean {
  return definition.mutating && !definition.name.startsWith("pptx_agent_");
}

export function parseToolProfile(value: string | undefined): ToolProfile {
  const normalized = (value ?? "").trim().toLowerCase();
  if (normalized === "" || normalized === "full") {
    return "full";
  }
  if (normalized === "core") {
    return "core";
  }
  throw new Error(`Unknown PPTX_TOOL_PROFILE "${value}". Use "full" or "core".`);
}

export function isToolInProfile(name: string, profile: ToolProfile): boolean {
  return profile === "full" || CORE_TOOLS.has(name);
}
