/** Guided workflows exposed as MCP prompts. Each one expands to a single user message. */

interface PromptArgument {
  name: string;
  description: string;
  required: boolean;
}

interface PromptDefinition {
  name: string;
  title: string;
  description: string;
  arguments: PromptArgument[];
  /** Tools the workflow cannot do without; the prompt is hidden when the tool profile lacks one. */
  requiredTools: string[];
  render: (args: Record<string, string>) => string;
}

export interface PromptMessage {
  role: "user";
  content: { type: "text"; text: string };
}

const VERIFY_STEP =
  "Check the result: read slides back with pptx_get_slide_text, and if pptx_get_slide_snapshot is available " +
  "(it needs LibreOffice in OOXML mode) look at a few slides with slide_indices. Fix what looks wrong.";

const PROMPTS: PromptDefinition[] = [
  {
    name: "create_deck_from_outline",
    title: "Create a deck from an outline",
    description: "Build a presentation from a text outline, optionally on top of a template, and save it.",
    arguments: [
      { name: "outline", description: "The outline: one entry per slide, with its title and content.", required: true },
      { name: "output_path", description: "Absolute path of the .pptx file to write.", required: true },
      { name: "template_path", description: "Absolute path of a .pptx/.potx template to start from.", required: false },
    ],
    requiredTools: ["pptx_create_presentation", "pptx_batch"],
    render: ({ outline, output_path: outputPath, template_path: templatePath }) =>
      [
        "Create a PowerPoint presentation from the outline below using the pptx_* tools.",
        "",
        "Steps:",
        templatePath
          ? `1. Call pptx_create_presentation with template_path "${templatePath}".`
          : "1. Call pptx_create_presentation.",
        "2. Call pptx_get_layouts, then pptx_get_layout_detail for the layouts you plan to use, so you know the " +
          "exact layout names and placeholder names. Never guess them.",
        "3. Plan one slide per outline entry and pick the layout that fits its content (title, bullets, " +
          "two columns, table, chart, image).",
        "4. Build the slides with pptx_batch, a few slides per batch: pptx_add_slide, then " +
          "pptx_set_placeholder_text for each placeholder. Prefer placeholders over free text boxes, so the " +
          "template's formatting applies. Use pptx_add_table or pptx_add_chart for tabular or numeric content.",
        "5. If the template came with slides of its own that the outline does not need, delete them.",
        `6. ${VERIFY_STEP}`,
        `7. Call pptx_save_presentation with output_path "${outputPath}", then pptx_close_presentation.`,
        "",
        "Keep slide text short: a title and at most five or six bullets per slide. Put detail in speaker notes " +
          "with pptx_set_slide_notes.",
        "",
        "Outline:",
        outline,
      ].join("\n"),
  },
  {
    name: "apply_template_to_deck",
    title: "Bring a deck in line with a template",
    description: "Check an existing presentation against a reference template and fix the deviations.",
    arguments: [
      { name: "file_path", description: "Absolute path of the presentation to fix.", required: true },
      { name: "template_path", description: "Absolute path of the reference template.", required: true },
      { name: "output_path", description: "Absolute path for the corrected copy. Defaults to asking.", required: false },
    ],
    requiredTools: ["pptx_check_template_conformance", "pptx_check_visual_consistency", "pptx_check_positions"],
    render: ({ file_path: filePath, template_path: templatePath, output_path: outputPath }) =>
      [
        `Bring the presentation "${filePath}" in line with the template "${templatePath}" using the pptx_* tools.`,
        "",
        "Steps:",
        `1. Call pptx_open_presentation with file_path "${filePath}". The original file is not modified until you save.`,
        `2. Call pptx_check_template_conformance with template_path "${templatePath}" to list deviations in ` +
          "theme colors, fonts and layouts.",
        "3. Call pptx_check_visual_consistency and pptx_check_positions to find slides whose fonts, colors or " +
          "element positions differ from the rest of the deck.",
        "4. Summarise what you found and what you intend to change before changing anything substantial.",
        "5. Fix the deviations slide by slide, grouping the edits for one slide in a pptx_batch: " +
          "pptx_set_shape_properties for positions, sizes and fill colors, the text formatting tools for fonts, " +
          "pptx_set_placeholder_text when text must move into a proper placeholder. Do not change the wording " +
          "of the content.",
        `6. ${VERIFY_STEP} Run the checks from steps 2 and 3 again to confirm the deviations are gone.`,
        outputPath
          ? `7. Call pptx_save_presentation with output_path "${outputPath}".`
          : "7. Ask where to save the corrected copy, then call pptx_save_presentation. Do not overwrite the " +
            "original unless told to.",
      ].join("\n"),
  },
];

const PROMPT_MAP = new Map(PROMPTS.map((prompt) => [prompt.name, prompt]));

type ToolFilter = (name: string) => boolean;

export function listPrompts(isToolAvailable: ToolFilter): Array<Omit<PromptDefinition, "render" | "requiredTools">> {
  return PROMPTS.filter((prompt) => prompt.requiredTools.every(isToolAvailable)).map(
    ({ render: _render, requiredTools: _requiredTools, ...metadata }) => metadata,
  );
}

/** Returns undefined for an unknown or unavailable prompt; throws if a required argument is missing. */
export function getPrompt(
  name: string,
  args: Record<string, string>,
  isToolAvailable: ToolFilter,
): { description: string; messages: PromptMessage[] } | undefined {
  const prompt = PROMPT_MAP.get(name);
  if (!prompt || !prompt.requiredTools.every(isToolAvailable)) {
    return undefined;
  }

  const missing = prompt.arguments.filter((argument) => argument.required && !args[argument.name]?.trim());
  if (missing.length > 0) {
    throw new Error(`Missing required argument(s): ${missing.map((argument) => argument.name).join(", ")}`);
  }

  return {
    description: prompt.description,
    messages: [{ role: "user", content: { type: "text", text: prompt.render(args) } }],
  };
}
