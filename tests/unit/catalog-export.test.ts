import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

import { TOOL_DEFINITIONS } from "../../src/tools/catalog.js";
import { exportToolCatalog, PYTHON_CATALOG_PATH, renderToolCatalog } from "../../src/tools/export.js";

describe("catalog exported to Python", () => {
  it("is up to date with src/tools/catalog.ts (run `npm run gen:catalog` if this fails)", () => {
    const committed = readFileSync(new URL(`../../${PYTHON_CATALOG_PATH}`, import.meta.url), "utf8");

    expect(committed).toBe(renderToolCatalog());
  });

  it("lists every tool with its category and an object input schema", () => {
    const exported = exportToolCatalog();

    expect(exported.map((tool) => tool.name)).toEqual(TOOL_DEFINITIONS.map((tool) => tool.name));
    for (const tool of exported) {
      expect(tool.category, tool.name).toBeTruthy();
      expect(tool.input_schema.type, tool.name).toBe("object");
    }
  });

  it("keeps measurement units case-insensitive for Python's regex engine", () => {
    const addTextBox = exportToolCatalog().find((tool) => tool.name === "pptx_add_text_box")!;
    const left = (addTextBox.input_schema.properties as Record<string, { anyOf: Array<{ pattern?: string }> }>).left;

    expect(left.anyOf.find((variant) => variant.pattern)?.pattern).toMatch(/^\(\?i\)/);
  });
});
