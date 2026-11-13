import { spawn } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { TOOL_DEFINITIONS } from "../../src/tools/catalog.js";

const REPO_ROOT = fileURLToPath(new URL("../../", import.meta.url));
const TSX_CLI = path.join(REPO_ROOT, "node_modules/tsx/dist/cli.mjs");
const SERVER_ENTRY = path.join(REPO_ROOT, "src/index.ts");
const SERVER_ARGS = [TSX_CLI, SERVER_ENTRY];

// 1x1 red PNG.
const PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";

interface ToolResult {
  isError?: boolean;
  content: Array<{ type: string; text?: string }>;
}

describe("MCP server over stdio", () => {
  let workDir: string;
  let client: Client;

  async function callTool(name: string, args: Record<string, unknown>): Promise<{ isError: boolean; data: any }> {
    const result = (await client.callTool({ name, arguments: args })) as ToolResult;
    return { isError: result.isError === true, data: JSON.parse(result.content[0].text ?? "null") };
  }

  beforeAll(async () => {
    workDir = await mkdtemp(path.join(os.tmpdir(), "pptx-mcp-it-"));
    client = new Client({ name: "integration-test", version: "0.0.0" });
    // Launch from an unrelated directory, as MCP clients do.
    await client.connect(new StdioClientTransport({ command: process.execPath, args: SERVER_ARGS, cwd: workDir }));
  }, 30_000);

  afterAll(async () => {
    await client?.close();
    await rm(workDir, { recursive: true, force: true });
  });

  it("advertises every tool with a usable input schema", async () => {
    const { tools } = await client.listTools();

    expect(tools.map((tool) => tool.name).sort()).toEqual(TOOL_DEFINITIONS.map((tool) => tool.name).sort());
    for (const tool of tools) {
      expect(tool.inputSchema.type, tool.name).toBe("object");
      expect(tool.inputSchema.properties, tool.name).toBeDefined();
    }

    const addSlide = tools.find((tool) => tool.name === "pptx_add_slide");
    expect(Object.keys(addSlide?.inputSchema.properties ?? {})).toEqual([
      "presentation_id",
      "layout_name",
      "position",
      "include_state",
    ]);
    expect(addSlide?.inputSchema.required).toEqual(["presentation_id", "layout_name"]);
  });

  it("builds, duplicates, saves and reopens a deck containing an image", async () => {
    const imagePath = path.join(workDir, "pixel.png");
    const outputPath = path.join(workDir, "deck.pptx");
    await writeFile(imagePath, Buffer.from(PNG_BASE64, "base64"));

    const created = await callTool("pptx_create_presentation", {});
    expect(created.isError).toBe(false);
    const presentationId = created.data.presentation_id as string;

    const steps: Array<[string, Record<string, unknown>]> = [
      ["pptx_add_slide", { presentation_id: presentationId, layout_name: "Title Slide" }],
      ["pptx_add_image", { presentation_id: presentationId, slide_index: 1, image_path: imagePath, left: "1in", top: "1in" }],
      ["pptx_duplicate_slide", { presentation_id: presentationId, source_index: 1 }],
      ["pptx_save_presentation", { presentation_id: presentationId, output_path: outputPath }],
    ];
    for (const [name, args] of steps) {
      const result = await callTool(name, args);
      expect(result.isError, `${name}: ${JSON.stringify(result.data)}`).toBe(false);
    }

    const reopened = await callTool("pptx_open_presentation", { file_path: outputPath });
    expect(reopened.isError).toBe(false);
    expect(reopened.data.slide_count).toBe(2);

    const duplicate = await callTool("pptx_get_slide", {
      presentation_id: reopened.data.presentation_id,
      slide_index: 2,
    });
    expect(duplicate.isError).toBe(false);
    expect(JSON.stringify(duplicate.data)).toContain("PICTURE");
  }, 30_000);

  it("returns results that satisfy every advertised output schema", async () => {
    // listTools makes the SDK client validate structuredContent against each tool's outputSchema.
    const { tools } = await client.listTools();
    const withOutputSchema = tools.filter((tool) => tool.outputSchema).map((tool) => tool.name);
    const created = await callTool("pptx_create_presentation", {});
    const presentationId = created.data.presentation_id as string;
    await callTool("pptx_add_slide", { presentation_id: presentationId, layout_name: "Title Slide" });

    const argumentsByTool: Record<string, Record<string, unknown>> = {
      pptx_list_open_presentations: {},
      pptx_get_presentation_state: { presentation_id: presentationId },
      pptx_get_layouts: { presentation_id: presentationId },
      pptx_get_layout_detail: { presentation_id: presentationId, layout_name: "Title Slide" },
      pptx_get_placeholders: { presentation_id: presentationId, slide_index: 1 },
    };
    expect(Object.keys(argumentsByTool).sort()).toEqual(withOutputSchema.sort());

    for (const [name, args] of Object.entries(argumentsByTool)) {
      const result = (await client.callTool({ name, arguments: args })) as ToolResult & { structuredContent?: unknown };
      expect(result.isError, name).toBeFalsy();
      expect(result.structuredContent, name).toEqual(JSON.parse(result.content[0].text ?? "null"));
    }

    await callTool("pptx_close_presentation", { presentation_id: presentationId });
  });

  it("runs a batch in one call and rolls it back when an operation fails", async () => {
    const created = await callTool("pptx_create_presentation", {});
    const presentationId = created.data.presentation_id as string;
    expect(created.data.presentation_state).toBeUndefined();
    expect(created.data.slide_count).toBe(0);

    const built = await callTool("pptx_batch", {
      presentation_id: presentationId,
      operations: [
        { tool: "pptx_add_slide", arguments: { layout_name: "Title Slide" } },
        { tool: "pptx_add_slide", arguments: { layout_name: "Blank" } },
        {
          tool: "pptx_add_text_box",
          arguments: { slide_index: 2, text_content: "Été 2026", left: "1in", top: "1in", width: "4in", height: "1in" },
        },
      ],
    });
    expect(built.isError, JSON.stringify(built.data)).toBe(false);
    expect(built.data).toMatchObject({ executed: 3, slide_count: 2 });

    const failed = await callTool("pptx_batch", {
      presentation_id: presentationId,
      operations: [
        { tool: "pptx_delete_slide", arguments: { slide_index: 1 } },
        { tool: "pptx_add_slide", arguments: { layout_name: "No Such Layout" } },
      ],
    });
    expect(failed.isError).toBe(true);
    expect(failed.data.code).toBe("not_found");
    expect(failed.data.details).toMatchObject({ failed_operation_index: 1, rolled_back: true });
    expect(failed.data.hint).toContain("pptx_get_layouts");

    const state = await callTool("pptx_get_presentation_state", { presentation_id: presentationId });
    expect(state.data.slides.map((slide: { layout: string }) => slide.layout)).toEqual(["Title Slide", "Blank"]);

    const text = await callTool("pptx_get_slide_text", { presentation_id: presentationId, slide_index: 2 });
    expect(JSON.stringify(text.data)).toContain("Été 2026");

    await callTool("pptx_close_presentation", { presentation_id: presentationId });
  }, 30_000);

  it("serves presentations as resources and rejects unknown URIs", async () => {
    const created = await callTool("pptx_create_presentation", {});
    const presentationId = created.data.presentation_id as string;

    const { resources } = await client.listResources();
    expect(resources.map((resource) => resource.uri)).toContain(`pptx://presentations/${presentationId}`);

    const layouts = await client.readResource({ uri: `pptx://presentations/${presentationId}/layouts` });
    expect(JSON.parse((layouts.contents[0] as { text: string }).text).layouts.length).toBeGreaterThan(0);

    await expect(client.readResource({ uri: "pptx://presentations/not-open" })).rejects.toMatchObject({ code: -32002 });
    await expect(client.readResource({ uri: "pptx://elsewhere" })).rejects.toMatchObject({ code: -32002 });

    await callTool("pptx_close_presentation", { presentation_id: presentationId });
  });

  it("reports invalid arguments as a validation error", async () => {
    const result = await callTool("pptx_add_slide", { presentation_id: "not-a-uuid" });

    expect(result.isError).toBe(true);
    expect(result.data.code).toBe("validation_error");
  });
});

describe("MCP server lifecycle", () => {
  it("exits when the client closes stdin", async () => {
    const child = spawn(process.execPath, SERVER_ARGS, { cwd: os.tmpdir(), stdio: ["pipe", "pipe", "pipe"] });
    const exited = new Promise<number | null>((resolve) => child.on("exit", (code) => resolve(code)));

    try {
      // Wait for the initialize reply so the Python bridge is known to be running.
      const ready = new Promise<void>((resolve) => child.stdout.once("data", () => resolve()));
      child.stdin.write(
        `${JSON.stringify({
          jsonrpc: "2.0",
          id: 1,
          method: "initialize",
          params: { protocolVersion: "2025-03-26", capabilities: {}, clientInfo: { name: "lifecycle-test", version: "0.0.0" } },
        })}\n`,
      );
      await ready;

      child.stdin.end();

      const timedOut = Symbol("timeout");
      const outcome = await Promise.race([
        exited,
        new Promise<typeof timedOut>((resolve) => setTimeout(() => resolve(timedOut), 5_000)),
      ]);
      expect(outcome, "server still running 5s after stdin closed").toBe(0);
    } finally {
      child.kill("SIGKILL");
    }
  }, 30_000);
});
