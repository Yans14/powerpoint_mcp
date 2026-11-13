import { readFileSync } from "node:fs";

import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { ResourceListChangedNotificationSchema } from "@modelcontextprotocol/sdk/types.js";
import { describe, expect, it } from "vitest";

import type { PythonBridgeClient } from "../../src/bridge/client.js";
import { BridgeClientError } from "../../src/bridge/types.js";
import { createMcpServer } from "../../src/server.js";
import { prepareBatch } from "../../src/tools/batch.js";
import { TOOL_DEFINITIONS, TOOL_MAP } from "../../src/tools/catalog.js";
import { OUTPUT_SCHEMAS } from "../../src/tools/outputs.js";
import {
  isAgentEnabled,
  parseToolProfile,
  toolAnnotations,
  toolTitle,
  ToolProfile,
  TRAIT_SETS,
} from "../../src/tools/traits.js";

const PRESENTATION_ID = "00000000-0000-4000-8000-000000000000";
const OTHER_ID = "11111111-1111-4111-8111-111111111111";

type BridgeCall = { method: string; params: Record<string, unknown> };
type BridgeHandler = (method: string, params: Record<string, unknown>) => unknown;

async function connect(
  handler: BridgeHandler = () => ({}),
  profile: ToolProfile = "full",
  agentEnabled = true,
) {
  const calls: BridgeCall[] = [];
  const bridge = {
    call: async (method: string, params: Record<string, unknown>) => {
      calls.push({ method, params });
      return handler(method, params);
    },
  } as unknown as PythonBridgeClient;

  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await createMcpServer(bridge, { profile, agentEnabled }).connect(serverTransport);
  const client = new Client({ name: "unit-test", version: "0.0.0" });
  await client.connect(clientTransport);
  return { client, calls };
}

type ToolResult = { isError?: boolean; structuredContent?: unknown; content: Array<{ type: string; text?: string }> };

async function callTool(client: Client, name: string, args: Record<string, unknown>) {
  const result = (await client.callTool({ name, arguments: args })) as ToolResult;
  return { ...result, data: JSON.parse(result.content[0].text ?? "null") };
}

async function rejection(promise: Promise<unknown>): Promise<{ code?: number; message: string }> {
  try {
    await promise;
  } catch (error) {
    return error as { code?: number; message: string };
  }
  throw new Error("Expected the request to be rejected.");
}

describe("server identity", () => {
  it("reports the version from package.json", async () => {
    const { client } = await connect();
    const { version } = JSON.parse(readFileSync(new URL("../../package.json", import.meta.url), "utf8"));

    expect(client.getServerVersion()).toMatchObject({ name: "powerpoint-mcp-server", version });
  });
});

describe("tool traits", () => {
  const names = new Set(TOOL_DEFINITIONS.map((definition) => definition.name));
  const mutating = new Set(TOOL_DEFINITIONS.filter((definition) => definition.mutating).map((definition) => definition.name));

  it("only reference tools that exist", () => {
    for (const [trait, tools] of Object.entries(TRAIT_SETS)) {
      for (const name of tools) {
        expect(names.has(name), `${trait}: ${name}`).toBe(true);
      }
    }
  });

  it("only mark mutating tools as destructive or as idempotent mutations", () => {
    for (const name of [...TRAIT_SETS.destructive, ...TRAIT_SETS.idempotentMutation]) {
      expect(mutating.has(name), name).toBe(true);
    }
  });

  it("derive annotations from the declared traits, not from tool names", () => {
    expect(toolAnnotations(TOOL_MAP.get("pptx_save_presentation")!)).toEqual({
      title: "Save Presentation",
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
      openWorldHint: false,
    });
    expect(toolAnnotations(TOOL_MAP.get("pptx_get_slide")!)).toMatchObject({
      readOnlyHint: true,
      destructiveHint: false,
      idempotentHint: true,
    });
    expect(toolAnnotations(TOOL_MAP.get("pptx_add_slide")!)).toMatchObject({
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
    });
    expect(toolAnnotations(TOOL_MAP.get("pptx_agent_start")!)).toMatchObject({
      idempotentHint: false,
      openWorldHint: true,
    });
    expect(toolTitle("pptx_get_slide_xml")).toBe("Get Slide XML");
  });

  it("are advertised in full for every tool", async () => {
    const { client } = await connect();
    const { tools } = await client.listTools();

    expect(tools).toHaveLength(TOOL_DEFINITIONS.length);
    for (const tool of tools) {
      expect(tool.annotations?.title, tool.name).toBeTruthy();
      expect(Object.keys(tool.annotations ?? {}).sort(), tool.name).toEqual([
        "destructiveHint",
        "idempotentHint",
        "openWorldHint",
        "readOnlyHint",
        "title",
      ]);
    }
  });
});

describe("embedded agent switch", () => {
  it("hides the agent tools unless the agent is enabled", async () => {
    const { client, calls } = await connect(() => ({}), "full", false);

    const { tools } = await client.listTools();
    const agentTools = TOOL_DEFINITIONS.filter((definition) => definition.category === "Agent");
    expect(agentTools.length).toBeGreaterThan(0);
    expect(tools).toHaveLength(TOOL_DEFINITIONS.length - agentTools.length);
    expect(tools.some((tool) => tool.name.startsWith("pptx_agent_"))).toBe(false);

    const hidden = await callTool(client, "pptx_agent_status", { task_id: "t" });
    expect(hidden.data.code).toBe("not_found");
    expect(hidden.data.hint).toContain("PPTX_ENABLE_AGENT=1");
    expect(calls).toEqual([]);
  });

  it("reads the switch like the Python side does", () => {
    expect(isAgentEnabled(undefined)).toBe(false);
    expect(isAgentEnabled("")).toBe(false);
    expect(isAgentEnabled("0")).toBe(false);
    expect(isAgentEnabled("1")).toBe(true);
    expect(isAgentEnabled(" True ")).toBe(true);
  });
});

describe("tool profiles", () => {
  it("parses the profile name and rejects unknown ones", () => {
    expect(parseToolProfile(undefined)).toBe("full");
    expect(parseToolProfile(" Core ")).toBe("core");
    expect(() => parseToolProfile("minimal")).toThrow(/PPTX_TOOL_PROFILE/);
  });

  it("exposes only core tools and prompts in the core profile", async () => {
    const { client, calls } = await connect(() => ({}), "core");

    const { tools } = await client.listTools();
    expect(tools.map((tool) => tool.name).sort()).toEqual([...TRAIT_SETS.core].sort());

    const hidden = await callTool(client, "pptx_get_slide_xml", { presentation_id: PRESENTATION_ID, slide_index: 1 });
    expect(hidden.isError).toBe(true);
    expect(hidden.data.code).toBe("not_found");
    expect(hidden.data.hint).toContain("PPTX_TOOL_PROFILE=full");
    expect(calls).toEqual([]);

    const { prompts } = await client.listPrompts();
    expect(prompts.map((prompt) => prompt.name)).toEqual(["create_deck_from_outline"]);
  });
});

describe("tool results", () => {
  it("offers include_state on mutating tools only, and forwards it to the bridge", async () => {
    const { client, calls } = await connect(() => ({ success: true, slide_count: 1 }));
    const { tools } = await client.listTools();
    const properties = (name: string) => Object.keys(tools.find((tool) => tool.name === name)?.inputSchema.properties ?? {});

    expect(properties("pptx_add_slide")).toContain("include_state");
    expect(properties("pptx_get_slide")).not.toContain("include_state");
    expect(properties("pptx_agent_execute")).not.toContain("include_state");

    await callTool(client, "pptx_add_slide", { presentation_id: PRESENTATION_ID, layout_name: "Blank", include_state: true });
    await callTool(client, "pptx_add_slide", { presentation_id: PRESENTATION_ID, layout_name: "Blank" });

    expect(calls.map((call) => call.params)).toEqual([
      { presentation_id: PRESENTATION_ID, layout_name: "Blank", include_state: true },
      { presentation_id: PRESENTATION_ID, layout_name: "Blank" },
    ]);
  });

  it("rejects include_state where it does not apply or is not a boolean", async () => {
    const { client, calls } = await connect();

    const onRead = await callTool(client, "pptx_get_slide", {
      presentation_id: PRESENTATION_ID,
      slide_index: 1,
      include_state: true,
    });
    const notBoolean = await callTool(client, "pptx_add_slide", {
      presentation_id: PRESENTATION_ID,
      layout_name: "Blank",
      include_state: "yes",
    });

    expect(onRead.data.code).toBe("validation_error");
    expect(notBoolean.data.code).toBe("validation_error");
    expect(calls).toEqual([]);
  });

  it("returns compact JSON text", async () => {
    const { client } = await connect(() => ({ a: 1, b: [1, 2] }));

    const result = await callTool(client, "pptx_get_slide", { presentation_id: PRESENTATION_ID, slide_index: 1 });

    expect(result.content[0].text).toBe('{"a":1,"b":[1,2]}');
    expect(result.structuredContent).toBeUndefined();
  });

  it("advertises an output schema and returns structured content for the tools that declare one", async () => {
    const state = {
      presentation_id: PRESENTATION_ID,
      engine: "OOXML",
      slide_count: 1,
      slides: [{ index: 1, title: "Intro", layout: "Title Slide", shape_count: 2 }],
    };
    const { client } = await connect(() => state);

    const { tools } = await client.listTools();
    const withOutputSchema = tools.filter((tool) => tool.outputSchema).map((tool) => tool.name);
    expect(withOutputSchema.sort()).toEqual(Object.keys(OUTPUT_SCHEMAS).sort());

    // The SDK client validates structuredContent against the advertised schema.
    const result = await callTool(client, "pptx_get_presentation_state", { presentation_id: PRESENTATION_ID });
    expect(result.structuredContent).toEqual(state);
    expect(result.data).toEqual(state);
  });

  it("passes the bridge's hint through on errors", async () => {
    const { client } = await connect(() => {
      throw new BridgeClientError({ code: "not_found", message: "Layout 'X' not found.", hint: "Call pptx_get_layouts." });
    });

    const result = await callTool(client, "pptx_add_slide", { presentation_id: PRESENTATION_ID, layout_name: "X" });

    expect(result.isError).toBe(true);
    expect(result.data).toEqual({ code: "not_found", message: "Layout 'X' not found.", hint: "Call pptx_get_layouts." });
  });
});

describe("pptx_batch", () => {
  const available = (name: string) => TOOL_MAP.has(name);

  it("validates each operation with its own schema and injects the presentation id", () => {
    const prepared = prepareBatch(
      {
        presentation_id: PRESENTATION_ID,
        operations: [
          { tool: "pptx_add_slide", arguments: { layout_name: "Blank", include_state: true } },
          { tool: "pptx_get_slide", arguments: { slide_index: 1, presentation_id: PRESENTATION_ID } },
        ],
      },
      available,
    );

    expect(prepared).toEqual({
      ok: true,
      params: {
        presentation_id: PRESENTATION_ID,
        operations: [
          { tool: "pptx_add_slide", arguments: { presentation_id: PRESENTATION_ID, layout_name: "Blank" } },
          { tool: "pptx_get_slide", arguments: { presentation_id: PRESENTATION_ID, slide_index: 1 } },
        ],
      },
    });
  });

  it("reports every invalid operation by index", () => {
    const prepared = prepareBatch(
      {
        presentation_id: PRESENTATION_ID,
        atomic: false,
        operations: [
          { tool: "pptx_add_slide", arguments: { layout_name: "Blank" } },
          { tool: "pptx_add_slide", arguments: {} },
          { tool: "pptx_no_such_tool" },
          { tool: "pptx_save_presentation", arguments: { output_path: "/tmp/x.pptx" } },
          { tool: "pptx_get_engine_info" },
          { tool: "pptx_agent_status", arguments: { task_id: "t" } },
          { tool: "pptx_get_slide", arguments: { slide_index: 1, presentation_id: OTHER_ID } },
        ],
      },
      available,
    );

    expect(prepared.ok).toBe(false);
    if (prepared.ok) {
      return;
    }
    expect(prepared.problems.map((problem) => problem.operation_index)).toEqual([1, 2, 3, 4, 5, 6]);
    expect(prepared.hint).toContain("pptx_save_presentation");
  });

  it("sends nothing to the bridge when an operation is invalid", async () => {
    const { client, calls } = await connect();

    const result = await callTool(client, "pptx_batch", {
      presentation_id: PRESENTATION_ID,
      operations: [{ tool: "pptx_add_slide", arguments: {} }],
    });

    expect(result.isError).toBe(true);
    expect(result.data.code).toBe("validation_error");
    expect(result.data.details.operations[0]).toMatchObject({ operation_index: 0, tool: "pptx_add_slide" });
    expect(calls).toEqual([]);
  });

  it("sends a valid batch as one bridge call", async () => {
    const { client, calls } = await connect(() => ({ success: true, executed: 1, results: [], slide_count: 1 }));

    const result = await callTool(client, "pptx_batch", {
      presentation_id: PRESENTATION_ID,
      atomic: false,
      operations: [{ tool: "pptx_add_slide", arguments: { layout_name: "Blank" } }],
    });

    expect(result.isError).toBeFalsy();
    expect(calls).toEqual([
      {
        method: "pptx_batch",
        params: {
          presentation_id: PRESENTATION_ID,
          atomic: false,
          operations: [{ tool: "pptx_add_slide", arguments: { presentation_id: PRESENTATION_ID, layout_name: "Blank" } }],
        },
      },
    ]);
  });
});

describe("resources", () => {
  const presentations = {
    presentations: [
      { presentation_id: PRESENTATION_ID, original_path: "/decks/q3-review.pptx" },
      { presentation_id: OTHER_ID, original_path: "" },
    ],
  };

  it("lists open presentations with a single bridge call", async () => {
    const { client, calls } = await connect(() => presentations);

    const { resources } = await client.listResources();

    expect(calls.map((call) => call.method)).toEqual(["pptx_list_open_presentations"]);
    expect(resources.map((resource) => [resource.uri, resource.title])).toEqual([
      ["pptx://presentations", "Open presentations"],
      [`pptx://presentations/${PRESENTATION_ID}`, "q3-review.pptx"],
      [`pptx://presentations/${OTHER_ID}`, "Untitled presentation"],
    ]);
  });

  it("describes per-presentation resources as templates", async () => {
    const { client } = await connect();

    const { resourceTemplates } = await client.listResourceTemplates();

    expect(resourceTemplates.map((template) => template.uriTemplate)).toContain(
      "pptx://presentations/{presentation_id}/slides/{slide_index}/snapshot",
    );
    expect(resourceTemplates.find((template) => template.name === "slide-snapshot")?.mimeType).toBe("image/jpeg");
  });

  it("serves JSON resources as text and snapshots as image blobs", async () => {
    const { client, calls } = await connect((method) =>
      method === "pptx_get_slide_snapshot" ? { mime_type: "image/jpeg", snapshot_base64: "QUJD" } : { layouts: [] },
    );

    const layouts = await client.readResource({ uri: `pptx://presentations/${PRESENTATION_ID}/layouts` });
    const snapshot = await client.readResource({ uri: `pptx://presentations/${PRESENTATION_ID}/slides/2/snapshot` });

    expect(layouts.contents).toEqual([
      { uri: `pptx://presentations/${PRESENTATION_ID}/layouts`, mimeType: "application/json", text: '{"layouts":[]}' },
    ]);
    expect(snapshot.contents).toEqual([
      { uri: `pptx://presentations/${PRESENTATION_ID}/slides/2/snapshot`, mimeType: "image/jpeg", blob: "QUJD" },
    ]);
    expect(calls[1]).toEqual({
      method: "pptx_get_slide_snapshot",
      params: { presentation_id: PRESENTATION_ID, slide_index: 2 },
    });
  });

  it.each([
    "pptx://nothing",
    "file:///etc/passwd",
    `pptx://presentations/${PRESENTATION_ID}/unknown`,
    `pptx://presentations/${PRESENTATION_ID}/slides/abc`,
    `pptx://presentations/${PRESENTATION_ID}/slides/0`,
    `pptx://presentations/${PRESENTATION_ID}/slides/1/extra`,
  ])("rejects the unknown URI %s with a resource-not-found error", async (uri) => {
    const { client, calls } = await connect();

    const error = await rejection(client.readResource({ uri }));

    expect(error.code).toBe(-32002);
    expect(calls).toEqual([]);
  });

  it("turns a bridge not_found into a protocol error instead of resource content", async () => {
    const { client } = await connect(() => {
      throw new BridgeClientError({ code: "not_found", message: "Presentation session not found." });
    });

    const error = await rejection(client.readResource({ uri: `pptx://presentations/${PRESENTATION_ID}` }));

    expect(error.code).toBe(-32002);
    expect(error.message).toContain("Presentation session not found.");
  });

  it("notifies that the resource list changed when a presentation is opened or closed", async () => {
    const { client } = await connect(() => ({ success: true, presentation_id: PRESENTATION_ID }));
    let notifications = 0;
    client.setNotificationHandler(ResourceListChangedNotificationSchema, () => {
      notifications += 1;
    });

    await callTool(client, "pptx_create_presentation", {});
    await callTool(client, "pptx_get_presentation_state", { presentation_id: PRESENTATION_ID }).catch(() => undefined);
    await callTool(client, "pptx_close_presentation", { presentation_id: PRESENTATION_ID });
    await new Promise((resolve) => setTimeout(resolve, 20));

    expect(notifications).toBe(2);
  });
});

describe("prompts", () => {
  it("lists the guided workflows with their arguments", async () => {
    const { client } = await connect();

    const { prompts } = await client.listPrompts();

    expect(prompts.map((prompt) => prompt.name)).toEqual(["create_deck_from_outline", "apply_template_to_deck"]);
    expect(prompts[0].arguments?.filter((argument) => argument.required).map((argument) => argument.name)).toEqual([
      "outline",
      "output_path",
    ]);
  });

  it("renders a prompt with the caller's arguments", async () => {
    const { client } = await connect();

    const prompt = await client.getPrompt({
      name: "create_deck_from_outline",
      arguments: { outline: "1. Intro\n2. Results", output_path: "/out/deck.pptx", template_path: "/tpl/brand.potx" },
    });

    const text = (prompt.messages[0].content as { text: string }).text;
    expect(prompt.messages).toHaveLength(1);
    expect(text).toContain("1. Intro\n2. Results");
    expect(text).toContain('template_path "/tpl/brand.potx"');
    expect(text).toContain('output_path "/out/deck.pptx"');
  });

  it("only names tools that exist", async () => {
    const { client } = await connect();
    const rendered = await Promise.all([
      client.getPrompt({ name: "create_deck_from_outline", arguments: { outline: "x", output_path: "/o.pptx" } }),
      client.getPrompt({ name: "apply_template_to_deck", arguments: { file_path: "/a.pptx", template_path: "/t.potx" } }),
    ]);

    for (const prompt of rendered) {
      const text = (prompt.messages[0].content as { text: string }).text;
      for (const toolName of text.match(/pptx_[a-z_]+/g) ?? []) {
        expect(toolName === "pptx_" || TOOL_MAP.has(toolName), toolName).toBe(true);
      }
    }
  });

  it("rejects unknown prompts and missing required arguments", async () => {
    const { client } = await connect();

    const unknown = await rejection(client.getPrompt({ name: "no_such_prompt" }));
    const missing = await rejection(client.getPrompt({ name: "create_deck_from_outline", arguments: { outline: "x" } }));

    expect(unknown.code).toBe(-32602);
    expect(missing.code).toBe(-32602);
    expect(missing.message).toContain("output_path");
  });
});
