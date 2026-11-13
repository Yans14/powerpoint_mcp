import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { describe, expect, it } from "vitest";

import type { PythonBridgeClient } from "../../src/bridge/client.js";
import { createMcpServer, toInputSchema } from "../../src/server.js";
import { TOOL_DEFINITIONS, TOOL_MAP } from "../../src/tools/catalog.js";

const PRESENTATION_ID = "00000000-0000-4000-8000-000000000000";

async function connect(bridgeResult: unknown): Promise<Client> {
  const bridge = { call: async () => bridgeResult } as unknown as PythonBridgeClient;
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await createMcpServer(bridge).connect(serverTransport);
  const client = new Client({ name: "unit-test", version: "0.0.0" });
  await client.connect(clientTransport);
  return client;
}

describe("toInputSchema", () => {
  it("produces an object schema with properties for every tool", () => {
    for (const definition of TOOL_DEFINITIONS) {
      const schema = toInputSchema(definition.schema);
      expect(schema.type, definition.name).toBe("object");
      expect(schema.properties, definition.name).toBeDefined();
      expect(schema.$schema, definition.name).toBeUndefined();
    }
  });

  it("lists required parameters and leaves optional and defaulted ones out", () => {
    const addSlide = toInputSchema(TOOL_MAP.get("pptx_add_slide")!.schema);
    expect(addSlide.required).toEqual(["presentation_id", "layout_name"]);

    const addConnector = toInputSchema(TOOL_MAP.get("pptx_add_connector")!.schema);
    expect(addConnector.required).not.toContain("connector_type");
  });

  it("omits converter noise from uuid and integer fields", () => {
    const { properties } = toInputSchema(TOOL_MAP.get("pptx_add_slide")!.schema) as {
      properties: Record<string, Record<string, unknown>>;
    };

    expect(properties.presentation_id).toMatchObject({ type: "string", format: "uuid" });
    expect(properties.presentation_id.pattern).toBeUndefined();
    expect(properties.position).toMatchObject({ type: "integer", minimum: 1 });
    expect(properties.position.maximum).toBeUndefined();
  });
});

describe("pptx_get_slide_snapshot result", () => {
  it("returns the image once, as an image block, and keeps it out of the text block", async () => {
    const client = await connect({ slide_index: 1, mime_type: "image/jpeg", snapshot_base64: "QUJD", width_px: 640 });

    const result = (await client.callTool({
      name: "pptx_get_slide_snapshot",
      arguments: { presentation_id: PRESENTATION_ID, slide_index: 1 },
    })) as { content: Array<{ type: string; text?: string; data?: string; mimeType?: string }> };

    expect(result.content).toHaveLength(2);
    expect(JSON.parse(result.content[0].text ?? "")).toEqual({ slide_index: 1, mime_type: "image/jpeg", width_px: 640 });
    expect(result.content[1]).toEqual({ type: "image", data: "QUJD", mimeType: "image/jpeg" });
  });

  it("returns one image block per slide for a multi-slide snapshot", async () => {
    const client = await connect({
      presentation_id: PRESENTATION_ID,
      mime_type: "image/jpeg",
      width_px: 640,
      snapshots: [
        { slide_index: 3, snapshot_base64: "QUFB" },
        { slide_index: 1, snapshot_base64: "QkJC" },
      ],
    });

    const result = (await client.callTool({
      name: "pptx_get_slide_snapshot",
      arguments: { presentation_id: PRESENTATION_ID, slide_indices: [3, 1] },
    })) as { content: Array<{ type: string; text?: string; data?: string }> };

    expect(result.content).toHaveLength(3);
    expect(JSON.parse(result.content[0].text ?? "")).toEqual({
      presentation_id: PRESENTATION_ID,
      mime_type: "image/jpeg",
      width_px: 640,
      snapshots: [{ slide_index: 3 }, { slide_index: 1 }],
    });
    expect(result.content.slice(1).map((block) => block.data)).toEqual(["QUFB", "QkJC"]);
  });

  it.each([
    ["neither slide_index nor slide_indices", {}],
    ["both slide_index and slide_indices", { slide_index: 1, slide_indices: [1] }],
  ])("rejects %s", async (_label, selection) => {
    const client = await connect({});

    const result = (await client.callTool({
      name: "pptx_get_slide_snapshot",
      arguments: { presentation_id: PRESENTATION_ID, ...selection },
    })) as { isError?: boolean; content: Array<{ text?: string }> };

    expect(result.isError).toBe(true);
    expect(JSON.parse(result.content[0].text ?? "").code).toBe("validation_error");
  });

  it("leaves other tools' results untouched", async () => {
    const payload = { presentation_id: PRESENTATION_ID, slide_count: 0, slides: [] };
    const client = await connect(payload);

    const result = (await client.callTool({
      name: "pptx_get_presentation_state",
      arguments: { presentation_id: PRESENTATION_ID },
    })) as { content: Array<{ type: string; text?: string }> };

    expect(result.content).toHaveLength(1);
    expect(JSON.parse(result.content[0].text ?? "")).toEqual(payload);
  });
});
