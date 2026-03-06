import { zodToJsonSchema } from "zod-to-json-schema";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ListResourcesRequestSchema,
  ListToolsRequestSchema,
  ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { normalizeToolError, PythonBridgeClient } from "./bridge/client.js";
import { readResource, listResources } from "./resources/index.js";
import { TOOL_DEFINITIONS, TOOL_MAP } from "./tools/catalog.js";

export function createMcpServer(bridge: PythonBridgeClient): Server {
  const server = new Server(
    {
      name: "powerpoint-mcp-server",
      version: "0.1.0",
    },
    {
      capabilities: {
        tools: {},
        resources: {},
      },
    },
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: TOOL_DEFINITIONS.map((definition) => {
        const schema = zodToJsonSchema(definition.schema, {
          $refStrategy: "none",
          target: "jsonSchema7",
        }) as Record<string, unknown>;

        delete schema.$schema;

        return {
          name: definition.name,
          description: definition.description,
          inputSchema: schema,
          annotations: {
            readOnlyHint: !definition.mutating,
            destructiveHint: definition.name.includes("delete") || definition.name.includes("close"),
          },
        };
      }),
    };
  });

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const toolName = request.params.name;
    const definition = TOOL_MAP.get(toolName);

    if (!definition) {
      return {
        isError: true,
        content: [
          {
            type: "text",
            text: JSON.stringify({
              code: "not_found",
              message: `Unknown tool: ${toolName}`,
            }),
          },
        ],
      };
    }

    const parseResult = definition.schema.safeParse(request.params.arguments ?? {});
    if (!parseResult.success) {
      return {
        isError: true,
        content: [
          {
            type: "text",
            text: JSON.stringify({
              code: "validation_error",
              message: `Invalid parameters for ${toolName}`,
              details: parseResult.error.flatten(),
            }),
          },
        ],
      };
    }

    try {
      const result = await bridge.call(toolName, parseResult.data);
      const content: Array<Record<string, unknown>> = [
        {
          type: "text",
          text: JSON.stringify(result, null, 2),
        },
      ];

      if (toolName === "pptx_get_slide_snapshot" && typeof (result as { snapshot_base64?: unknown }).snapshot_base64 === "string") {
        content.push({
          type: "image",
          data: (result as { snapshot_base64: string }).snapshot_base64,
          mimeType: "image/jpeg",
        });
      }

      return { content };
    } catch (error) {
      const payload = normalizeToolError(error);
      return {
        isError: true,
        content: [
          {
            type: "text",
            text: JSON.stringify(payload),
          },
        ],
      };
    }
  });

  server.setRequestHandler(ListResourcesRequestSchema, async () => {
    const resources = await listResources(bridge);
    return { resources };
  });

  server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
    try {
      const data = await readResource(request.params.uri, bridge);
      return {
        contents: [
          {
            uri: request.params.uri,
            mimeType: "application/json",
            text: JSON.stringify(data, null, 2),
          },
        ],
      };
    } catch (error) {
      const payload = normalizeToolError(error);
      return {
        contents: [
          {
            uri: request.params.uri,
            mimeType: "application/json",
            text: JSON.stringify(payload),
          },
        ],
      };
    }
  });

  return server;
}
