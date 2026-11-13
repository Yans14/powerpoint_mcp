import { readFileSync } from "node:fs";

import { z } from "zod";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ErrorCode,
  GetPromptRequestSchema,
  ListPromptsRequestSchema,
  ListResourceTemplatesRequestSchema,
  ListResourcesRequestSchema,
  ListToolsRequestSchema,
  McpError,
  ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { normalizeToolError, PythonBridgeClient } from "./bridge/client.js";
import { BridgeErrorPayload } from "./bridge/types.js";
import { getPrompt, listPrompts } from "./prompts/index.js";
import { listResources, readResource, RESOURCE_TEMPLATES, ResourceNotFoundError } from "./resources/index.js";
import { BATCH_TOOL_NAME, prepareBatch } from "./tools/batch.js";
import { TOOL_DEFINITIONS, TOOL_MAP } from "./tools/catalog.js";
import { OUTPUT_SCHEMAS, toOutputSchema } from "./tools/outputs.js";
import {
  acceptsIncludeState,
  isToolInProfile,
  parseToolProfile,
  toolAnnotations,
  ToolProfile,
} from "./tools/traits.js";

/** JSON-RPC error code the MCP specification assigns to an unknown resource. */
const RESOURCE_NOT_FOUND = -32002;

/** Tools whose success changes the set of open presentations, and so the resource listing. */
const SESSION_LIFECYCLE_TOOLS = new Set([
  "pptx_create_presentation",
  "pptx_open_presentation",
  "pptx_close_presentation",
]);

// Repeated on every mutating tool, so the description is kept to one short sentence.
const INCLUDE_STATE_SCHEMA = {
  type: "boolean",
  description: "Also return the per-slide presentation_state, not just slide_count. Default false.",
};

const PACKAGE_VERSION = (
  JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as { version: string }
).version;

export interface McpServerOptions {
  /** Which tools to expose. Defaults to PPTX_TOOL_PROFILE, then "full". */
  profile?: ToolProfile;
}

/**
 * Convert a tool's zod schema into the JSON Schema advertised over MCP.
 * Uses zod 4's native converter; `io: "input"` keeps fields with defaults optional for callers.
 */
export function toInputSchema(schema: z.ZodTypeAny, options: { includeState?: boolean } = {}): Record<string, unknown> {
  const jsonSchema = z.toJSONSchema(schema, {
    target: "draft-7",
    io: "input",
    unrepresentable: "any",
    // Drop converter noise that costs tokens in every tools/list without informing the model:
    // the UUID regex (format already says it) and JS safe-integer bounds.
    override: ({ jsonSchema: node }) => {
      if (node.format === "uuid") {
        delete node.pattern;
      }
      if (node.type === "integer") {
        if (node.maximum === Number.MAX_SAFE_INTEGER) {
          delete node.maximum;
        }
        if (node.minimum === Number.MIN_SAFE_INTEGER) {
          delete node.minimum;
        }
      }
    },
  }) as Record<string, unknown>;

  delete jsonSchema.$schema;

  if (jsonSchema.type !== "object") {
    throw new Error("Tool input schema must be a JSON Schema object.");
  }

  if (options.includeState) {
    // Handled by the server for every mutating tool, so it is added here rather than in each schema.
    jsonSchema.properties = { ...(jsonSchema.properties as Record<string, unknown>), include_state: INCLUDE_STATE_SCHEMA };
  }

  return jsonSchema;
}

/**
 * Split a snapshot result into its metadata and its image payloads. Image bytes are returned as
 * image blocks only: repeating them as base64 in the text block would double their token cost.
 * Handles both response shapes: one slide (`snapshot_base64`) and several (`snapshots[]`).
 */
function splitSnapshotResult(result: unknown): { metadata: unknown; images: string[] } {
  if (result === null || typeof result !== "object") {
    return { metadata: result, images: [] };
  }

  const { snapshot_base64: single, snapshots, ...rest } = result as Record<string, unknown>;
  const images: string[] = [];
  const metadata: Record<string, unknown> = rest;

  if (typeof single === "string") {
    images.push(single);
  }

  if (Array.isArray(snapshots)) {
    metadata.snapshots = snapshots.map((entry: Record<string, unknown>) => {
      const { snapshot_base64: image, ...entryMetadata } = entry;
      if (typeof image === "string") {
        images.push(image);
      }
      return entryMetadata;
    });
  }

  return { metadata, images };
}

function errorResult(payload: BridgeErrorPayload | Record<string, unknown>) {
  return { isError: true, content: [{ type: "text", text: JSON.stringify(payload) }] };
}

/** Bridge failures while serving a resource become protocol errors, not resource contents. */
function toResourceError(error: unknown, uri: string): McpError {
  if (error instanceof ResourceNotFoundError) {
    return new McpError(RESOURCE_NOT_FOUND, error.message, { uri });
  }

  const payload = normalizeToolError(error);
  const code =
    payload.code === "not_found" || payload.code === "session_lost"
      ? RESOURCE_NOT_FOUND
      : payload.code === "validation_error"
        ? ErrorCode.InvalidParams
        : ErrorCode.InternalError;
  return new McpError(code, payload.message, { uri, ...payload });
}

export function createMcpServer(bridge: PythonBridgeClient, options: McpServerOptions = {}): Server {
  const profile = options.profile ?? parseToolProfile(process.env.PPTX_TOOL_PROFILE);
  const isToolAvailable = (name: string): boolean => TOOL_MAP.has(name) && isToolInProfile(name, profile);

  const server = new Server(
    {
      name: "powerpoint-mcp-server",
      version: PACKAGE_VERSION,
    },
    {
      capabilities: {
        tools: {},
        resources: { listChanged: true },
        prompts: {},
      },
    },
  );

  // Schemas never change at runtime; convert them once instead of on every tools/list.
  const advertisedTools = TOOL_DEFINITIONS.filter((definition) => isToolAvailable(definition.name)).map((definition) => {
    const annotations = toolAnnotations(definition);
    const outputSchema = OUTPUT_SCHEMAS[definition.name];
    return {
      name: definition.name,
      // The display title travels in annotations.title only: clients fall back to it, and sending
      // it twice per tool would be paid for in every tools/list.
      description: definition.description,
      inputSchema: toInputSchema(definition.schema, { includeState: acceptsIncludeState(definition) }),
      ...(outputSchema ? { outputSchema: toOutputSchema(outputSchema) } : {}),
      annotations,
    };
  });

  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: advertisedTools }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const toolName = request.params.name;
    const definition = TOOL_MAP.get(toolName);

    if (!definition || !isToolAvailable(toolName)) {
      return errorResult({
        code: "not_found",
        message: `Unknown tool: ${toolName}`,
        ...(definition
          ? {
              hint: `${toolName} is not part of the "${profile}" tool profile. Start the server with PPTX_TOOL_PROFILE=full to use it.`,
            }
          : {}),
      });
    }

    let rawArguments: Record<string, unknown> = request.params.arguments ?? {};
    let includeState: boolean | undefined;

    if (acceptsIncludeState(definition) && "include_state" in rawArguments) {
      const { include_state: requested, ...rest } = rawArguments;
      if (typeof requested !== "boolean") {
        return errorResult({
          code: "validation_error",
          message: `Invalid parameters for ${toolName}`,
          details: { fieldErrors: { include_state: ["Expected a boolean."] } },
        });
      }
      includeState = requested;
      rawArguments = rest;
    }

    const parseResult = definition.schema.safeParse(rawArguments);
    if (!parseResult.success) {
      return errorResult({
        code: "validation_error",
        message: `Invalid parameters for ${toolName}`,
        details: parseResult.error.flatten(),
      });
    }

    let bridgeParams = parseResult.data as Record<string, unknown>;

    if (toolName === BATCH_TOOL_NAME) {
      const prepared = prepareBatch(bridgeParams as unknown as Parameters<typeof prepareBatch>[0], isToolAvailable);
      if (!prepared.ok) {
        return errorResult({
          code: "validation_error",
          message: `Invalid operations for ${toolName}. Nothing was executed.`,
          details: { operations: prepared.problems },
          ...(prepared.hint ? { hint: prepared.hint } : {}),
        });
      }
      bridgeParams = prepared.params;
    }

    if (includeState !== undefined) {
      bridgeParams = { ...bridgeParams, include_state: includeState };
    }

    try {
      const result = await bridge.call(toolName, bridgeParams);
      const { metadata, images } =
        toolName === "pptx_get_slide_snapshot" ? splitSnapshotResult(result) : { metadata: result, images: [] };

      // Compact JSON: the reader is a model, and indentation is paid for in tokens on every call.
      const content: Array<Record<string, unknown>> = [{ type: "text", text: JSON.stringify(metadata) }];

      // One image block per slide, in the order listed in the metadata.
      for (const data of images) {
        content.push({ type: "image", data, mimeType: "image/jpeg" });
      }

      if (SESSION_LIFECYCLE_TOOLS.has(toolName)) {
        // Best effort: a client that went away must not turn a successful call into a failure.
        void server.sendResourceListChanged().catch(() => undefined);
      }

      if (toolName in OUTPUT_SCHEMAS) {
        return { content, structuredContent: result as Record<string, unknown> };
      }
      return { content };
    } catch (error) {
      return errorResult(normalizeToolError(error));
    }
  });

  server.setRequestHandler(ListResourcesRequestSchema, async () => {
    try {
      return { resources: await listResources(bridge) };
    } catch (error) {
      throw toResourceError(error, "pptx://presentations");
    }
  });

  server.setRequestHandler(ListResourceTemplatesRequestSchema, async () => ({
    resourceTemplates: RESOURCE_TEMPLATES,
  }));

  server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
    const { uri } = request.params;
    try {
      return { contents: [await readResource(uri, bridge)] };
    } catch (error) {
      throw toResourceError(error, uri);
    }
  });

  server.setRequestHandler(ListPromptsRequestSchema, async () => ({ prompts: listPrompts(isToolAvailable) }));

  server.setRequestHandler(GetPromptRequestSchema, async (request) => {
    const { name, arguments: promptArguments } = request.params;
    let prompt;
    try {
      prompt = getPrompt(name, promptArguments ?? {}, isToolAvailable);
    } catch (error) {
      throw new McpError(ErrorCode.InvalidParams, error instanceof Error ? error.message : String(error));
    }
    if (!prompt) {
      throw new McpError(ErrorCode.InvalidParams, `Unknown prompt: ${name}`);
    }
    return prompt;
  });

  return server;
}
