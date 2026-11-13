import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { PythonBridgeClient } from "./bridge/client.js";
import { BridgeClientError } from "./bridge/types.js";
import { createMcpServer } from "./server.js";

const SHUTDOWN_GRACE_MS = 3000;

async function main(): Promise<void> {
  const bridge = new PythonBridgeClient();
  await bridge.start();

  const server = createMcpServer(bridge);
  const transport = new StdioServerTransport();
  await server.connect(transport);

  let shuttingDown = false;
  const shutdown = async (): Promise<void> => {
    if (shuttingDown) {
      return;
    }
    shuttingDown = true;

    // Never let a stuck bridge keep the server alive after the client is gone.
    setTimeout(() => process.exit(0), SHUTDOWN_GRACE_MS).unref();

    try {
      await bridge.close();
    } finally {
      process.exit(0);
    }
  };

  process.on("SIGINT", () => {
    void shutdown();
  });

  process.on("SIGTERM", () => {
    void shutdown();
  });

  // The MCP client signals the end of a stdio session by closing stdin; without this the
  // Python bridge child keeps the process alive forever.
  process.stdin.on("end", () => {
    void shutdown();
  });

  process.stdin.on("close", () => {
    void shutdown();
  });

  server.onclose = () => {
    void shutdown();
  };
}

main().catch((error: unknown) => {
  let message: string;
  if (error instanceof BridgeClientError) {
    // Expected failure modes (missing interpreter or packages): the message is the diagnosis, a stack adds noise.
    message = error.details ? `${error.message}\n${JSON.stringify(error.details, null, 2)}` : error.message;
  } else {
    message = error instanceof Error ? error.stack ?? error.message : String(error);
  }
  process.stderr.write(`Fatal startup error: ${message}\n`);
  process.exit(1);
});
