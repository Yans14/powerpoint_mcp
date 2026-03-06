import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { PythonBridgeClient } from "./bridge/client.js";
import { createMcpServer } from "./server.js";

async function main(): Promise<void> {
  const bridge = new PythonBridgeClient();
  await bridge.start();

  const server = createMcpServer(bridge);
  const transport = new StdioServerTransport();
  await server.connect(transport);

  const shutdown = async (): Promise<void> => {
    await bridge.close();
    process.exit(0);
  };

  process.on("SIGINT", () => {
    void shutdown();
  });

  process.on("SIGTERM", () => {
    void shutdown();
  });
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.stack ?? error.message : String(error);
  process.stderr.write(`Fatal startup error: ${message}\n`);
  process.exit(1);
});
