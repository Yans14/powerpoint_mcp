import { fileURLToPath } from "node:url";

import { afterEach, describe, expect, it } from "vitest";

import { PythonBridgeClient } from "../../src/bridge/client.js";
import { BridgeClientError } from "../../src/bridge/types.js";

const FAKE_BRIDGE = fileURLToPath(new URL("../fixtures/fake-bridge.mjs", import.meta.url));

const clients: PythonBridgeClient[] = [];

function createClient(options: { pythonBin?: string; callTimeoutMs?: number } = {}): PythonBridgeClient {
  const client = new PythonBridgeClient({
    pythonBin: process.execPath,
    bridgeScriptPath: FAKE_BRIDGE,
    callTimeoutMs: 5_000,
    ...options,
  });
  clients.push(client);
  return client;
}

async function failure(promise: Promise<unknown>): Promise<BridgeClientError> {
  try {
    await promise;
  } catch (error) {
    expect(error).toBeInstanceOf(BridgeClientError);
    return error as BridgeClientError;
  }
  throw new Error("Expected the call to fail.");
}

async function bridgePid(client: PythonBridgeClient): Promise<number> {
  return (await client.call<{ pid: number }>("pptx_get_engine_info", {})).pid;
}

afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.close()));
});

describe("PythonBridgeClient", () => {
  it("returns results and maps bridge errors", async () => {
    const client = createClient();

    expect(await client.call("echo", { value: 1 })).toEqual({ method: "echo", params: { value: 1 } });

    const error = await failure(client.call("fail", {}));
    expect(error.code).toBe("not_found");
    expect(error.message).toBe("nope");
  });

  it("starts a single process for concurrent first calls", async () => {
    const client = createClient();

    const [first, second] = await Promise.all([bridgePid(client), bridgePid(client)]);

    expect(first).toBe(second);
  });

  it("reports a missing interpreter with an actionable error and can be retried", async () => {
    const client = createClient({ pythonBin: "pptx-mcp-no-such-interpreter" });

    for (let attempt = 0; attempt < 2; attempt += 1) {
      const error = await failure(client.start());
      expect(error.code).toBe("dependency_missing");
      expect(error.message).toContain("PPTX_PYTHON_BIN");
    }
  });

  it("ignores non-protocol output on the bridge's stdout", async () => {
    const client = createClient();

    expect(await client.call("noise", {})).toEqual({ ok: true });
  });

  it("times out a stuck call, restarts the bridge and reports lost sessions", async () => {
    const client = createClient({ callTimeoutMs: 300 });
    const firstPid = await bridgePid(client);
    const { presentation_id: presentationId } = await client.call<{ presentation_id: string }>(
      "pptx_create_presentation",
      {},
    );

    const timeout = await failure(client.call("hang", {}));
    expect(timeout.code).toBe("timeout");
    expect(timeout.details).toMatchObject({ method: "hang", timeout_ms: 300 });

    const lost = await failure(client.call("echo", { presentation_id: presentationId }));
    expect(lost.code).toBe("session_lost");
    expect(lost.details).toMatchObject({ presentation_id: presentationId });

    // Unrelated calls go through on a fresh process.
    expect(await bridgePid(client)).not.toBe(firstPid);
  });

  it("fails calls queued behind a stuck one instead of leaving them pending", async () => {
    const client = createClient({ callTimeoutMs: 300 });
    await client.start();

    const stuck = failure(client.call("hang", {}));
    const queued = failure(client.call("hang", {}));

    expect((await stuck).code).toBe("timeout");
    expect((await queued).code).toBe("bridge_unavailable");
  });

  it("reports a crash, marks its sessions lost and recovers on the next call", async () => {
    const client = createClient();
    const { presentation_id: presentationId } = await client.call<{ presentation_id: string }>(
      "pptx_create_presentation",
      {},
    );

    const crash = await failure(client.call("crash", {}));
    expect(crash.code).toBe("bridge_unavailable");
    expect(crash.details).toMatchObject({ exit_code: 3 });
    expect(crash.details?.stderr_tail).toContain("fake bridge crashing on request");

    expect((await failure(client.call("echo", { presentation_id: presentationId }))).code).toBe("session_lost");
    expect(await client.call("echo", {})).toEqual({ method: "echo", params: {} });
  });

  it("does not report closed sessions as lost", async () => {
    const client = createClient();
    const { presentation_id: presentationId } = await client.call<{ presentation_id: string }>(
      "pptx_create_presentation",
      {},
    );
    await client.call("pptx_close_presentation", { presentation_id: presentationId });

    await failure(client.call("crash", {}));

    // The fake bridge echoes; the real one would answer not_found for a closed session.
    expect(await client.call("echo", { presentation_id: presentationId })).toMatchObject({ method: "echo" });
  });
});
