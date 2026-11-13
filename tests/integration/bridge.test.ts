import { readdir, mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { PythonBridgeClient } from "../../src/bridge/client.js";
import { BridgeClientError } from "../../src/bridge/types.js";

async function waitFor(condition: () => boolean, timeoutMs: number): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (!condition()) {
    if (Date.now() > deadline) {
      throw new Error("Condition not met in time.");
    }
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
}

describe("PythonBridgeClient with the real Python bridge", () => {
  let workspaceRoot: string;
  let client: PythonBridgeClient;
  const previousWorkspaceDir = process.env.PPTX_WORKSPACE_DIR;

  beforeAll(async () => {
    workspaceRoot = await mkdtemp(path.join(os.tmpdir(), "pptx-mcp-ws-"));
    process.env.PPTX_WORKSPACE_DIR = workspaceRoot;
    client = new PythonBridgeClient();
  });

  afterAll(async () => {
    await client.close();
    if (previousWorkspaceDir === undefined) {
      delete process.env.PPTX_WORKSPACE_DIR;
    } else {
      process.env.PPTX_WORKSPACE_DIR = previousWorkspaceDir;
    }
    await rm(workspaceRoot, { recursive: true, force: true });
  });

  it("recovers from a killed bridge: lost sessions are reported and leftovers are swept", async () => {
    const { presentation_id: lostId } = await client.call<{ presentation_id: string }>("pptx_create_presentation", {});
    const [abandonedWorkspace] = await readdir(workspaceRoot);
    expect(abandonedWorkspace).toMatch(/^bridge-\d+-/);

    // Simulate a hard crash: SIGKILL gives the bridge no chance to clean up.
    const internals = client as unknown as { child?: { kill(signal: string): boolean } };
    const crashed = internals.child;
    crashed?.kill("SIGKILL");
    await waitFor(() => internals.child !== crashed, 5_000);

    const lost = await client.call("pptx_get_presentation_state", { presentation_id: lostId }).catch((error) => error);
    expect(lost).toBeInstanceOf(BridgeClientError);
    expect((lost as BridgeClientError).code).toBe("session_lost");

    // The next call starts a fresh bridge, which removes the dead one's working files.
    const { presentation_id: freshId } = await client.call<{ presentation_id: string }>("pptx_create_presentation", {});
    expect(freshId).not.toBe(lostId);

    const workspaces = await readdir(workspaceRoot);
    expect(workspaces).toHaveLength(1);
    expect(workspaces).not.toContain(abandonedWorkspace);

    await client.close();
    expect(await readdir(workspaceRoot)).toEqual([]);
  }, 30_000);
});
