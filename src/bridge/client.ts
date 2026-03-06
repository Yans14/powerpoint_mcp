import { createInterface } from "node:readline";
import { ChildProcessWithoutNullStreams, spawn } from "node:child_process";
import path from "node:path";
import { randomUUID } from "node:crypto";

import {
  BridgeClientError,
  BridgeErrorPayload,
  BridgeFailureResponse,
  BridgeRequest,
  BridgeResponse,
  BridgeSuccessResponse,
} from "./types.js";

interface PendingRequest {
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
  method: string;
}

export class PythonBridgeClient {
  private child?: ChildProcessWithoutNullStreams;
  private readonly pending = new Map<string, PendingRequest>();
  private readonly pythonBin: string;
  private readonly bridgeScriptPath: string;

  public constructor() {
    this.pythonBin = process.env.PPTX_PYTHON_BIN ?? "python3";
    this.bridgeScriptPath = process.env.PPTX_PY_BRIDGE ?? path.resolve(process.cwd(), "python/bridge.py");
  }

  public async start(): Promise<void> {
    if (this.child) {
      return;
    }

    this.child = spawn(this.pythonBin, [this.bridgeScriptPath], {
      stdio: ["pipe", "pipe", "pipe"],
      env: {
        ...process.env,
        PYTHONUNBUFFERED: "1",
      },
    });

    const stdoutRl = createInterface({ input: this.child.stdout });
    stdoutRl.on("line", (line) => this.handleResponseLine(line));

    const stderrRl = createInterface({ input: this.child.stderr });
    stderrRl.on("line", (line) => {
      process.stderr.write(`[python-bridge] ${line}\n`);
    });

    this.child.on("exit", (code, signal) => {
      const message = `Python bridge exited (code=${code ?? "null"}, signal=${signal ?? "null"}).`;
      for (const [id, request] of this.pending.entries()) {
        request.reject(new Error(`${message} Pending method: ${request.method}`));
        this.pending.delete(id);
      }
      this.child = undefined;
    });

    // Light startup probe so bridge errors fail fast.
    await this.call("pptx_get_engine_info", {});
  }

  public async call<TResult = unknown>(method: string, params: Record<string, unknown>): Promise<TResult> {
    await this.start();

    if (!this.child || !this.child.stdin.writable) {
      throw new Error("Python bridge process is not available.");
    }

    const id = randomUUID();
    const payload: BridgeRequest = { id, method, params };

    const responsePromise = new Promise<TResult>((resolve, reject) => {
      this.pending.set(id, { resolve: resolve as (value: unknown) => void, reject, method });
    });

    this.child.stdin.write(`${JSON.stringify(payload)}\n`);
    return responsePromise;
  }

  public async close(): Promise<void> {
    if (!this.child) {
      return;
    }

    try {
      await this.call("__shutdown__", {});
    } catch {
      // Ignore shutdown errors and terminate process.
    }

    this.child.kill();
    this.child = undefined;
  }

  private handleResponseLine(line: string): void {
    if (!line.trim()) {
      return;
    }

    let parsed: BridgeResponse;
    try {
      parsed = JSON.parse(line) as BridgeResponse;
    } catch (error) {
      process.stderr.write(`[python-bridge] Invalid JSON response: ${line}\n`);
      return;
    }

    const pending = this.pending.get(parsed.id);
    if (!pending) {
      process.stderr.write(`[python-bridge] Dropped unmatched response id=${parsed.id}.\n`);
      return;
    }

    this.pending.delete(parsed.id);

    if (this.isFailure(parsed)) {
      pending.reject(new BridgeClientError(parsed.error));
      return;
    }

    pending.resolve((parsed as BridgeSuccessResponse).result);
  }

  private isFailure(response: BridgeResponse): response is BridgeFailureResponse {
    return "error" in response;
  }
}

export function normalizeToolError(error: unknown): BridgeErrorPayload {
  if (error instanceof BridgeClientError) {
    return {
      code: error.code,
      message: error.message,
      details: error.details,
    };
  }

  if (error instanceof Error) {
    return {
      code: "internal_error",
      message: error.message,
    };
  }

  return {
    code: "internal_error",
    message: "Unknown bridge error.",
  };
}
