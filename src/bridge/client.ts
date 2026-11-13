import { createInterface } from "node:readline";
import { ChildProcessWithoutNullStreams, spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { randomUUID } from "node:crypto";

import {
  BridgeClientError,
  BridgeErrorPayload,
  BridgeFailureResponse,
  BridgeRequest,
  BridgeResponse,
  BridgeSuccessResponse,
} from "./types.js";

const DEFAULT_CALL_TIMEOUT_MS = 120_000;
/** Agent tools wait on LLM calls and run whole plans, so they get a longer budget. */
const AGENT_CALL_TIMEOUT_MS = 600_000;
const SHUTDOWN_TIMEOUT_MS = 2_000;
/** Delay between asking a stuck bridge to stop (SIGTERM) and forcing it (SIGKILL). */
const KILL_GRACE_MS = 2_000;
const STDERR_TAIL_LINES = 10;

/** Every parameter name that carries a presentation session id. */
const SESSION_ID_PARAMS = [
  "presentation_id",
  "source_presentation_id",
  "target_presentation_id",
  "presentation_id_a",
  "presentation_id_b",
] as const;

interface PendingRequest {
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
  method: string;
  timer: NodeJS.Timeout;
}

export interface BridgeClientOptions {
  /** Interpreter to launch. Defaults to PPTX_PYTHON_BIN, then "python3". */
  pythonBin?: string;
  /** Bridge script to run. Defaults to PPTX_PY_BRIDGE, then the bundled python/bridge.py. */
  bridgeScriptPath?: string;
  /** Time allowed for one bridge call. Defaults to PPTX_BRIDGE_TIMEOUT_MS, then 120 s. */
  callTimeoutMs?: number;
}

function positiveInt(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}

function log(message: string): void {
  process.stderr.write(`[bridge-client] ${message}\n`);
}

export class PythonBridgeClient {
  private child?: ChildProcessWithoutNullStreams;
  private starting?: Promise<void>;
  private closing = false;
  private stderrTail: string[] = [];
  private readonly pending = new Map<string, PendingRequest>();
  /** Session ids known to exist in the running bridge process. */
  private readonly liveSessions = new Set<string>();
  /** Session ids that died with a previous bridge process, with the reason. */
  private readonly lostSessions = new Map<string, string>();
  private readonly pythonBin: string;
  private readonly bridgeScriptPath: string;
  private readonly callTimeoutMs: number;

  public constructor(options: BridgeClientOptions = {}) {
    this.pythonBin = options.pythonBin ?? process.env.PPTX_PYTHON_BIN ?? "python3";
    // Resolve the bundled bridge relative to this module (src/bridge or dist/bridge), never the caller's cwd:
    // MCP clients launch the server from arbitrary working directories.
    const bundledBridgePath = fileURLToPath(new URL("../../python/bridge.py", import.meta.url));
    const override = options.bridgeScriptPath ?? process.env.PPTX_PY_BRIDGE;
    this.bridgeScriptPath = override ? path.resolve(override) : bundledBridgePath;
    this.callTimeoutMs =
      options.callTimeoutMs ?? positiveInt(process.env.PPTX_BRIDGE_TIMEOUT_MS) ?? DEFAULT_CALL_TIMEOUT_MS;
  }

  /** Launch the bridge process if it is not running. Safe to call concurrently. */
  public start(): Promise<void> {
    if (!this.starting) {
      const launch = this.launch();
      this.starting = launch;
      launch.catch(() => {
        // Let the next call retry instead of replaying a stale failure.
        if (this.starting === launch) {
          this.starting = undefined;
        }
      });
    }
    return this.starting;
  }

  public async call<TResult = unknown>(method: string, params: Record<string, unknown>): Promise<TResult> {
    const sessionIds = SESSION_ID_PARAMS.map((name) => params[name]).filter(
      (value): value is string => typeof value === "string",
    );

    for (const sessionId of sessionIds) {
      const reason = this.lostSessions.get(sessionId);
      if (reason !== undefined) {
        throw new BridgeClientError({
          code: "session_lost",
          message:
            `Presentation session ${sessionId} no longer exists: ${reason}. Unsaved changes are gone. ` +
            "Open the file again with pptx_open_presentation (or recreate it) and use the new presentation_id.",
          details: { presentation_id: sessionId, reason },
        });
      }
    }

    await this.start();
    const result = await this.request<TResult>(method, params, this.timeoutFor(method));
    this.trackSessions(method, sessionIds, result);
    return result;
  }

  public async close(): Promise<void> {
    const child = this.child;
    if (!child) {
      return;
    }

    this.closing = true;
    try {
      await this.request("__shutdown__", {}, SHUTDOWN_TIMEOUT_MS);
    } catch {
      // Ignore shutdown errors and terminate the process below.
    }

    if (this.child === child) {
      this.child = undefined;
      this.starting = undefined;
    }
    this.liveSessions.clear();
    child.kill();
    this.closing = false;
  }

  private timeoutFor(method: string): number {
    return method.startsWith("pptx_agent_") ? Math.max(this.callTimeoutMs, AGENT_CALL_TIMEOUT_MS) : this.callTimeoutMs;
  }

  private async launch(): Promise<void> {
    this.stderrTail = [];
    const child = spawn(this.pythonBin, [this.bridgeScriptPath], {
      stdio: ["pipe", "pipe", "pipe"],
      env: {
        ...process.env,
        PYTHONUNBUFFERED: "1",
      },
    });
    this.child = child;

    createInterface({ input: child.stdout }).on("line", (line) => this.handleResponseLine(line));

    createInterface({ input: child.stderr }).on("line", (line) => {
      this.stderrTail.push(line);
      if (this.stderrTail.length > STDERR_TAIL_LINES) {
        this.stderrTail.shift();
      }
      process.stderr.write(`[python-bridge] ${line}\n`);
    });

    // Writes to a dead child fail with EPIPE; the exit/error handlers below report the real cause.
    child.stdin.on("error", () => undefined);

    child.on("error", (error: NodeJS.ErrnoException) => {
      const notFound = error.code === "ENOENT";
      this.handleChildGone(
        child,
        new BridgeClientError({
          code: notFound ? "dependency_missing" : "bridge_unavailable",
          message: notFound
            ? `Python interpreter "${this.pythonBin}" was not found. Install Python 3.11+ and the packages in ` +
              "requirements.txt, or set PPTX_PYTHON_BIN to the interpreter to use."
            : `Could not start the Python bridge: ${error.message}`,
          details: { python_bin: this.pythonBin, bridge_script: this.bridgeScriptPath },
        }),
        "the Python bridge could not be started",
      );
    });

    child.on("exit", (code, signal) => {
      const details = { exit_code: code, signal, python_bin: this.pythonBin, stderr_tail: [...this.stderrTail] };
      const missingModule = this.stderrTail
        .map((line) => /ModuleNotFoundError: No module named '([^']+)'/.exec(line)?.[1])
        .find((name) => name !== undefined);

      this.handleChildGone(
        child,
        missingModule
          ? new BridgeClientError({
              code: "dependency_missing",
              message:
                `Python package "${missingModule}" is not installed for interpreter "${this.pythonBin}". ` +
                "Install the packages in requirements.txt for that interpreter, or set PPTX_PYTHON_BIN to one that has them.",
              details,
            })
          : new BridgeClientError({
              code: "bridge_unavailable",
              message: `Python bridge exited unexpectedly (code=${code ?? "null"}, signal=${signal ?? "null"}).`,
              details,
            }),
        "the Python bridge process exited unexpectedly",
      );
    });

    // Light startup probe so bridge errors (missing interpreter, missing packages) fail fast.
    await this.request("pptx_get_engine_info", {}, this.callTimeoutMs);
  }

  private request<TResult>(method: string, params: Record<string, unknown>, timeoutMs: number): Promise<TResult> {
    const child = this.child;
    if (!child || !child.stdin.writable) {
      return Promise.reject(
        new BridgeClientError({ code: "bridge_unavailable", message: "Python bridge process is not available." }),
      );
    }

    const id = randomUUID();
    const payload: BridgeRequest = { id, method, params };

    return new Promise<TResult>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(
          new BridgeClientError({
            code: "timeout",
            message:
              `${method} did not finish within ${timeoutMs} ms. The Python bridge was stopped and will restart ` +
              "on the next call; presentations that were open must be opened again.",
            details: { method, timeout_ms: timeoutMs },
          }),
        );
        // The bridge handles one request at a time, so a stuck call blocks every later one.
        this.terminate(child, `a call to ${method} timed out and the Python bridge was restarted`);
      }, timeoutMs);

      this.pending.set(id, { resolve: resolve as (value: unknown) => void, reject, method, timer });
      child.stdin.write(`${JSON.stringify(payload)}\n`);
    });
  }

  private terminate(child: ChildProcessWithoutNullStreams, reason: string): void {
    this.handleChildGone(
      child,
      new BridgeClientError({ code: "bridge_unavailable", message: `Python bridge stopped: ${reason}.` }),
      reason,
    );

    child.kill("SIGTERM");
    setTimeout(() => {
      if (child.exitCode === null && child.signalCode === null) {
        child.kill("SIGKILL");
      }
    }, KILL_GRACE_MS).unref();
  }

  /** Detach from a bridge process that died or is being killed. Runs once per process. */
  private handleChildGone(child: ChildProcessWithoutNullStreams, error: BridgeClientError, reason: string): void {
    if (this.child !== child) {
      return;
    }
    this.child = undefined;
    this.starting = undefined;

    for (const [id, request] of this.pending.entries()) {
      clearTimeout(request.timer);
      request.reject(error);
      this.pending.delete(id);
    }

    if (this.closing) {
      return;
    }

    for (const sessionId of this.liveSessions) {
      this.lostSessions.set(sessionId, reason);
    }
    if (this.liveSessions.size > 0) {
      log(`${this.liveSessions.size} presentation session(s) lost: ${reason}.`);
    }
    this.liveSessions.clear();
    log(error.message);
  }

  /** Remember which sessions live in the current bridge process, to report them if it dies. */
  private trackSessions(method: string, usedIds: string[], result: unknown): void {
    if (method === "pptx_close_presentation") {
      for (const sessionId of usedIds) {
        this.liveSessions.delete(sessionId);
      }
      return;
    }

    for (const sessionId of usedIds) {
      this.liveSessions.add(sessionId);
    }

    const created = (result as { presentation_id?: unknown } | null)?.presentation_id;
    if (typeof created === "string") {
      this.liveSessions.add(created);
    }
  }

  private handleResponseLine(line: string): void {
    if (!line.trim()) {
      return;
    }

    let parsed: BridgeResponse;
    try {
      parsed = JSON.parse(line) as BridgeResponse;
    } catch {
      log(`Ignored a non-JSON line on the bridge protocol stream: ${line.slice(0, 200)}`);
      return;
    }

    const pending = this.pending.get(parsed.id);
    if (!pending) {
      log(`Dropped a response with no matching request (id=${parsed.id}).`);
      return;
    }

    this.pending.delete(parsed.id);
    clearTimeout(pending.timer);

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
      ...(error.hint ? { hint: error.hint } : {}),
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
