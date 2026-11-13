export type BridgeErrorCode =
  | "validation_error"
  | "not_found"
  | "engine_error"
  | "dependency_missing"
  | "conflict"
  | "internal_error"
  // A call exceeded its time budget (bridge call in the client, or slide rendering in the bridge).
  | "timeout"
  // Raised by the TypeScript client only:
  | "session_lost"
  | "bridge_unavailable";

export interface BridgeErrorPayload {
  code: BridgeErrorCode;
  message: string;
  details?: Record<string, unknown>;
  /** One sentence telling the caller what to do next. */
  hint?: string;
}

export interface BridgeRequest<TParams = Record<string, unknown>> {
  id: string;
  method: string;
  params: TParams;
}

export interface BridgeSuccessResponse<TResult = unknown> {
  id: string;
  result: TResult;
}

export interface BridgeFailureResponse {
  id: string;
  error: BridgeErrorPayload;
}

export type BridgeResponse<TResult = unknown> =
  | BridgeSuccessResponse<TResult>
  | BridgeFailureResponse;

export interface SlideSummary {
  index: number;
  title: string;
  layout: string;
  shape_count: number;
}

export interface PresentationState {
  presentation_id: string;
  slide_count: number;
  slides: SlideSummary[];
}

export interface MutatingResponseEnvelope {
  success: true;
  warning?: string;
  /** Always present; replaces presentation_state unless the call passed include_state: true. */
  slide_count: number;
  presentation_state?: PresentationState;
  [key: string]: unknown;
}

export class BridgeClientError extends Error {
  public readonly code: BridgeErrorCode;
  public readonly details?: Record<string, unknown>;
  public readonly hint?: string;

  public constructor(payload: BridgeErrorPayload) {
    super(payload.message);
    this.name = "BridgeClientError";
    this.code = payload.code;
    this.details = payload.details;
    this.hint = payload.hint;
  }
}
