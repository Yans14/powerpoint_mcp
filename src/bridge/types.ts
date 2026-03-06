export type BridgeErrorCode =
  | "validation_error"
  | "not_found"
  | "engine_error"
  | "dependency_missing"
  | "conflict"
  | "internal_error";

export interface BridgeErrorPayload {
  code: BridgeErrorCode;
  message: string;
  details?: Record<string, unknown>;
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
  presentation_state: PresentationState;
  [key: string]: unknown;
}

export class BridgeClientError extends Error {
  public readonly code: BridgeErrorCode;
  public readonly details?: Record<string, unknown>;

  public constructor(payload: BridgeErrorPayload) {
    super(payload.message);
    this.name = "BridgeClientError";
    this.code = payload.code;
    this.details = payload.details;
  }
}
