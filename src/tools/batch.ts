import { TOOL_MAP } from "./catalog.js";

export const BATCH_TOOL_NAME = "pptx_batch";

/** Saving and closing cannot be undone by a rollback, so they stay outside batches. */
const EXCLUDED_FROM_BATCH = new Set([BATCH_TOOL_NAME, "pptx_save_presentation", "pptx_close_presentation"]);

const NOT_BATCHABLE_HINT =
  "Batches accept tools that take presentation_id, except pptx_save_presentation, pptx_close_presentation, " +
  "pptx_batch and the pptx_agent_* tools. Call those separately.";

interface BatchInput {
  presentation_id: string;
  operations: Array<{ tool: string; arguments?: Record<string, unknown> }>;
  atomic?: boolean;
}

interface OperationProblem {
  operation_index: number;
  tool: string;
  problem: unknown;
}

export type PreparedBatch =
  | { ok: true; params: Record<string, unknown> }
  | { ok: false; problems: OperationProblem[]; hint?: string };

/**
 * Validate every operation of a batch against its own tool schema before anything runs, so a
 * typo in the last operation is reported without executing (and rolling back) the first ones.
 * Returns the bridge parameters, with each operation's arguments parsed (defaults applied).
 */
export function prepareBatch(input: BatchInput, isToolAvailable: (name: string) => boolean): PreparedBatch {
  const problems: OperationProblem[] = [];
  let hint: string | undefined;

  const operations = input.operations.map((operation, index) => {
    const { tool } = operation;
    const definition = TOOL_MAP.get(tool);
    const { include_state: _ignored, ...rawArguments } = operation.arguments ?? {};

    if (!definition || !isToolAvailable(tool)) {
      problems.push({ operation_index: index, tool, problem: `Unknown tool: ${tool}` });
      return operation;
    }

    if (EXCLUDED_FROM_BATCH.has(tool) || tool.startsWith("pptx_agent_")) {
      problems.push({ operation_index: index, tool, problem: `${tool} cannot be used inside ${BATCH_TOOL_NAME}.` });
      hint = NOT_BATCHABLE_HINT;
      return operation;
    }

    const ownId = rawArguments.presentation_id;
    if (ownId !== undefined && ownId !== input.presentation_id) {
      problems.push({
        operation_index: index,
        tool,
        problem: "Every operation runs on the batch's presentation_id; remove presentation_id from its arguments.",
      });
      return operation;
    }

    const parsed = definition.schema.safeParse({ ...rawArguments, presentation_id: input.presentation_id });
    if (!parsed.success) {
      const unknownKeys = parsed.error.issues.some(
        (issue) => issue.code === "unrecognized_keys" && issue.keys.includes("presentation_id"),
      );
      if (unknownKeys) {
        problems.push({ operation_index: index, tool, problem: `${tool} cannot be used inside ${BATCH_TOOL_NAME}.` });
        hint = NOT_BATCHABLE_HINT;
      } else {
        problems.push({ operation_index: index, tool, problem: parsed.error.flatten() });
      }
      return operation;
    }

    return { tool, arguments: parsed.data as Record<string, unknown> };
  });

  if (problems.length > 0) {
    return { ok: false, problems, hint };
  }

  return {
    ok: true,
    params: { presentation_id: input.presentation_id, operations, ...(input.atomic === undefined ? {} : { atomic: input.atomic }) },
  };
}
