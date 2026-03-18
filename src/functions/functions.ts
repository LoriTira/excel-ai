/* global CustomFunctions */

import * as cache from "../services/cache";
import * as queue from "../services/queue";
import { complete } from "../services/lmstudio";

const MAX_CELL_LENGTH = 32767;

function serializeContext(context: unknown[][] | undefined): string {
  if (!context) return "";

  const rows = context.map((row) =>
    row
      .map((cell) => {
        if (cell === null || cell === undefined || cell === "") return "";
        return String(cell);
      })
      .join("\t")
  );

  return rows.join("\n").trim();
}

/**
 * Send a prompt to a local AI model and return the response.
 * @customfunction
 * @param {string} prompt The instruction or question to send to the AI model.
 * @param [context] A cell or range whose values are included as context.
 * @param [model] Override the default model loaded in LM Studio.
 * @returns The model's response.
 */
export async function ai(prompt: string, context?: unknown[][], model?: string): Promise<string> {
  const serialized = serializeContext(context);
  const cacheKey = cache.buildKey(prompt, serialized);

  const cached = cache.get(cacheKey);
  if (cached !== undefined) {
    return cached;
  }

  const fullPrompt = serialized ? `${prompt}\n\nContext:\n${serialized}` : prompt;

  try {
    const result = await queue.enqueue(() => complete(fullPrompt, model));
    const truncated = result.length > MAX_CELL_LENGTH ? result.slice(0, MAX_CELL_LENGTH) : result;
    cache.set(cacheKey, truncated);
    return truncated;
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "Unknown error";
    return `#ERROR: ${message}`;
  }
}

CustomFunctions.associate("AI", ai);
