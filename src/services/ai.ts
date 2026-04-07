import { getSettings } from "./settings";

const DEFAULT_TEMPERATURE = 0.3;
const DEFAULT_MAX_TOKENS = 512;
const REQUEST_TIMEOUT_MS = 30000;

const SYSTEM_PROMPT = "You are a helpful assistant embedded in Excel. Give concise answers suitable for spreadsheet cells.";

// Cache the auto-detected local model name
let cachedLocalModel: string | undefined;

interface ChatMessage {
  role: "system" | "user";
  content: string;
}

interface ChatCompletionResponse {
  choices: Array<{
    message: {
      content: string;
    };
  }>;
}

interface ModelsResponse {
  data: Array<{ id: string }>;
}

/** Detect the first model available in Ollama. */
async function detectLocalModel(baseUrl: string): Promise<string | undefined> {
  if (cachedLocalModel) return cachedLocalModel;
  try {
    const resp = await fetch(`${baseUrl}/v1/models`, { signal: AbortSignal.timeout(3000) });
    if (!resp.ok) return undefined;
    const data: ModelsResponse = await resp.json();
    cachedLocalModel = data?.data?.[0]?.id;
    return cachedLocalModel;
  } catch {
    return undefined;
  }
}

export async function complete(prompt: string, model?: string): Promise<string> {
  const settings = getSettings();

  let baseUrl: string;
  const headers: Record<string, string> = { "Content-Type": "application/json" };

  if (settings.provider === "api") {
    baseUrl = settings.apiEndpoint;
    if (settings.apiKey) {
      headers["Authorization"] = `Bearer ${settings.apiKey}`;
    }
  } else {
    // Local mode: relative URL works in both dev (webpack proxy) and production (Go server)
    baseUrl = settings.localAddress || "";
  }

  let resolvedModel: string | undefined;
  if (model) {
    resolvedModel = model;
  } else if (settings.provider === "api") {
    resolvedModel = settings.apiModel;
  } else {
    // Auto-detect whatever model Ollama has loaded
    resolvedModel = await detectLocalModel(baseUrl);
  }

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

  const messages: ChatMessage[] = [
    { role: "system", content: SYSTEM_PROMPT },
    { role: "user", content: prompt },
  ];

  let response: Response;
  try {
    response = await fetch(`${baseUrl}/v1/chat/completions`, {
      method: "POST",
      headers,
      body: JSON.stringify({
        ...(resolvedModel ? { model: resolvedModel } : {}),
        messages,
        temperature: DEFAULT_TEMPERATURE,
        max_tokens: DEFAULT_MAX_TOKENS,
        // Limit context window to reduce RAM usage (~1.5 GB vs 5+ GB at default 128K)
        ...(settings.provider === "local" ? { num_ctx: 2048 } : {}),
      }),
      signal: controller.signal,
    });
  } catch (err: unknown) {
    if (err instanceof DOMException && err.name === "AbortError") {
      throw new Error("Request timed out (30s)");
    }
    if (settings.provider === "api") {
      throw new Error(`Cannot reach API at ${settings.apiEndpoint}`);
    }
    throw new Error("Ollama not running. Make sure the Excel AI server and Ollama are started.");
  } finally {
    clearTimeout(timeout);
  }

  if (!response.ok) {
    const text = await response.text().catch(() => "Unknown error");
    throw new Error(`Model error: ${text}`);
  }

  const data: ChatCompletionResponse = await response.json();
  const content = data?.choices?.[0]?.message?.content;

  if (!content) {
    throw new Error("Empty response from model");
  }

  return content.trim();
}
