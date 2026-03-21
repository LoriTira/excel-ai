import { getSettings } from "./settings";

const DEFAULT_TEMPERATURE = 0.3;
const DEFAULT_MAX_TOKENS = 512;
const REQUEST_TIMEOUT_MS = 30000;

const SYSTEM_PROMPT = "You are a helpful assistant embedded in Excel. Give concise answers suitable for spreadsheet cells.";

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
    // Local mode: use the dev-server proxy when running locally, otherwise call LM Studio directly
    const isDevServer = window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";
    baseUrl = isDevServer ? "/lmstudio" : settings.localAddress;
  }

  const resolvedModel = model || (settings.provider === "api" ? settings.apiModel : undefined);

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
    throw new Error(`LM Studio not running at ${settings.localAddress}`);
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
