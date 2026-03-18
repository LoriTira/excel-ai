const BASE_URL = "http://localhost:1234";
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
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

  const messages: ChatMessage[] = [
    { role: "system", content: SYSTEM_PROMPT },
    { role: "user", content: prompt },
  ];

  let response: Response;
  try {
    response = await fetch(`${BASE_URL}/v1/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        ...(model ? { model } : {}),
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
    throw new Error(`LM Studio not running at ${BASE_URL}`);
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
