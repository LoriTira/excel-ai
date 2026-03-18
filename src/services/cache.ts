const store = new Map<string, string>();

export function buildKey(prompt: string, context: string): string {
  return `${prompt}||${context}`;
}

export function get(key: string): string | undefined {
  return store.get(key);
}

export function set(key: string, value: string): void {
  store.set(key, value);
}

export function clear(): void {
  store.clear();
}
