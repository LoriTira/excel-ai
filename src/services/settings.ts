const STORAGE_KEY = "excelai_settings";

export interface ProviderSettings {
  provider: "local" | "api";
  localAddress: string;
  apiEndpoint: string;
  apiKey: string;
  apiModel: string;
}

export function getDefaults(): ProviderSettings {
  return {
    provider: "local",
    localAddress: "",
    apiEndpoint: "",
    apiKey: "",
    apiModel: "",
  };
}

export function getSettings(): ProviderSettings {
  const defaults = getDefaults();
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaults;
    const parsed = JSON.parse(raw);
    return { ...defaults, ...parsed };
  } catch {
    return defaults;
  }
}

export function saveSettings(settings: ProviderSettings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}
