/* global Office */

import { getSettings, saveSettings, ProviderSettings } from "../services/settings";
import * as cache from "../services/cache";

Office.onReady(() => {
  loadSettingsIntoForm();

  // Provider radio toggle
  document.querySelectorAll<HTMLInputElement>('input[name="provider"]').forEach((radio) => {
    radio.addEventListener("change", () => togglePanels(radio.value));
  });

  document.getElementById("save-btn")!.addEventListener("click", handleSave);
});

function togglePanels(provider: string): void {
  document.getElementById("local-panel")!.classList.toggle("active", provider === "local");
  document.getElementById("api-panel")!.classList.toggle("active", provider === "api");
}

function loadSettingsIntoForm(): void {
  const s = getSettings();

  // Set provider radio
  const radio = document.querySelector<HTMLInputElement>(`input[name="provider"][value="${s.provider}"]`);
  if (radio) radio.checked = true;
  togglePanels(s.provider);

  // Populate fields
  (document.getElementById("local-address") as HTMLInputElement).value = s.localAddress;
  (document.getElementById("api-endpoint") as HTMLInputElement).value = s.apiEndpoint;
  (document.getElementById("api-key") as HTMLInputElement).value = s.apiKey;
  (document.getElementById("api-model") as HTMLInputElement).value = s.apiModel;
}

function handleSave(): void {
  const provider = document.querySelector<HTMLInputElement>('input[name="provider"]:checked')!.value as "local" | "api";
  const localAddress = (document.getElementById("local-address") as HTMLInputElement).value.trim();
  const apiEndpoint = (document.getElementById("api-endpoint") as HTMLInputElement).value.trim();
  const apiKey = (document.getElementById("api-key") as HTMLInputElement).value.trim();
  const apiModel = (document.getElementById("api-model") as HTMLInputElement).value.trim();

  // Validate
  if (provider === "local" && !localAddress) {
    showStatus("Server address is required.", true);
    return;
  }
  if (provider === "api") {
    if (!apiEndpoint) { showStatus("API endpoint is required.", true); return; }
    if (!apiKey) { showStatus("API key is required.", true); return; }
    if (!apiModel) { showStatus("Model name is required.", true); return; }
  }

  const settings: ProviderSettings = { provider, localAddress, apiEndpoint, apiKey, apiModel };
  saveSettings(settings);
  cache.clear();
  showStatus("Settings saved.");
}

function showStatus(msg: string, isError = false): void {
  const el = document.getElementById("status-msg")!;
  el.textContent = msg;
  el.className = isError ? "status-err" : "status-ok";
  setTimeout(() => { el.textContent = ""; el.className = ""; }, 3000);
}
