/* global Office */

import { getSettings, saveSettings, ProviderSettings } from "../services/settings";
import * as cache from "../services/cache";
import { freezeWorkbook } from "../services/freeze";

Office.onReady(() => {
  // Ribbon button handler — must be registered at load, not gated on taskpane DOM.
  if (Office.actions && typeof Office.actions.associate === "function") {
    Office.actions.associate("freezeAIResults", freezeFromRibbon);
  }

  // Taskpane-only wiring below. Bail if the DOM hasn't been injected (ribbon-only load).
  if (!document.getElementById("save-btn")) return;

  loadSettingsIntoForm();

  // Heartbeat keeps Ollama alive while Excel is open; proxy starts/stops it on demand
  const sendHeartbeat = () => {
    fetch("/heartbeat", { signal: AbortSignal.timeout(3000) }).catch(() => {});
  };
  setInterval(sendHeartbeat, 30000);

  // First heartbeat triggers Ollama startup; detect models once it confirms running
  fetch("/heartbeat", { signal: AbortSignal.timeout(15000) })
    .then(() => detectLocalModel())
    .catch(() => detectLocalModel()); // try anyway on timeout

  // Provider radio toggle
  document.querySelectorAll<HTMLInputElement>('input[name="provider"]').forEach((radio) => {
    radio.addEventListener("change", () => togglePanels(radio.value));
  });

  document.getElementById("save-btn")!.addEventListener("click", handleSave);
  document.getElementById("freeze-btn")!.addEventListener("click", handleFreezeClick);
});

async function handleFreezeClick(): Promise<void> {
  const ok = window.confirm(
    "Freeze replaces every =EXCELAI.AI(...) formula in this workbook with its current value. " +
      "This cannot be undone with Ctrl+Z after saving.\n\nContinue?"
  );
  if (!ok) return;

  showStatus("Freezing AI results\u2026");
  try {
    const { frozen, skipped } = await freezeWorkbook();
    const parts = [`Froze ${frozen} cell${frozen === 1 ? "" : "s"}`];
    if (skipped > 0) parts.push(`skipped ${skipped} error cell${skipped === 1 ? "" : "s"}`);
    showStatus(parts.join(", ") + ".");
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : "Unknown error";
    showStatus(`Freeze failed: ${msg}`, true);
  }
}

async function freezeFromRibbon(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await freezeWorkbook();
  } catch {
    // Ribbon handler has no UI surface; errors are visible in cells already.
  } finally {
    event.completed();
  }
}

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

async function detectLocalModel(): Promise<void> {
  const s = getSettings();
  const baseUrl = s.localAddress || "";
  const el = document.getElementById("local-model-info")!;
  try {
    const resp = await fetch(`${baseUrl}/v1/models`, { signal: AbortSignal.timeout(3000) });
    if (!resp.ok) throw new Error();
    const data = await resp.json();
    const models: string[] = (data?.data || []).map((m: { id: string }) => m.id);
    if (models.length > 0) {
      el.innerHTML = `Using <strong>${models[0]}</strong>` +
        (models.length > 1 ? ` <span style="color:#999">(+${models.length - 1} more available)</span>` : "");
    } else {
      el.innerHTML = 'No models found. Run <code>ollama pull qwen2.5:1.5b</code>';
    }
  } catch {
    el.innerHTML = 'Cannot reach Ollama. Check troubleshooting below.';
  }
}

async function handleSave(): Promise<void> {
  const provider = document.querySelector<HTMLInputElement>('input[name="provider"]:checked')!.value as "local" | "api";
  const localAddress = (document.getElementById("local-address") as HTMLInputElement).value.trim();
  const apiEndpoint = (document.getElementById("api-endpoint") as HTMLInputElement).value.trim();
  const apiKey = (document.getElementById("api-key") as HTMLInputElement).value.trim();
  const apiModel = (document.getElementById("api-model") as HTMLInputElement).value.trim();

  // Validate
  if (provider === "api") {
    if (!apiEndpoint) { showStatus("API endpoint is required.", true); return; }
    if (!apiKey) { showStatus("API key is required.", true); return; }
    if (!apiModel) { showStatus("Model name is required.", true); return; }
  }

  const settings: ProviderSettings = { provider, localAddress, apiEndpoint, apiKey, apiModel };
  saveSettings(settings);
  cache.clear();

  // Test connection to Ollama (works in both dev and production via proxy)
  if (provider === "local") {
    showStatus("Saved. Testing connection\u2026");
    const testUrl = localAddress ? `${localAddress}/v1/models` : "/v1/models";
    const ok = await testLocalConnection(testUrl);
    if (!ok) {
      showStatus("Saved, but cannot reach Ollama. Check troubleshooting tips below.", true);
      return;
    }
    detectLocalModel();
  }

  showStatus("Settings saved.");
}

async function testLocalConnection(url: string): Promise<boolean> {
  try {
    const resp = await fetch(url, { signal: AbortSignal.timeout(5000) });
    return resp.ok;
  } catch {
    return false;
  }
}

function showStatus(msg: string, isError = false): void {
  const el = document.getElementById("status-msg")!;
  el.textContent = msg;
  el.className = isError ? "status-err" : "status-ok";
  setTimeout(() => { el.textContent = ""; el.className = ""; }, 3000);
}
