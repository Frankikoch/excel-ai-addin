/**
 * Excel AI Add-in - Main Application (v2 - Direct AI Mode)
 */

// Configuration - Direct AI API (OpenRouter or OpenCode)
const CONFIG = {
  aiProvider: localStorage.getItem("aiProvider") || "openrouter", // "openrouter" or "opencode"
  apiKey: localStorage.getItem("apiKey") || "",
  model: localStorage.getItem("model") || "openai/gpt-4o-mini"
};

// State
const state = {
  privacyMode: localStorage.getItem("privacyMode") !== "false",
  model: CONFIG.model,
  messages: [],
  isLoading: false,
  excelConnected: false
};

// DOM Elements
const elements = {
  userInput: document.getElementById("user-input"),
  sendBtn: document.getElementById("send-btn"),
  messages: document.getElementById("messages"),
  status: document.getElementById("status"),
  privacyToggle: document.getElementById("privacy-toggle"),
  settingsBtn: document.getElementById("settings-btn"),
  modelDisplay: document.getElementById("model-display")
};

// Initialize
document.addEventListener("DOMContentLoaded", init);

async function init() {
  console.log("🤖 Excel AI v2 initializing...");

  if (elements.privacyToggle) {
    elements.privacyToggle.checked = state.privacyMode;
  }
  if (elements.modelDisplay) {
    elements.modelDisplay.textContent = `Modelo: ${CONFIG.model}`;
  }

  elements.sendBtn?.addEventListener("click", sendMessage);
  elements.userInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });
  elements.privacyToggle?.addEventListener("change", togglePrivacy);
  elements.settingsBtn?.addEventListener("click", openSettings);

  document.querySelectorAll(".quick-btn").forEach(btn => {
    btn.addEventListener("click", () => sendMessage(btn.dataset.prompt));
  });

  document.querySelectorAll(".toolbar-btn[data-cmd]").forEach(btn => {
    btn.addEventListener("click", () => {
      const cmdMap = {
        analyze: "Analiza los datos seleccionados en esta hoja",
        explain: "Explica la fórmula en la celda activa",
        fix: "Busca y corrige errores en esta hoja",
        summarize: "Resume el contenido seleccionado"
      };
      if (cmdMap[btn.dataset.cmd]) sendMessage(cmdMap[btn.dataset.cmd]);
    });
  });

  await connectToExcel();
  await checkAIConnection();

  updateStatus(CONFIG.apiKey ? "🟢 IA lista" : "⚠️ Configura API Key");
}

/// Check AI API connection
async function checkAIConnection() {
  if (!CONFIG.apiKey && !CONFIG.aiEndpoint) {
    updateStatus("⚠️ Sin configurar - Modo demo");
    return false;
  }
  return true;
}

/// Connect to Excel via Office.js
async function connectToExcel() {
  try {
    if (typeof Excel !== "undefined") {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await context.sync();
        state.excelConnected = true;
        console.log("📊 Connected to Excel:", sheet.name);
        updateStatus("📊 Excel conectado");
      });
    }
  } catch (error) {
    console.warn("Excel context:", error.message);
  }
}

/// Send message to AI
async function sendMessage(prompt) {
  const text = prompt || elements.userInput?.value?.trim();
  if (!text || state.isLoading) return;

  addMessage("user", text);
  if (!prompt && elements.userInput) elements.userInput.value = "";

  state.isLoading = true;
  updateStatus("⏳ Procesando...");

  try {
    const context = await getCellContext();
    const response = await callAI(text, context);
    addMessage("assistant", response);
  } catch (error) {
    addMessage("assistant", "❌ Error: " + error.message);
  } finally {
    state.isLoading = false;
    updateStatus(CONFIG.apiKey ? "🟢 IA lista" : "⚠️ Modo demo");
  }
}

/// Get current cell context from Excel
async function getCellContext() {
  try {
    if (typeof Excel === "undefined") return [];

    return await Excel.run(async (excelContext) => {
      const range = excelContext.workbook.getSelectedRange();
      range.load(["address", "values", "formulas", "numberFormat"]);
      await excelContext.sync();

      const cells = [];
      const values = range.values;
      const address = range.address;

      for (let i = 0; i < Math.min(values.length, 50); i++) {
        for (let j = 0; j < Math.min(values[i].length, 10); j++) {
          if (values[i][j] !== null && values[i][j] !== "") {
            cells.push({
              address: address.split("!")[1] || address,
              value: values[i][j],
              row: i + 1,
              col: j + 1
            });
          }
        }
      }
      return cells;
    });
  } catch {
    return [];
  }
}

/// Apply privacy filter
function applyPrivacyFilter(cells) {
  return cells.map(cell => ({
    ...cell,
    value: typeof cell.value === "string"
      ? cell.value.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[EMAIL]")
          .replace(/\b\d{3}[-.]?\d{2}[-.]?\d{4}\b/g, "[SSN]")
      : cell.value
  }));
}

/// Call AI API (supports OpenRouter and OpenCode)
async function callAI(text, cellContext) {
  // Demo mode if no API key
  if (!CONFIG.apiKey) {
    return generateMockResponse(text, cellContext?.length || 0);
  }

  const hasContext = cellContext && cellContext.length > 0;
  const contextInfo = hasContext 
    ? `\n\n📊 **Context (${cellContext.length} cells):**\n` + cellContext.slice(0, 15).map(c => 
        `${c.address}: ${c.value}`).join("\n")
    : "\n\n⚠️ No cells selected. Select a range first.";


  const systemPrompt = "You are an Excel data analysis assistant. Help analyze spreadsheets, explain formulas, fix errors. Respond in user's language.";

  const body = {
    model: CONFIG.model,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: text + contextInfo }
    ],
    max_tokens: 1000
  };

  let url, headers;


  if (CONFIG.aiProvider === "openrouter") {
    url = "https://openrouter.ai/api/v1/chat/completions";
    headers = {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${CONFIG.apiKey}`,
      "HTTP-Referer": "https://frankikoch.github.io",
      "X-Title": "Excel AI"
    };
  } else {
    url = "https://opencode.ai/v1/chat/completions";
    headers = {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${CONFIG.apiKey}`
    };
  }

  try {
    const response = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
      signal: AbortSignal.timeout(45000)
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error?.message || err.message || `HTTP ${response.status}`);
    }

    const data = await response.json();
    return data.choices?.[0]?.message?.content 
      || data.choices?.[0]?.text?.content
      || data.output?.[0]?.content?.text
      || "No response received";
  } catch (error) {
    console.error("AI Error:", error);
    return generateMockResponse(text, cellContext?.length || 0);
  }
}

/// Generate mock response
function generateMockResponse(prompt, cellCount) {
  const lower = prompt.toLowerCase();

  if (lower.includes("analizar") || lower.includes("análisis")) {
    return `📊 **Análisis de datos**\n\n` +
      `Celdas detectadas: ${cellCount}\n\n` +
      `**Para análisis real:**\n` +
      `1. Haz clic en ⚙️ (arriba derecha)\n` +
      `2. Introduce tu API Key de OpenCode\n` +
      `3. Selecciona celdas en Excel\n` +
      `4. Envía tu pregunta\n\n` +
      `**Sin API Key = modo demo** (responde ejemplos genéricos)`;
  }

  if (lower.includes("fórmula") || lower.includes("explicar")) {
    return `📝 **Explicador de fórmulas**\n\n` +
      `Selecciona una celda con fórmula y pregúntame.\n\n` +
      `Ejemplos de fórmulas Excel:\n` +
      `- =SUM(A1:A10) — Suma valores\n` +
      `- =BUSCARV(...) — Busca en tabla\n` +
      `- =SI(C1>10,"Mayor","Menor") — Condicional`;
  }

  if (lower.includes("error") || lower.includes("corregir")) {
    return `🐛 **Corrector de errores**\n\n` +
      `Selecciona celdas con errores y te ayudo.\n\n` +
      `Errores comunes:\n` +
      `- #¡DIV/0! — División por cero\n` +
      `- #¡REF! — Referencia inválida\n` +
      `- #¿NOMBRE? — Función desconocida`;
  }

  if (lower.includes("resumir")) {
    return `📋 **Resumidor**\n\n` +
      `${cellCount} celdas analizadas.\n\n` +
      `Selecciona más datos para un resumen completo.`;
  }

  return `🤖 **Excel AI**\n\n` +
    `Recibido: "${prompt.slice(0, 40)}${prompt.length > 40 ? "..." : ""}"\n` +
    `Celdas: ${cellCount}\n` +
    `Modo: ${state.privacyMode ? "🔒 Privacidad" : "🔓 Normal"}\n\n` +
    `💡 Configura tu API Key en ⚙️ para activar la IA real.`;
}

/// Add message to chat
function addMessage(role, content) {
  if (!elements.messages) return;

  const welcome = elements.messages.querySelector(".welcome-message");
  if (welcome && role === "assistant") welcome.style.display = "none";

  const div = document.createElement("div");
  div.className = `message ${role}`;
  div.textContent = content;
  elements.messages.appendChild(div);
  elements.messages.scrollTop = elements.messages.scrollHeight;
  state.messages.push({ role, content, timestamp: Date.now() });
}

/// Toggle privacy
function togglePrivacy() {
  state.privacyMode = !state.privacyMode;
  localStorage.setItem("privacyMode", state.privacyMode);
  updateStatus(state.privacyMode ? "🔒 Privacidad ON" : "🔓 Normal");
}

/// Update status
function updateStatus(text) {
  if (elements.status) elements.status.textContent = text;
}

/// Settings
function openSettings() {
  initSettingsModal();
  const modal = document.getElementById("settings-modal");
  if (modal) modal.classList.add("show");
  
  document.querySelectorAll(".settings-content").forEach(c => c.classList.add("active"));

  const providerSelect = document.getElementById("ai-provider");
  const apiKeyInput = document.getElementById("api-key");
  const modelSelect = document.getElementById("model-select");
  
  if (providerSelect) providerSelect.value = CONFIG.aiProvider;
  if (apiKeyInput) apiKeyInput.value = CONFIG.apiKey;
  if (modelSelect) modelSelect.value = state.model;
}

function closeSettings() {
  const modal = document.getElementById("settings-modal");
  if (modal) modal.classList.remove("show");
  document.querySelectorAll(".settings-content").forEach(c => c.classList.remove("active"));
}

function saveSettings() {
  const provider = document.getElementById("ai-provider")?.value;
  const apiKey = document.getElementById("api-key")?.value?.trim();
  const model = document.getElementById("model-select")?.value;

  if (provider) {
    localStorage.setItem("aiProvider", provider);
    CONFIG.aiProvider = provider;
  }
  if (apiKey) {
    localStorage.setItem("apiKey", apiKey);
    CONFIG.apiKey = apiKey;
  }
  if (model) {
    localStorage.setItem("model", model);
    state.model = model;
    if (elements.modelDisplay) elements.modelDisplay.textContent = `Modelo: ${model}`;
  }

  closeSettings();
  checkAIConnection();
}

async function testConnection() {
  const statusEl = document.getElementById("settings-status");
  if (statusEl) statusEl.textContent = "⏳ Probando...";

  if (!CONFIG.apiKey) {
    if (statusEl) statusEl.textContent = "⚠️ Sin API Key";
    return;
  }


  const url = CONFIG.aiProvider === "openrouter" 
    ? "https://openrouter.ai/api/v1/models" 
    : "https://opencode.ai/v1/models";

  try {
    const response = await fetch(url, {
      headers: { "Authorization": `Bearer ${CONFIG.apiKey}` },
      signal: AbortSignal.timeout(10000)
    });

    if (response.ok) {
      if (statusEl) statusEl.textContent = "✅ ¡Conectado!";
    } else {
      if (statusEl) statusEl.textContent = `⚠️ HTTP ${response.status}`;
    }
  } catch (error) {
    if (statusEl) statusEl.textContent = `❌ ${error.message.split("\n")[0]}`;
  }
}

/// Settings modal
function initSettingsModal() {
  if (document.getElementById("settings-modal")) return;

  const modal = document.createElement("div");
  modal.id = "settings-modal";
  modal.className = "settings-modal";
  modal.innerHTML = `
    <div class="settings-content">
      <div class="settings-header">
        <h2>⚙️ Configuración</h2>
        <button class="settings-close" id="settings-close-btn">✕</button>
      </div>
      <div class="settings-body">
        <div class="settings-group">
          <label for="ai-provider">Proveedor</label>
          <select id="ai-provider">
            <option value="openrouter">OpenRouter (Recomendado)</option>
            <option value="opencode">OpenCode</option>
          </select>
        </div>
        <div class="settings-group">
          <label for="api-key">API Key</label>
          <input type="password" id="api-key" placeholder="sk-..." />
          <small>Obtén tu key en <a href="https://openrouter.ai" target="_blank">openrouter.ai</a></small>
        </div>
        <div class="settings-group">
          <label for="model-select">Modelo</label>
          <select id="model-select">
            <option value="openai/gpt-4o-mini">GPT-4o Mini (Gratis)</option>
            <option value="anthropic/claude-3.5-haiku">Claude 3.5 Haiku</option>
            <option value="google/gemini-flash-1.5">Gemini Flash 1.5</option>
            <option value="meta-llama/llama-3.1-8b-instruct">Llama 3.1</option>
          </select>
        </div>
        <div class="settings-divider"></div>
        <div class="settings-info">
          <p><strong>Estado:</strong> <span id="settings-status">—</span></p>
          <p><strong>Excel:</strong> <span id="excel-status-display">Conectando...</span></p>
          <p><strong>Versión:</strong> v2.1</p>
        </div>
        <div class="settings-footer">
          <button class="btn-secondary" id="test-conn-btn">🔗 Probar</button>
          <button class="btn-primary" id="save-settings-btn">💾 Guardar</button>
        </div>
      </div>
    </div>
  `;
  document.body.appendChild(modal);

  modal.querySelector("#settings-close-btn")?.addEventListener("click", closeSettings);
  modal.querySelector("#save-settings-btn")?.addEventListener("click", saveSettings);
  modal.querySelector("#test-conn-btn")?.addEventListener("click", testConnection);
  modal.addEventListener("click", (e) => { if (e.target === modal) closeSettings(); });

  if (typeof Excel !== "undefined") {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
      const el = document.getElementById("excel-status-display");
      if (el) el.textContent = "✅ " + sheet.name;
    }).catch(() => {
      const el = document.getElementById("excel-status-display");
      if (el) el.textContent = "⚠️ No detectable";
    });
  } else {
    const el = document.getElementById("excel-status-display");
    if (el) el.textContent = "⚠️ Fuera de Excel";
  }
}

window.ExcelAI = { sendMessage, getCellContext, togglePrivacy, openSettings, closeSettings, testConnection, state, CONFIG };