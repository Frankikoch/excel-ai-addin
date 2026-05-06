/**
 * Excel AI Add-in - Main Application (v2 - with MCP connection)
 */

// Configuration
const CONFIG = {
  mcpEndpoint: localStorage.getItem("mcpEndpoint") || "https://YOUR-MCP-SERVER.trycloudflare.com",
  provider: localStorage.getItem("provider") || "opencode",
  model: localStorage.getItem("model") || "minimax-m2.5-free"
};

// State
const state = {
  privacyMode: localStorage.getItem("privacyMode") !== "false",
  provider: CONFIG.provider,
  model: CONFIG.model,
  messages: [],
  isLoading: false
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

  // Load saved settings
  if (elements.privacyToggle) {
    elements.privacyToggle.checked = state.privacyMode;
  }
  if (elements.modelDisplay) {
    elements.modelDisplay.textContent = `Modelo: ${CONFIG.model}`;
  }

  // Event listeners
  elements.sendBtn?.addEventListener("click", sendMessage);
  elements.userInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });
  elements.privacyToggle?.addEventListener("change", togglePrivacy);
  elements.settingsBtn?.addEventListener("click", openSettings);

  // Quick action buttons
  document.querySelectorAll(".quick-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const prompt = btn.dataset.prompt;
      sendMessage(prompt);
    });
  });

  // Toolbar buttons
  document.querySelectorAll(".toolbar-btn[data-cmd]").forEach(btn => {
    btn.addEventListener("click", () => {
      const cmd = btn.dataset.cmd;
      const cmdMap = {
        analyze: "Analiza los datos seleccionados en esta hoja",
        explain: "Explica la fórmula en la celda activa",
        fix: "Busca y corrige errores en esta hoja",
        summarize: "Resume el contenido seleccionado"
      };
      if (cmdMap[cmd]) sendMessage(cmdMap[cmd]);
    });
  });

  // Connect to Excel
  await connectToExcel();

  // Check MCP connection
  await checkMCPConnection();

  updateStatus("🟢 Listo");
}

/// Check MCP server connection
async function checkMCPConnection() {
  try {
    const response = await fetch(`${CONFIG.mcpEndpoint}/health`, {
      method: "GET",
      signal: AbortSignal.timeout(5000)
    });

    if (response.ok) {
      const data = await response.json();
      updateStatus(`🟢 Conectado (${data.history} msgs)`);
      console.log("✅ MCP Server connected");
    }
  } catch (error) {
    console.warn("MCP Server not available:", error.message);
    updateStatus("⚠️ MCP offline (modo demo)");
  }
}

/// Connect to Excel via Office.js
async function connectToExcel() {
  try {
    if (typeof Excel !== "undefined") {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await context.sync();
        console.log("📊 Connected to Excel sheet:", sheet.name);
      });
    }
  } catch (error) {
    console.warn("Excel context not available:", error.message);
  }
}

/// Send message to AI via MCP
async function sendMessage(prompt) {
  const text = prompt || elements.userInput?.value?.trim();
  if (!text || state.isLoading) return;

  // Add user message
  addMessage("user", text);
  if (!prompt && elements.userInput) elements.userInput.value = "";

  state.isLoading = true;
  updateStatus("⏳ Procesando...");

  try {
    // Get selected cells context
    const context = await getCellContext();

    // Send to MCP server
    const response = await callMCP(text, context);

    // Add assistant response
    addMessage("assistant", response);
  } catch (error) {
    addMessage("assistant", "❌ Error: " + error.message);
    console.error("MCP Error:", error);
  } finally {
    state.isLoading = false;
    updateStatus("🟢 Listo");
  }
}

/// Get current cell context from Excel
async function getCellContext() {
  try {
    if (typeof Excel === "undefined") return [];

    const context = await Excel.run(async (excelContext) => {
      const range = excelContext.workbook.getSelectedRange();
      range.load(["address", "values", "formulas", "numberFormat"]);
      await excelContext.sync();

      const cells = [];
      const values = range.values;
      const address = range.address;
      const formulas = range.formulas || [];
      const formats = range.numberFormat || [];

      for (let i = 0; i < Math.min(values.length, 20); i++) {
        for (let j = 0; j < Math.min(values[i].length, 5); j++) {
          cells.push({
            address: address.split("!")[1] || address,
            value: values[i][j],
            formula: formulas[i]?.[j] ? `=${formulas[i][j]}` : undefined,
            dataType: typeof values[i][j],
            format: formats[i]?.[j]
          });
        }
      }
      return cells;
    });

    // Apply privacy if enabled
    if (state.privacyMode) {
      return applyPrivacyFilter(context);
    }

    return context;
  } catch {
    return [];
  }
}

/// Apply privacy filter (anonymize sensitive data)
function applyPrivacyFilter(cells) {
  return cells.map(cell => ({
    ...cell,
    value: typeof cell.value === "string"
      ? cell.value.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[EMAIL]")
          .replace(/\b\d{3}[-.]?\d{2}[-.]?\d{4}\b/g, "[SSN]")
          .replace(/\b\d{10,}\b/g, "[ID]")
      : cell.value,
    formula: undefined // Never send formulas to external AI
  }));
}

/// Call MCP Server
async function callMCP(text, cellContext) {
  try {
    const headers = { "Content-Type": "application/json" };
    const apiKey = localStorage.getItem("apiKey");
    if (apiKey) headers["Authorization"] = `Bearer ${apiKey}`;

    const response = await fetch(`${CONFIG.mcpEndpoint}/chat`, {
      method: "POST",
      headers,
      body: JSON.stringify({
        message: text,
        context: cellContext,
        provider: state.provider,
        model: state.model
      }),
      signal: AbortSignal.timeout(30000)
    });

    if (!response.ok) {
      throw new Error(`MCP error: ${response.status}`);
    }

    const data = await response.json();
    return data.response;
  } catch (error) {
    // Fallback to mock if MCP unavailable
    console.warn("MCP call failed, using mock:", error.message);
    return generateMockResponse(text, cellContext?.length || 0);
  }
}

/// Generate mock response (fallback)
function generateMockResponse(prompt, cellCount) {
  const lower = prompt.toLowerCase();

  if (lower.includes("analizar") || lower.includes("análisis")) {
    return `📊 **Análisis de datos**\n\n` +
      `He detectado ${cellCount} celdas en tu selección.\n\n` +
      `**Modo demo activo** — Conecta un MCP Server para análisis real:\n\n` +
      `1. Abre Settings ⚙️ (icono arriba a la derecha)\n` +
      `2. Introduce la URL de tu servidor MCP\n` +
      `3. Pulsa "Probar conexión" para verificar\n\n` +
      `**Endpoints disponibles:**\n` +
      `- POST /chat — Chat con IA\n` +
      `- GET /health — Estado del servidor\n` +
      `- GET /history — Historial de mensajes`;
  }

  if (lower.includes("fórmula") || lower.includes("explicar")) {
    return `📝 **Explicador de fórmulas**\n\n` +
      `**Modo demo** — Conecta el MCP Server para análisis real.\n\n` +
      `**Fórmulas disponibles en Excel:**\n` +
      `- =AIAnalyze(A1:B10) — Analiza un rango\n` +
      `- =AIExplain(A1) — Explica una fórmula\n` +
      `- =AISummarize(A1:A5) — Resume celdas de texto`;
  }

  if (lower.includes("error") || lower.includes("corregir") || lower.includes("bug")) {
    return `🐛 **Depurador de errores**\n\n` +
      `Para análisis completo, conecta el MCP Server:\n\n` +
      `Configuración → MCP Server URL → Probar conexión`;
  }

  if (lower.includes("resumir") || lower.includes("resumen")) {
    return `📋 **Resumidor**\n\n` +
      `${cellCount} celdas analizadas.\n\n` +
      `Conecta el MCP Server para generar resúmenes con IA.`;
  }

  return `🤖 **Excel AI v2**\n\n` +
    `Recibido: "${prompt.slice(0, 50)}${prompt.length > 50 ? "..." : ""}"\n` +
    `Celdas contexto: ${cellCount}\n` +
    `Modo privacidad: ${state.privacyMode ? "🔒 ON" : "🔓 OFF"}\n` +
    `Modelo: ${state.model}\n\n` +
    `💡 Usa ⚙️ para configurar el MCP Server y activar el modo IA real.`;
}

/// Add message to chat
function addMessage(role, content) {
  if (!elements.messages) return;

  // Hide welcome message after first real message
  const welcome = elements.messages.querySelector(".welcome-message");
  if (welcome && role === "assistant") {
    welcome.style.display = "none";
  }

  const div = document.createElement("div");
  div.className = `message ${role}`;
  div.textContent = content;

  elements.messages.appendChild(div);
  elements.messages.scrollTop = elements.messages.scrollHeight;

  state.messages.push({ role, content, timestamp: Date.now() });
}

/// Toggle privacy mode
function togglePrivacy() {
  state.privacyMode = !state.privacyMode;
  localStorage.setItem("privacyMode", state.privacyMode);
  updateStatus(state.privacyMode ? "🔒 Privacidad ON" : "🔓 Privacidad OFF");
}

/// Update status text
function updateStatus(text) {
  if (elements.status) {
    elements.status.textContent = text;
  }
}

/// Open settings modal
function openSettings() {
  initSettingsModal();
  const modal = document.getElementById("settings-modal");
  if (modal) {
    modal.classList.add("show");
    const content = modal.querySelector(".settings-content");
    if (content) content.classList.add("active");
    const endpointInput = modal.querySelector("#mcp-endpoint");
    const providerSelect = modal.querySelector("#provider-select");
    const modelSelect = modal.querySelector("#model-select");
    const statusEl = modal.querySelector("#settings-status");
    if (endpointInput) endpointInput.value = CONFIG.mcpEndpoint;
    const apiKeyInput = modal.querySelector("#api-key");
    if (apiKeyInput) apiKeyInput.value = localStorage.getItem("apiKey") || "";
    if (providerSelect) {
      providerSelect.value = state.provider;
      // Trigger change to populate models
      providerSelect.dispatchEvent(new Event("change"));
    }
    if (modelSelect) setTimeout(() => { modelSelect.value = state.model; }, 0);
    if (statusEl) statusEl.textContent = state.privacyMode ? "🔒 Privacidad" : "🔓 Normal";
  }
}

function closeSettings() {
  const modal = document.getElementById("settings-modal");
  if (modal) modal.classList.remove("show");
  const content = modal?.querySelector(".settings-content");
  if (content) content.classList.remove("active");
}

function saveSettings() {
  const endpoint = document.getElementById("mcp-endpoint")?.value?.trim();
  const provider = document.getElementById("provider-select")?.value;
  const model = document.getElementById("model-select")?.value;

  if (endpoint) {
    localStorage.setItem("mcpEndpoint", endpoint);
    CONFIG.mcpEndpoint = endpoint;
  }
  const apiKey = document.getElementById("api-key")?.value?.trim();
  if (apiKey) {
    localStorage.setItem("apiKey", apiKey);
  }
  if (provider) {
    localStorage.setItem("provider", provider);
    state.provider = provider;
  }
  if (model) {
    localStorage.setItem("model", model);
    state.model = model;
    if (elements.modelDisplay) elements.modelDisplay.textContent = `Modelo: ${model.split(":")[0] || model}`;
  }

  closeSettings();
  checkMCPConnection();
}

async function testConnection() {
  const statusEl = document.getElementById("settings-conn-status");
  if (statusEl) statusEl.textContent = "⏳ Probando...";

  const endpoint = document.getElementById("mcp-endpoint")?.value || CONFIG.mcpEndpoint;

  try {
    const response = await fetch(`${endpoint}/health`, {
      signal: AbortSignal.timeout(5000)
    });

    if (response.ok) {
      const data = await response.json();
      if (statusEl) statusEl.textContent = `✅ Conectado (${data.history || 0} msgs)`;
    } else {
      if (statusEl) statusEl.textContent = `⚠️ HTTP ${response.status}`;
    }
  } catch (error) {
    if (statusEl) statusEl.textContent = `❌ ${error.message.split("\n")[0]}`;
  }
}

/// Settings modal (lazy init)
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
          <label for="mcp-endpoint">MCP Server URL</label>
          <input type="url" id="mcp-endpoint" placeholder="https://tu-tunnel.loca.lt" />
          <small>URL de tu servidor MCP (Flask en Colab)</small>
        </div>
        <div class="settings-group">
          <label for="api-key">API Key</label>
          <input type="password" id="api-key" placeholder="Tu API key (NVIDIA/OpenRouter)" />
          <small>API key que se pasará al servidor MCP</small>
        </div>
        <div class="settings-group">
          <label for="provider-select">Proveedor</label>
          <select id="provider-select">
            <option value="opencode">OpenCode (MiniMax Free)</option>
            <option value="openrouter">OpenRouter (Free)</option>
            <option value="openai">OpenAI</option>
            <option value="anthropic">Anthropic</option>
            <option value="google">Google</option>
            <option value="nvidia">NVIDIA</option>
          </select>
        </div>
        <div class="settings-group">
          <label for="model-select">Modelo</label>
          <select id="model-select">
            <option value="minimax-m2.5-free">MiniMax M2.5 Free</option>
          </select>
        </div>
        <div class="settings-divider"></div>
        <div class="settings-info">
          <p><strong>Estado conexión:</strong> <span id="settings-conn-status">—</span></p>
          <p><strong>Modo:</strong> <span id="settings-status">${state.privacyMode ? "🔒 Privacidad" : "🔓 Normal"}</span></p>
          <p><strong>Versión:</strong> v0.3.1</p>
          <p><strong>Excel:</strong> <span id="excel-status-display">Conectando...</span></p>
        </div>
        <div class="settings-footer">
          <button class="btn-secondary" id="test-conn-btn">🔗 Probar</button>
          <button class="btn-primary" id="save-settings-btn">💾 Guardar</button>
        </div>
      </div>
    </div>
  `;
  document.body.appendChild(modal);

  // Event listeners
  modal.querySelector("#settings-close-btn")?.addEventListener("click", closeSettings);
  modal.querySelector("#save-settings-btn")?.addEventListener("click", saveSettings);
  modal.querySelector("#test-conn-btn")?.addEventListener("click", testConnection);
  modal.addEventListener("click", (e) => {
    if (e.target === modal) closeSettings();
  });

  // Modelos por proveedor
  const MODELS_BY_PROVIDER = {
    "opencode": [
      { id: "minimax-m2.5-free", name: "MiniMax M2.5 Free" }
    ],
    "openrouter": [
      { id: "meta-llama/llama-3.1-8b-instruct:free", name: "Meta Llama 3.1 8B" },
      { id: "qwen/qwen2.5-7b-instruct:free", name: "Qwen 2.5 7B" },
      { id: "mistralai/mistral-7b-instruct:free", name: "Mistral 7B" },
      { id: "google/gemma-3-1b-it:free", name: "Google Gemma 3 1B" },
      { id: "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning:free", name: "Nvidia Nemotron 3 Nano" },
      { id: "google/gemma-4-26b-a4b-it:free", name: "Google Gemma 4 26B" }
    ],
    "openai": [
      { id: "openai/gpt-4o-mini", name: "GPT-4o Mini" },
      { id: "openai/gpt-4o", name: "GPT-4o" },
      { id: "openai/gpt-4-turbo", name: "GPT-4 Turbo" },
      { id: "openai/gpt-3.5-turbo", name: "GPT-3.5 Turbo" },
      { id: "openai/o1-mini", name: "o1 Mini" },
      { id: "openai/o1", name: "o1" }
    ],
    "anthropic": [
      { id: "anthropic/claude-3.5-haiku", name: "Claude 3.5 Haiku" },
      { id: "anthropic/claude-3.5-sonnet", name: "Claude 3.5 Sonnet" },
      { id: "anthropic/claude-3-opus", name: "Claude 3 Opus" },
      { id: "anthropic/claude-3-sonnet", name: "Claude 3 Sonnet" },
      { id: "anthropic/claude-3-haiku", name: "Claude 3 Haiku" },
      { id: "anthropic/claude-2.1", name: "Claude 2.1" }
    ],
    "google": [
      { id: "google/gemini-flash-1.5-8b", name: "Gemini Flash 1.5" },
      { id: "google/gemini-pro-1.5", name: "Gemini Pro 1.5" },
      { id: "google/gemini-1.5-flash-8b", name: "Gemini 1.5 Flash 8B" },
      { id: "google/gemini-1.0-pro", name: "Gemini 1.0 Pro" },
      { id: "google/gemini-2.0-flash-exp", name: "Gemini 2.0 Flash" },
      { id: "google/gemini-2.0-pro-exp", name: "Gemini 2.0 Pro" }
    ],
    "nvidia": [
      { id: "minimaxai/minimax-m2.7", name: "MiniMax M2.7" },
      { id: "nvidia/llama-3.1-nemotron-70b-instruct:free", name: "Nemotron 70B" },
      { id: "nvidia/llama-3.3-70b-instruct:free", name: "Llama 3.3 70B" }
    ]
  };

  // Cuando cambia el proveedor, actualizar modelos
  const providerSelect = modal.querySelector("#provider-select");
  const modelSelect = modal.querySelector("#model-select");
  providerSelect?.addEventListener("change", function() {
    const models = MODELS_BY_PROVIDER[this.value] || [];
    modelSelect.innerHTML = models.map(m => `<option value="${m.id}">${m.name}</option>`).join("");
  });

  // Excel status
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

// Export for console access
window.ExcelAI = {
  sendMessage,
  getCellContext,
  togglePrivacy,
  openSettings,
  closeSettings,
  testConnection,
  state,
  CONFIG
};