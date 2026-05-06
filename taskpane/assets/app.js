/**
 * Excel AI Add-in - Main Application (v2 - with MCP connection)
 */

// Configuration
const CONFIG = {
  mcpEndpoint: localStorage.getItem("mcpEndpoint") || "http://192.168.1.44:3749",
  excelMcpEndpoint: localStorage.getItem("excelMcpEndpoint") || "http://localhost:8765",
  model: localStorage.getItem("model") || "opencode"
};

// State
const state = {
  privacyMode: localStorage.getItem("privacyMode") !== "false",
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
  elements.sendBtn?.addEventListener("click", () => sendMessage());
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
  
  // Excel automation buttons (when Excel MCP is selected)
  setupExcelAutomationButtons();
  
  // Connect to Excel
  await connectToExcel();
  
  // Check MCP connection
  await checkMCPConnection();
  
  updateStatus("🟢 Listo");

  // Sync header model selector -> app state
  window.addEventListener("modelChanged", function(e) {
    const model = e?.detail?.model;
    if (!model) return;
    state.model = model;
    localStorage.setItem("model", model);
    if (elements.modelDisplay) elements.modelDisplay.textContent = `Modelo: ${model}`;
  });
}

/// Check MCP server connection
async function checkMCPConnection() {
  // Check regular AI MCP
  try {
    const response = await fetch(`${CONFIG.mcpEndpoint}/health`, {
      method: "GET",
      signal: AbortSignal.timeout(5000)
    });
    
    if (response.ok) {
      const data = await response.json();
      updateStatus(`🟢 Conectado IA (${data.history} msgs)`);
      console.log("✅ AI MCP Server connected");
    }
  } catch (error) {
    console.warn("AI MCP Server not available:", error.message);
  }
  
  // Check Excel MCP connection
  if (state.model === "excel-mcp") {
    try {
      const excelResponse = await fetch(`${CONFIG.excelMcpEndpoint}/health`, {
        method: "GET",
        signal: AbortSignal.timeout(5000)
      });
      
      if (excelResponse.ok) {
        updateStatus("🟢 Excel MCP conectado");
        console.log("✅ Excel MCP connected");
      }
    } catch (error) {
      console.warn("Excel MCP not available:", error.message);
      updateStatus("⚠️ Excel MCP no disponible");
    }
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
  elements.userInput.value = "";
  
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

/// Call MCP Server (AI or Excel Automation)
async function callMCP(text, cellContext) {
  // Use Excel MCP for automation commands
  if (state.model === "excel-mcp") {
    return await callExcelMCP(text, cellContext);
  }
  
  // Default: Use regular AI MCP
  try {
    const response = await fetch(`${CONFIG.mcpEndpoint}/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        message: text,
        context: cellContext
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

/// Call Excel MCP Server for automation
async function callExcelMCP(text, cellContext) {
  try {
    const response = await fetch(`${CONFIG.excelMcpEndpoint}/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        message: text,
        context: cellContext
      }),
      signal: AbortSignal.timeout(60000)
    });
    
    if (!response.ok) {
      throw new Error(`Excel MCP error: ${response.status}`);
    }
    
    const data = await response.json();
    return data.response || data.message || "Comando enviado a Excel MCP";
  } catch (error) {
    return "❌ Excel MCP no disponible.\n\n" +
      "Para usar automatización de Excel:\n" +
      "1. Descarga mcp-excel.exe de:\n" +
      "   https://github.com/sbroenne/mcp-server-excel/releases\n" +
      "2. Ejecuta: node excel-mcp-bridge.js\n" +
      "3. Asegúrate de tener Excel abierto";
  }
}

/// Generate mock response (fallback)
function generateMockResponse(prompt, cellCount) {
  const promptStr = typeof prompt === 'string' ? prompt : String(prompt || '');
  const lower = promptStr.toLowerCase();
  
  if (lower.includes("analizar") || lower.includes("análisis")) {
    return `📊 **Análisis**\n\nCeldas seleccionadas: ${cellCount}\n\n` +
      `Para análisis completo, conecta con MCP Server:\n` +
      `node src/index.js\n\n` +
      `Endpoints disponibles:\n` +
      `- POST /chat - Chat con IA\n` +
      `- GET /history - Historial\n` +
      `- GET /health - Estado`;
  }
  
  if (lower.includes("fórmula") || lower.includes("explicar")) {
    return `📝 **Explicador de fórmulas**\n\n` +
      `Conecta el MCP Server para análisis real.\n\n` +
      `Comandos disponibles en Excel:\n` +
      `- =AIAnalyze(A1:B10)\n` +
      `- =AIExplain(A1)\n` +
      `- =AISummarize(A1:A5)`;
  }
  
  if (lower.includes("error") || lower.includes("corregir")) {
    return `🐛 **Depurador de errores**\n\n` +
      `Inicia MCP Server para análisis:\n\n` +
      `cd ~/Documents/excel-mcp-server\n` +
      `npm start`;
  }
  
  return `🤖 **Excel AI v2**\n\n` +
    `Mensaje: "${prompt.slice(0, 40)}..."\n` +
    `Celdas: ${cellCount}\n` +
    `Privacidad: ${state.privacyMode ? "🔒 ON" : "🔓 OFF"}\n\n` +
    `Inicia MCP Server para conexión real con OpenCode`;
}

/// Setup Excel automation buttons
function setupExcelAutomationButtons() {
  const excelBtns = document.querySelectorAll('.excel-mcp-btn');
  const excelPrompts = {
    pivot: "Crea una PivotTable con los datos seleccionados mostrando totales por categoría",
    chart: "Crea un gráfico de barras con los datos seleccionados",
    table: "Convierte este rango en una tabla de Excel con estilo",
    formula: "Agrega una fórmula para calcular totales en esta columna"
  };
  
  excelBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const cmd = btn.dataset.cmd;
      if (excelPrompts[cmd]) {
        // Switch to Excel MCP if not already selected
        if (state.model !== "excel-mcp") {
          state.model = "excel-mcp";
          localStorage.setItem("model", "excel-mcp");
          if (elements.modelDisplay) elements.modelDisplay.textContent = "Modelo: Excel MCP";
        }
        sendMessage(excelPrompts[cmd]);
      }
    });
  });
  
  // Toolbar command buttons
  document.querySelectorAll('.toolbar-btn[data-cmd]').forEach(btn => {
    if (!btn.classList.contains('excel-mcp-btn')) {
      btn.addEventListener('click', () => {
        const cmd = btn.dataset.cmd;
        const prompts = {
          analyze: "Analiza los datos seleccionados y sugiere insights",
          explain: "Explica la fórmula o valor de la celda activa",
          fix: "Identifica y corrige errores en esta hoja",
          summarize: "Resume los datos y estructura de esta hoja"
        };
        if (prompts[cmd]) sendMessage(prompts[cmd]);
      });
    }
  });
}

/// Open settings modal
function openSettings() {
  const existing = document.querySelector('.settings-modal');
  if (existing) {
    existing.remove();
    return;
  }
  
  const modal = document.createElement('div');
  modal.className = 'settings-modal';
  modal.innerHTML = `
    <div class="settings-content">
      <h2>⚙️ Configuración</h2>
      <div class="settings-group">
        <label>Endpoint IA (MCP)</label>
        <input type="text" id="mcp-endpoint-input" value="${CONFIG.mcpEndpoint}" placeholder="http://192.168.1.44:3749">
      </div>
      <div class="settings-group">
        <label>Endpoint Excel MCP</label>
        <input type="text" id="excel-mcp-endpoint-input" value="${CONFIG.excelMcpEndpoint}" placeholder="http://localhost:8765">
      </div>
      <div class="settings-group">
        <label>API Key (NVIDIA/OpenAI)</label>
        <input type="password" id="api-key-input" value="${localStorage.getItem('apiKey') || ''}" placeholder="Tu API key">
      </div>
      <div class="settings-actions">
        <button id="save-settings">💾 Guardar</button>
        <button id="close-settings">✖️ Cerrar</button>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
  
  modal.querySelector('#save-settings').addEventListener('click', () => {
    const mcpEndpoint = document.getElementById('mcp-endpoint-input').value;
    const excelMcpEndpoint = document.getElementById('excel-mcp-endpoint-input').value;
    const apiKey = document.getElementById('api-key-input').value;
    
    localStorage.setItem('mcpEndpoint', mcpEndpoint);
    localStorage.setItem('excelMcpEndpoint', excelMcpEndpoint);
    localStorage.setItem('apiKey', apiKey);
    
    CONFIG.mcpEndpoint = mcpEndpoint;
    CONFIG.excelMcpEndpoint = excelMcpEndpoint;
    
    updateStatus("✅ Configuración guardada");
    modal.remove();
    checkMCPConnection();
  });
  
  modal.querySelector('#close-settings').addEventListener('click', () => modal.remove());
  modal.addEventListener('click', (e) => {
    if (e.target === modal) modal.remove();
  });
}

/// Add message to chat
function addMessage(role, content) {
  if (!elements.messages) return;
  
  // Don't show welcome message after first real message
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

// Export for console access
window.ExcelAI = {
  sendMessage,
  getCellContext,
  togglePrivacy,
  state,
  CONFIG
};

