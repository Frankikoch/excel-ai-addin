/**
 * Excel AI Add-in - Main Application
 */

// State
const state = {
  privacyMode: true,
  model: "opencode",
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
  settingsBtn: document.getElementById("settings-btn")
};

// Initialize
document.addEventListener("DOMContentLoaded", init);

async function init() {
  console.log("🤖 Excel AI initializing...");
  
  // Event listeners
  elements.sendBtn?.addEventListener("click", sendMessage);
  elements.userInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });
  elements.privacyToggle?.addEventListener("change", togglePrivacy);
  
  // Quick action buttons
  document.querySelectorAll(".quick-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const prompt = btn.dataset.prompt;
      sendMessage(prompt);
    });
  });
  
  // Connect to Excel
  await connectToExcel();
  
  updateStatus("🟢 Listo");
}

/// Connect to Excel via Office.js
async function connectToExcel() {
  try {
    if (typeof Excel !== "undefined") {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await context.sync();
        console.log("📊 Connected to:", sheet.name);
      });
    }
  } catch (error) {
    console.warn("Excel context not available:", error.message);
  }
}

/// Send message to AI
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
    
    // Send to MCP server (mock for now)
    const response = await callAI(text, context);
    
    // Add assistant response
    addMessage("assistant", response);
  } catch (error) {
    addMessage("assistant", "❌ Error: " + error.message);
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
      range.load(["address", "values", "formulas"]);
      await excelContext.sync();
      
      const cells = [];
      const values = range.values;
      const address = range.address;
      
      for (let i = 0; i < Math.min(values.length, 10); i++) {
        cells.push({
          address: address,
          value: values[i][0],
          dataType: typeof values[i][0]
        });
      }
      return cells;
    });
    
    return context;
  } catch {
    return [];
  }
}

/// Call AI (MCP connector)
async function callAI(text, cellContext) {
  // MVP: Simulated response
  // TODO: Connect to real MCP server on Raspberry
  if (state.privacyMode) {
    // Anonymize before sending
    cellContext = cellContext.map(c => ({
      ...c,
      value: String(c.value).replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[EMAIL]")
    }));
  }
  
  // Mock response (MVP)
  return generateMockResponse(text, cellContext.length);
}

/// Generate mock response (MVP)
function generateMockResponse(prompt, cellCount) {
  const lower = prompt.toLowerCase();
  
  if (lower.includes("analizar") || lower.includes("análisis")) {
    return `📊 **Análisis**\n\n` +
      `Tienes ${cellCount} celdas seleccionadas.\n\n` +
      `Para un análisis completo, puedo:\n` +
      `- Calcular estadísticas\n` +
      `- Identificar tendencias\n` +
      `- Detectar outliers\n` +
      `- Generar gráficos\n\n` +
      `¿Qué análisis necesitas?`;
  }
  
  if (lower.includes("fórmula") || lower.includes("explicar")) {
    return `📝 **Explicación de fórmula**\n\n` +
      `Selecciona una celda con fórmula yAsk me to explain it.\n` +
      `Puedo analizar:\n` +
      `- Lógica de la fórmula\n` +
      `- Referencias circulares\n` +
      `- Errores #REF!, #VALUE!\n` +
      `- Sugerir alternativas`;
  }
  
  if (lower.includes("error") || lower.includes("corregir")) {
    return `🐛 **Depuración**\n\n` +
      `Ejecutaré un análisis de errores en la hoja.\n\n` +
      `Errores comunes:\n` +
      `- #DIV/0: división por cero\n` +
      `- #REF!: referencia inválida\n` +
      `- #VALUE!: tipo incorrecto\n` +
      `- #N/A: valor no encontrado`;
  }
  
  return `🤖 **Entendido**\n\n` +
    `He recibido: "${prompt.slice(0, 50)}..."\n\n` +
    `Contexto actual: ${cellCount} celdas\n` +
    `Modo privacidad: ${state.privacyMode ? "🔒 Activado" : "🔓 Desactivado"}\n\n` +
    `¿En qué puedo ayudarte con tu hoja de cálculo?`;
}

/// Add message to chat
function addMessage(role, content) {
  if (!elements.messages) return;
  
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
  updateStatus(state.privacyMode ? "🔒 Privacidad ON" : "🔓 Privacidad OFF");
}

/// Update status bar
function updateStatus(text) {
  if (elements.status) {
    elements.status.textContent = text;
  }
}

// Export for debugging
window.excelAI = { state, sendMessage, getCellContext };
