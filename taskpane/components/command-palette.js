/**
 * CommandPalette Component (Ctrl+K)
 * Quick access to all AI commands
 */

export function initCommandPalette() {
  const commands = [
    { id: "analyze", name: "📊 Analizar datos", shortcut: "Ctrl+Shift+A", action: "Analiza esta selección" },
    { id: "explain", name: "📝 Explicar fórmula", shortcut: "Ctrl+Shift+E", action: "Explica esta fórmula" },
    { id: "fix", name: "🐛 Corregir errores", shortcut: "Ctrl+Shift+F", action: "Corrige los errores" },
    { id: "summarize", name: "📋 Resumir selección", shortcut: "Ctrl+Shift+S", action: "Resume estas celdas" },
    { id: "format", name: "🎨 Formatear", shortcut: "Ctrl+Shift+O", action: "Aplica formato profesional" },
    { id: "chart", name: "📈 Crear gráfico", shortcut: "Ctrl+Shift+G", action: "Crea un gráfico" },
    { id: "template", name: "📄 Usar plantilla", shortcut: "Ctrl+Shift+T", action: "Aplica plantilla" },
    { id: "clear", name: "🗑️ Limpiar celdas", shortcut: "Ctrl+Shift+X", action: "Limpia la selección" },
    { id: "audit", name: "🔍 Auditar", shortcut: "Ctrl+Shift+D", action: "Muestra auditoría" },
    { id: "settings", name: "⚙️ Configuración", shortcut: "Ctrl+,", action: "Abre configuración" }
  ];
  
  // Create modal
  const modal = document.createElement("div");
  modal.className = "command-palette-modal";
  modal.innerHTML = `
    <div class="command-palette">
      <input type="text" class="command-input" placeholder="Escribe un comando..." />
      <div class="command-list"></div>
    </div>
  `;
  
  document.body.appendChild(modal);
  
  const input = modal.querySelector(".command-input");
  const list = modal.querySelector(".command-list");
  
  function renderCommands(filter) {
    const filtered = filter 
      ? commands.filter(c => c.name.toLowerCase().includes(filter.toLowerCase()))
      : commands;
    
    list.innerHTML = filtered.map((c, i) => 
      '<div class="command-item' + (i === 0 ? ' selected' : '') + '" data-id="' + c.id + '">' +
        '<span class="command-name">' + c.name + '</span>' +
        '<span class="command-shortcut">' + c.shortcut + '</span>' +
      '</div>'
    ).join("");
    
    list.querySelectorAll(".command-item").forEach(function(item) {
      item.addEventListener("click", function() {
        executeCommand(this.dataset.id, commands);
      });
    });
  }
  
  function executeCommand(id, commands) {
    const cmd = commands.find(c => c.id === id);
    if (cmd && window.excelAI) {
      window.excelAI.sendMessage(cmd.action);
    }
    close();
  }
  
  function open() {
    modal.classList.add("show");
    input.value = "";
    input.focus();
    renderCommands("");
  }
  
  function close() {
    modal.classList.remove("show");
  }
  
  // Keyboard shortcut
  document.addEventListener("keydown", function(e) {
    if ((e.ctrlKey || e.metaKey) && e.key === "k") {
      e.preventDefault();
      open();
    }
    if (e.key === "Escape" && modal.classList.contains("show")) {
      close();
    }
  });
  
  // Input handling
  input.addEventListener("input", function() {
    renderCommands(this.value);
  });
  
  input.addEventListener("keydown", function(e) {
    const selected = list.querySelector(".command-item.selected");
    const allItems = list.querySelectorAll(".command-item");
    const currentIndex = Array.from(allItems).indexOf(selected);
    
    if (e.key === "ArrowDown") {
      e.preventDefault();
      selected?.classList.remove("selected");
      allItems[Math.min(currentIndex + 1, allItems.length - 1)]?.classList.add("selected");
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      selected?.classList.remove("selected");
      allItems[Math.max(currentIndex - 1, 0)]?.classList.add("selected");
    } else if (e.key === "Enter") {
      e.preventDefault();
      selected && executeCommand(selected.dataset.id, commands);
    }
  });
  
  // Click outside to close
  modal.addEventListener("click", function(e) {
    if (e.target === modal) close();
  });
  
  return {
    open: open,
    close: close,
    getCommands: function() { return commands; }
  };
}

// Auto-register
if (typeof window !== "undefined") {
  window.CommandPalette = { init: initCommandPalette };
}
