/**
 * AuditLog Component
 * Logs all AI interactions for privacy compliance
 */

export function initAuditLog(container) {
  const template = `
    <div class="audit-log">
      <div class="audit-header">
        <h3>📋 Registro de Auditoría</h3>
        <button class="audit-export">Exportar</button>
      </div>
      <div class="audit-list"></div>
    </div>
  `;
  
  container.innerHTML = template;
  
  let entries = [];
  
  container.querySelector(".audit-export").addEventListener("click", function() {
    const blob = new Blob([JSON.stringify(entries, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "audit-" + new Date().toISOString().split("T")[0] + ".json";
    a.click();
  });
  
  return {
    log: function(action, details) {
      entries.push({
        timestamp: new Date().toISOString(),
        action: action,
        details: details
      });
      renderEntries(container.querySelector(".audit-list"), entries);
    },
    getEntries: function() { return entries; }
  };
}

function renderEntries(container, entries) {
  container.innerHTML = entries.slice(-10).reverse().map(function(entry) {
    return '<div class="audit-entry"><span class="time">' + 
      entry.timestamp.split("T")[1].split(".")[0] + 
      '</span><span class="action">' + entry.action + "</span></div>";
  }).join("");
}

// Auto-register
if (typeof window !== "undefined") {
  window.AuditLog = { init: initAuditLog };
}
