/**
 * ModelSelector Component
 * Switch between OpenCode, Claude, Gemini
 */

export function initModelSelector(container) {
  const models = [
    { id: "opencode", name: "OpenCode", icon: "🤖", desc: "IA local" },
    { id: "claude", name: "Claude", icon: "📘", desc: "Anthropic" },
    { id: "gemini", name: "Gemini", icon: "💎", desc: "Google" },
    { id: "openrouter", name: "OpenRouter", icon: "🔗", desc: "Múltiples modelos" }
  ];
  
  const current = localStorage.getItem("model") || "opencode";
  
  let html = '<div class="model-selector">';
  html += '<div class="model-current">';
  html += '<span class="model-icon">' + models.find(m => m.id === current)?.icon + '</span>';
  html += '<span class="model-name">' + models.find(m => m.id === current)?.name + '</span>';
  html += '<button class="model-toggle">▼</button>';
  html += '</div>';
  html += '<div class="model-dropdown">';
  
  for (const model of models) {
    html += '<div class="model-option' + (model.id === current ? ' active' : '') + '" data-model="' + model.id + '">';
    html += '<span class="model-icon">' + model.icon + '</span>';
    html += '<span class="model-info"><span class="model-name">' + model.name + '</span>';
    html += '<span class="model-desc">' + model.desc + '</span></span>';
    html += '</div>';
  }
  
  html += '</div></div>';
  
  container.innerHTML = html;
  
  const dropdown = container.querySelector(".model-dropdown");
  const toggle = container.querySelector(".model-toggle");
  const currentEl = container.querySelector(".model-current");
  
  toggle.addEventListener("click", function(e) {
    e.stopPropagation();
    dropdown.classList.toggle("show");
  });
  
  currentEl.addEventListener("click", function(e) {
    e.stopPropagation();
    dropdown.classList.toggle("show");
  });
  
  document.addEventListener("click", function() {
    dropdown.classList.remove("show");
  });
  
  container.querySelectorAll(".model-option").forEach(function(opt) {
    opt.addEventListener("click", function() {
      selectModel(this.dataset.model, models);
    });
  });
  
  return {
    getModel: function() { return localStorage.getItem("model") || "opencode"; },
    setModel: function(id) { selectModel(id, models); }
  };
}

function selectModel(id, models) {
  localStorage.setItem("model", id);
  
  const model = models.find(m => m.id === id);
  document.querySelectorAll(".model-option").forEach(function(opt) {
    opt.classList.toggle("active", opt.dataset.model === id);
  });
  
  const current = document.querySelector(".model-current");
  current.querySelector(".model-icon").textContent = model.icon;
  current.querySelector(".model-name").textContent = model.name;
  
  document.querySelector(".model-dropdown").classList.remove("show");
  
  // Notify app
  window.dispatchEvent(new CustomEvent("modelChanged", { detail: { model: id } }));
}

// Auto-register
if (typeof window !== "undefined") {
  window.ModelSelector = { init: initModelSelector };
}
