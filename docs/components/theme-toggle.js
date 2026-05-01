/**
 * Theme Toggle Component
 * Switch between Dark and Light mode
 */

export function initThemeToggle(container) {
  const isDark = localStorage.getItem("theme") !== "light";
  
  container.innerHTML = `
    <div class="theme-toggle ${isDark ? '' : 'light'}">
      <span>🌙</span>
      <div class="toggle-track">
        <div class="toggle-thumb"></div>
      </div>
      <span>☀️</span>
    </div>
  `;
  
  const toggle = container.querySelector(".theme-toggle");
  
  toggle.addEventListener("click", function() {
    toggle.classList.toggle("light");
    const isLight = toggle.classList.contains("light");
    
    localStorage.setItem("theme", isLight ? "light" : "dark");
    document.documentElement.setAttribute("data-theme", isLight ? "light" : "dark");
  });
  
  // Apply initial theme
  document.documentElement.setAttribute("data-theme", isDark ? "dark" : "light");
  
  return {
    isDark: function() { return !toggle.classList.contains("light"); },
    toggle: function() { toggle.click(); }
  };
}

// Auto-register
if (typeof window !== "undefined") {
  window.ThemeToggle = { init: initThemeToggle };
}
