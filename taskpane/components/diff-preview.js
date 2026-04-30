/**
 * DiffPreview Component
 * Shows differences between AI suggestions and original data
 */

export function initDiffPreview(container) {
  const template = `
    <div class="diff-preview">
      <div class="diff-header">
        <h3>📊 Vista de Diferencias</h3>
        <button class="diff-close">&times;</button>
      </div>
      <div class="diff-controls">
        <select id="diff-type">
          <option value="cells">Celdas</option>
          <option value="formulas">Fórmulas</option>
          <option value="values">Valores</option>
        </select>
        <button class="diff-apply">Aplicar Cambios</button>
        <button class="diff-revert">Revertir</button>
      </div>
      <div class="diff-content"></div>
    </div>
  `;
  
  container.innerHTML = template;
  
  // Event listeners
  container.querySelector(".diff-close").addEventListener("click", () => {
    container.style.display = "none";
  });
  
  return {
    show: function(original, suggested) {
      renderDiff(container.querySelector(".diff-content"), original, suggested);
      container.style.display = "block";
    },
    hide: function() {
      container.style.display = "none";
    }
  };
}

function renderDiff(container, original, suggested) {
  constrows = [];
  
  for (let i = 0; i < Math.min(original.length, suggested.length, 20); i++) {
    const orig = original[i];
    constSug = suggested[i];
    const isChanged = orig.value !== sug.value;
    
    rows.push(`
      <tr class="${isChanged ? 'changed' : ''}">
        <td class="address">${orig.address}</td>
        <td class="original">${orig.value}</td>
        <td class="arrow">→</td>
        <td class="suggested">${sug.value}</td>
      </tr>
    `);
  }
  
  container.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Celda</th>
          <th>Original</th>
          <th></th>
          <th>Cambio</th>
        </tr>
      </thead>
      <tbody>${rows.join("")}</tbody>
    </table>
  `;
}

// Auto-register if in browser
if (typeof window !== "undefined") {
  window.DiffPreview = { init: initDiffPreview };
}
