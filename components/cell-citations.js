/**
 * CellCitations Component
 * Shows which cells AI used for its response
 */

export function initCellCitations(container) {
  const template = `
    <div class="cell-citations">
      <div class="citations-header">
        <h3>📎 Celdas Referenciadas</h3>
        <button class="citations-toggle">▼</button>
      </div>
      <div class="citations-list"></div>
    </div>
  `;
  
  container.innerHTML = template;
  
  const list = container.querySelector(".citations-list");
  const toggle = container.querySelector(".citations-toggle");
  
  toggle.addEventListener("click", () => {
    list.classList.toggle("collapsed");
    toggle.textContent = list.classList.contains("collapsed") ? "▶" : "▼";
  });
  
  return {
    show: function(cells) {
      renderCitations(list, cells);
    }
  };
}

function renderCitations(list, cells) {
  if (!cells || cells.length === 0) {
    list.innerHTML = "<p class='no-citations'>Sin celdas referenciadas</p>";
    return;
  }
  
  list.innerHTML = cells.map(cell => 
    '<span class="citation-tag" data-address="' + cell.address + '">' + cell.address + "</span>"
  ).join("");
  
  // Click to highlight cell
  list.querySelectorAll(".citation-tag").forEach(function(tag) {
    tag.addEventListener("click", function() {
      highlightCell(this.dataset.address);
    });
  });
}

function highlightCell(address) {
  // Tell Excel to select this cell
  if (typeof Excel !== "undefined") {
    Excel.run(function(ctx) {
      var range = ctx.workbook.getRange(address);
      range.select();
      return ctx.sync();
    });
  }
}

// Auto-register
if (typeof window !== "undefined") {
  window.CellCitations = { init: initCellCitations };
}
