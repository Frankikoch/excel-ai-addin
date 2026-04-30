/**
 * Excel AI Functions
 * Custom formulas that can be used directly in Excel cells
 * 
 * Usage in Excel:
 * =AIAnalyze(A1:B10)
 * =AIExplain(A1)
 * =AICorrect(A1)
 * =AISummarize(A1:A5)
 */

// Register custom functions with Excel
if (typeof Excel !== "undefined") {
  Excel.run(async (context) => {
    // Register custom functions
    await context.workbook.functions.register(
      "AIAnalyze",
      async (range) => {
        // Get range values
        const values = range.values;
        const analysis = analyzeData(values);
        return analysis;
      }
    );
    
    await context.workbook.functions.register(
      "AIExplain",
      async (cell) => {
        const formula = cell.formula;
        return explainFormula(formula);
      }
    );
    
    console.log("📝 AI Functions registered");
  });
}

/**
 * Analyze data in range
 */
function analyzeData(values) {
  let numbers = [];
  let text = [];
  let empty = 0;
  
  for (let row of values) {
    for (let cell of row) {
      if (cell === "" || cell === null || cell === undefined) {
        empty++;
      } else if (typeof cell === "number") {
        numbers.push(cell);
      } else if (!isNaN(Number(cell))) {
        numbers.push(Number(cell));
      } else {
        text.push(String(cell));
      }
    }
  }
  
  const summary = [];
  if (numbers.length > 0) {
    const sum = numbers.reduce((a, b) => a + b, 0);
    const avg = sum / numbers.length;
    const min = Math.min(...numbers);
    const max = Math.max(...numbers);
    
    summary.push(`📊 Números: ${numbers.length}`);
    summary.push(`📈 Suma: ${sum.toFixed(2)}`);
    summary.push(`📉 Media: ${avg.toFixed(2)}`);
    summary.push(`⬇️ Mín: ${min}`);
    summary.push(`⬆️ Máx: ${max}`);
  }
  
  if (text.length > 0) {
    summary.push(`📝 Texto: ${text.length} celdas`);
  }
  
  summary.push(`⬜ Vacías: ${empty}`);
  
  return summary.join("\n");
}

/**
 * Explain a formula
 */
function explainFormula(formula) {
  if (!formula || formula[0] !== "=") {
    return "⚠️ No es una fórmula (no empieza por =)";
  }
  
  const parts = [];
  const funcMatch = formula.match(/([A-Z]+)\(/gi);
  
  if (funcMatch) {
    parts.push(`Funciones detectadas: ${funcMatch.join(", ")}`);
  }
  
  const refMatch = formula.match(/[A-Z]+[0-9]+/g);
  if (refMatch) {
    parts.push(`Celdas referenciadas: ${[...new Set(refMatch)].join(", ")}`);
  }
  
  if (formula.includes("+")) parts.push("Operación: Suma");
  if (formula.includes("-")) parts.push("Operación: Resta");
  if (formula.includes("*")) parts.push("Operación: Multiplicación");
  if (formula.includes("/")) parts.push("Operación: División");
  
  return parts.length > 0 
    ? parts.join("\n") 
    : "Fórmula detectada sin operaciones reconocibles";
}

/**
 * AI Summarize - Create summary of selected cells
 */
function summarizeCells(values) {
  let allText = [];
  
  for (let row of values) {
    for (let cell of row) {
      if (cell && typeof cell === "string" && cell.trim()) {
        allText.push(cell.trim());
      }
    }
  }
  
  if (allText.length === 0) {
    return "No hay texto para resumir";
  }
  
  // Simple word frequency
  const words = allText.join(" ").toLowerCase().split(/\s+/);
  const wordFreq = {};
  
  for (let word of words) {
    if (word.length > 3) {
      wordFreq[word] = (wordFreq[word] || 0) + 1;
    }
  }
  
  const topWords = Object.entries(wordFreq)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([word, count]) => `${word} (${count})`);
  
  return `📋 **Resumen**\n\nCeldas analizadas: ${allText.length}\n\nPalabras más frecuentes:\n${topWords.join("\n")}`;
}

// Export for use in app
if (typeof window !== "undefined") {
  window.AIFormulas = {
    analyze: analyzeData,
    explain: explainFormula,
    summarize: summarizeCells
  };
}
