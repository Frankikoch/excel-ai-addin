/** Privacy Layer - anonymizes before AI */ 
import type { CellContext, PrivacySettings } from "../shared/types";
export function anonymizeCell(cell: CellContext): CellContext {
  const value = String(cell.value).replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[EMAIL]")
    .replace(/\b\d{3}[-.]?\d{2}[-.]?\d{4}\b/g, "[SSN]").replace(/\b\d{10,}\b/g, "[ID]");
  return { ...cell, value, formula: undefined };
}
export function anonymizeRange(cells: CellContext[]): CellContext[] { 
  return cells.map(anonymizeCell); 
}
