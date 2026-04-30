/** OpenCode MCP Connector */ 
import type { AIMessage, CellContext, PrivacySettings } from "../shared/types";
import { anonymizeRange } from "./privacy";
export class OpenCodeMCPConnector {
  constructor(private endpoint: string) {}
  async sendMessage(message: AIMessage, cells: CellContext[], _privacy: PrivacySettings): Promise<string> {
    const a = anonymizeRange(cells);
    console.log("MCP", message.content.slice(0, 30), "| cells:", a.length);
    return `[AI] Response to: ${message.content.slice(0, 50)}...`;
  }
}
