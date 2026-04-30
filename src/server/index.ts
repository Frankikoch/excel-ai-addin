/** Server entry - Node.js backend for AI */ 
import { anonymizeRange } from "./privacy";
import { OpenCodeMCPConnector } from "./connector";
import type { AIMessage, CellContext, PrivacySettings } from "../shared/types";
export class AIService {
  private connector: OpenCodeMCPConnector;
  constructor(endpoint: string = "http://192.168.1.44:3749") { 
    this.connector = new OpenCodeMCPConnector(endpoint); 
  }
  async chat(message: string, cells: CellContext[], privacy: PrivacySettings): Promise<string> {
    const msg: AIMessage = { role: "user", content: message, cells, timestamp: Date.now() };
    return this.connector.sendMessage(msg, cells, privacy);
  }
}
