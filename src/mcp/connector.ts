/**
 * MCP Connector for OpenCode
 * Bridges Excel with OpenCode's Model Context Protocol
 */
import { spawn } from "child_process";
import type { CellContext, AIMessage } from "../types/";
import { anonymizeRange } from "../services/privacy";
import type { PrivacySettings } from "../types/";
import { logger } from "../utils/logger";
import path from "path";

export class OpenCodeMCPConnector {
  private process: ReturnType<typeof spawn> | null = null;
  private endpoint: string;
  private model: string;
  
  constructor(endpoint: string = "http://192.168.1.44:3749", model: string = "opencode") {
    this.endpoint = endpoint;
    this.model = model;
  }
  
  async connect(): Promise<void> {
    logger.info("Connecting to OpenCode MCP...", { endpoint: this.endpoint });
    // TODO: Implement MCP connection
    // - Spawn opencode process with --mcp flag
    // - Setup stdio communication
    // - Handle JSON-RPC messages
    this.process = spawn("opencode", ["--mcp", "--model", this.model], {
      stdio: ["pipe", "pipe", "pipe"],
    });
  }
  
  async sendMessage(message: AIMessage, context: CellContext[], privacy: PrivacySettings): Promise<string> {
    // Apply privacy layer before sending
    const anonymizedCells = anonymizeRange(context, privacy);
    const systemPrompt = this.buildSystemPrompt(anonymizedCells);
    
    logger.info("Sending message to AI", { 
      messageLength: message.content.length,
      cellsIncluded: anonymizedCells.length,
      privacyLayer: privacy.layer 
    });
    
    // TODO: Send to MCP and get response
    return `[MOCK] Respuesta IA simulada para: ${message.content.slice(0, 50)}...`;
  }
  
  private buildSystemPrompt(cells: CellContext[]): string {
    if (cells.length === 0) return "You are an Excel AI assistant."; // Default
    
    const cellsSummary = cells
      .slice(0, 10) // Limit context size
      .map(c => `  ${c.address}: ${c.value}`)
      .join("\n");
    
    return `You are an Excel AI assistant. Current spreadsheet context:\n${cellsSummary}\n\nProvide concise, actionable help.`; // Context injection
  }
  
  async disconnect(): Promise<void> {
    this.process?.kill();
    this.process = null;
    logger.info("Disconnected from OpenCode MCP");
  }
}
