/** Shared type definitions */ export interface CellContext {
  address: string;
  value: unknown; formula?: string; dataType: string;
}
export interface AIMessage {
  role: "user" | "assistant";
  content: string; cells?: CellContext[];
  timestamp: number; }
export interface PrivacySettings {
  layer: "mvp" | "advanced"; anonymizeFormulas: boolean; auditEnabled: boolean;
}
export interface AddinConfig {
  model: "opencode" | "claude" | "gemini"; privacy: PrivacySettings; raspberryEndpoint?: string;
}
