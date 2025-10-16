export type ChatRole = "user" | "assistant" | "system" | "action" | "error";

export interface ChatMessage {
  id: string;
  role: ChatRole;
  content: string;
  timestamp: string;
  metadata?: Record<string, unknown>;
}

export interface ChatConfig {
  endpoint?: string;
  apiKey?: string;
  timeoutMs?: number;
}

export interface AssistantReply {
  role: ChatRole;
  content: string;
  metadata?: Record<string, unknown>;
}
