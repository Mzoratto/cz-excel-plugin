import { ChatMessage, ChatRole } from "./types";

let idCounter = 0;

function nextId(): string {
  idCounter += 1;
  return `msg-${Date.now()}-${idCounter}`;
}

function nowISO(): string {
  return new Date().toISOString();
}

export class ChatSession {
  private history: ChatMessage[] = [];

  getMessages(): ChatMessage[] {
    return [...this.history];
  }

  addMessage(role: ChatRole, content: string, metadata?: Record<string, unknown>): ChatMessage {
    const message: ChatMessage = {
      id: nextId(),
      role,
      content,
      timestamp: nowISO(),
      metadata
    };
    this.history.push(message);
    return message;
  }

  reset(): void {
    this.history = [];
  }
}
