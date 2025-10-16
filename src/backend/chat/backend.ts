import { ChatMessage, AssistantReply, ChatConfig } from "./types";

const DEFAULT_TIMEOUT = 15000;

export class ChatBackend {
  private readonly endpoint?: string;
  private readonly apiKey?: string;
  private readonly timeoutMs: number;

  constructor(config: ChatConfig = {}) {
    this.endpoint =
      config.endpoint ??
      (typeof window !== "undefined" ? (window as unknown as Record<string, string>).__BYTEROVER_CHAT_ENDPOINT__ : undefined);
    this.apiKey =
      config.apiKey ??
      (typeof window !== "undefined" ? (window as unknown as Record<string, string>).__BYTEROVER_API_KEY__ : undefined);
    this.timeoutMs = config.timeoutMs ?? DEFAULT_TIMEOUT;
  }

  async generateReply(history: ChatMessage[], userMessage: string): Promise<AssistantReply> {
    if (!this.endpoint) {
      return {
        role: "assistant",
        content:
          "Zatím nemám přístup k LLM. Zkus formulovat požadavek pomocí podporovaných intentů (DPH, formát CZK, kurz ČNB, deduplikace, svátky, SLA)."
      };
    }

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);

    try {
      const response = await fetch(this.endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...(this.apiKey ? { Authorization: `Bearer ${this.apiKey}` } : {})
        },
        body: JSON.stringify({
          history,
          message: userMessage
        }),
        signal: controller.signal
      });

      if (!response.ok) {
        throw new Error(`Chat backend returned ${response.status}`);
      }

      const payload = await response.json();
      const content =
        typeof payload?.message === "string"
          ? payload.message
          : typeof payload?.reply === "string"
          ? payload.reply
          : "Chat backend nevrátil odpověď.";

      return {
        role: "assistant",
        content,
        metadata: typeof payload === "object" && payload ? { raw: payload } : undefined
      };
    } catch (error) {
      const message =
        error instanceof Error && error.name === "AbortError"
          ? "Chat backend neodpověděl v časovém limitu."
          : error instanceof Error
          ? error.message
          : "Chyba při volání chat backendu.";

      return {
        role: "error",
        content: message
      };
    } finally {
      clearTimeout(timer);
    }
  }
}
