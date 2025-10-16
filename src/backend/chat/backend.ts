import { ChatMessage, AssistantReply, ChatConfig } from "./types";

const DEFAULT_TIMEOUT = 15000;

const SYSTEM_PROMPT = `Jsi specializovaný český Excel Copilot.
- Odpovídej stručně česky (max. 2 věty), bez příkazu k provedení.
- Pokud je vhodné spustit deterministickou akci, vrať JSON bez formátování (bez code blocku):
  {
    "reply": "stručné lidské vysvětlení",
    "follow_up_intent": "věta v češtině, kterou agent dokáže zpracovat (např. 'Přidej DPH 21 % do sloupce C')",
    "notes": "volitelně doplňující informace"
  }
- Pokud akce není k dispozici, vrať buď text nebo JSON s polem "reply" a krátce vysvětli proč.
- Nepiš nic jiného než odpověď nebo JSON; žádné systémové komentáře.`;

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
      const messages = [
        { role: "system", content: SYSTEM_PROMPT },
        ...history.map((message) => ({
          role: mapRoleForLLM(message.role),
          content: message.content
        }))
      ];

      const response = await fetch(this.endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...(this.apiKey ? { Authorization: `Bearer ${this.apiKey}` } : {})
        },
        body: JSON.stringify({
          messages,
          input: userMessage
        }),
        signal: controller.signal
      });

      if (!response.ok) {
        throw new Error(`Chat backend returned ${response.status}`);
      }

      let payload: unknown;
      try {
        payload = await response.json();
      } catch (error) {
        console.warn("Chat backend nevrátil JSON", error);
      }

      const { content, metadata } = extractReply(payload, userMessage);

      return {
        role: "assistant",
        content,
        metadata
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

function mapRoleForLLM(role: ChatMessage["role"]): "user" | "assistant" | "system" {
  if (role === "user") {
    return "user";
  }
  return "assistant";
}

function extractReply(payload: unknown, fallback: string): { content: string; metadata?: Record<string, unknown> } {
  if (!payload || typeof payload !== "object") {
    return { content: typeof payload === "string" ? payload : fallback };
  }

  const data = payload as Record<string, unknown>;
  const contentCandidate =
    typeof data.reply === "string"
      ? data.reply
      : typeof data.message === "string"
      ? data.message
      : typeof data.response === "string"
      ? data.response
      : undefined;

  const metadata: Record<string, unknown> = { raw: payload };

  if (typeof data.follow_up_intent === "string" && data.follow_up_intent.trim().length > 0) {
    metadata.followUpIntent = data.follow_up_intent.trim();
  }

  if (typeof data.notes === "string") {
    metadata.notes = data.notes;
  }

  return { content: contentCandidate ?? fallback, metadata };
}
