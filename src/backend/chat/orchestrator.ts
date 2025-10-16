import { parseCzechRequest } from "../intents/parser";
import { buildPreview } from "../workbook/preview";
import { IntentPreview, ParsedIntentOutcome, isFailedPreview, FailedPreview } from "../intents/types";
import { ChatBackend } from "./backend";
import { ChatSession } from "./session";
import { ChatMessage } from "./types";

export type ChatOutcome =
  | {
      kind: "intent-preview";
      preview: IntentPreview;
      parsed: ParsedIntentOutcome;
      assistantMessage: ChatMessage;
      userMessage: ChatMessage;
    }
  | {
      kind: "assistant-message";
      assistantMessage: ChatMessage;
      userMessage: ChatMessage;
      failure?: FailedPreview;
    };

function summarizeIntent(preview: IntentPreview, parsed: ParsedIntentOutcome): string {
  const base = parsed.intent.type;
  const issuesNote = preview.issues.length > 0 ? " (pozor na doporučení níže)." : ".";
  return `Připravil jsem plán pro ${base}${issuesNote}`;
}

export class ChatOrchestrator {
  private readonly session: ChatSession;

  constructor(private readonly backend: ChatBackend, session?: ChatSession) {
    this.session = session ?? new ChatSession();
  }

  getHistory(): ChatMessage[] {
    return this.session.getMessages();
  }

  async handleUserMessage(text: string): Promise<ChatOutcome> {
    const trimmed = text.trim();
    if (!trimmed) {
      const userMessage = this.session.addMessage("user", "");
      const assistant = this.session.addMessage("system", "Zadej prosím požadavek, abych mohl pokračovat.");
      return { kind: "assistant-message", assistantMessage: assistant, userMessage };
    }

    const userMessage = this.session.addMessage("user", trimmed);

    const parsed = parseCzechRequest(trimmed);
    if (parsed) {
      const preview = await buildPreview(parsed.intent);
      if (isFailedPreview(preview)) {
        const assistant = this.session.addMessage(
          "error",
          `Nepodařilo se připravit plán: ${preview.error}`
        );
        return { kind: "assistant-message", assistantMessage: assistant, userMessage, failure: preview };
      }
      const assistant = this.session.addMessage("action", summarizeIntent(preview, parsed));
      return {
        kind: "intent-preview",
        preview,
        parsed,
        assistantMessage: assistant,
        userMessage
      };
    }

    const reply = await this.backend.generateReply(this.session.getMessages(), trimmed);
    const assistant = this.session.addMessage(reply.role, reply.content, reply.metadata);
    return {
      kind: "assistant-message",
      assistantMessage: assistant,
      userMessage
    };
  }
}
