import { parseCzechRequest } from "../intents/parser";
import { buildPreview } from "../workbook/preview";
import { IntentPreview, ParsedIntentOutcome, isFailedPreview, FailedPreview } from "../intents/types";
import { ChatBackend } from "./backend";
import { ChatSession } from "./session";
import { ChatMessage } from "./types";
import { logTelemetryEvent } from "../workbook/telemetry";

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
      await logTelemetryEvent({ event: "preview", intent: parsed.intent.type, detail: "deterministic" });
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
    const followUpIntentText =
      typeof reply.metadata?.followUpIntent === "string" ? reply.metadata.followUpIntent : undefined;

    if (followUpIntentText) {
      const synthetic = parseCzechRequest(followUpIntentText);
      if (synthetic) {
        const preview = await buildPreview(synthetic.intent);
        if (!isFailedPreview(preview)) {
          const summary = summarizeIntent(preview, synthetic);
          const content = reply.content ? `${reply.content}\n\n${summary}` : summary;
          const assistant = this.session.addMessage("assistant", content, reply.metadata);
          await logTelemetryEvent({ event: "preview", intent: synthetic.intent.type, detail: "llm" });
          return {
            kind: "intent-preview",
            preview,
            parsed: synthetic,
            assistantMessage: assistant,
            userMessage
          };
        }

        const failureAssistant = this.session.addMessage(
          "error",
          `Navrženou akci se nepodařilo připravit: ${preview.error}`
        );
        await logTelemetryEvent({ event: "chat_fallback", detail: "llm_intent_failed" });
        return {
          kind: "assistant-message",
          assistantMessage: failureAssistant,
          userMessage,
          failure: preview
        };
      }
    }

    const assistant = this.session.addMessage(reply.role, reply.content, reply.metadata);
    await logTelemetryEvent({ event: "chat_fallback", detail: reply.metadata?.followUpIntent ? "intent_parse_failed" : "conversation" });
    return {
      kind: "assistant-message",
      assistantMessage: assistant,
      userMessage
    };
  }
}
