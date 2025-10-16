import { applyIntent } from "../../backend/workbook/apply";
import { performUndo } from "../../backend/workbook/undo";
import { ChatBackend } from "../../backend/chat/backend";
import { ChatOrchestrator } from "../../backend/chat/orchestrator";
import {
  appendChatMessage,
  clearPlan,
  hideTypingIndicator,
  setApplyEnabled,
  showPlan,
  showPlanFailure,
  showTypingIndicator
} from "../ui/state";
import { IntentPreview } from "../../backend/intents/types";

const orchestrator = new ChatOrchestrator(new ChatBackend());

let lastPreview: IntentPreview | null = null;

function updatePlanState(preview: IntentPreview | null): void {
  if (preview) {
    showPlan(preview);
    setApplyEnabled(true);
  } else {
    clearPlan();
    setApplyEnabled(false);
  }
}

export async function handleUserRequest(input: string): Promise<void> {
  const request = input.trim();
  if (!request) {
    return;
  }

  showTypingIndicator();
  try {
    const outcome = await orchestrator.handleUserMessage(request);
    appendChatMessage(outcome.userMessage);

    if (outcome.kind === "intent-preview") {
      lastPreview = outcome.preview;
      updatePlanState(outcome.preview);
      appendChatMessage(outcome.assistantMessage);
    } else {
      lastPreview = null;
      if (outcome.failure) {
        showPlanFailure(outcome.failure);
      } else {
        clearPlan();
      }
      setApplyEnabled(false);
      appendChatMessage(outcome.assistantMessage);
    }
  } catch (error) {
    lastPreview = null;
    setApplyEnabled(false);
    const message =
      error instanceof Error
        ? error.message
        : "Došlo k neočekávané chybě při zpracování požadavku.";
    appendChatMessage({
      id: `error-${Date.now()}`,
      role: "error",
      content: message,
      timestamp: new Date().toISOString()
    });
  } finally {
    hideTypingIndicator();
  }
}

export async function handleApplyRequest(): Promise<void> {
  if (!lastPreview) {
    appendChatMessage({
      id: `warn-${Date.now()}`,
      role: "system",
      content: "Nejdříve musí být připraven plán. Napiš mi, co mám udělat.",
      timestamp: new Date().toISOString()
    });
    setApplyEnabled(false);
    return;
  }

  try {
    const result = await applyIntent(lastPreview);
    appendChatMessage({
      id: `apply-${Date.now()}`,
      role: "assistant",
      content: result.message,
      timestamp: new Date().toISOString()
    });
    if (result.warnings) {
      for (const warning of result.warnings) {
        appendChatMessage({
          id: `warn-${Date.now()}-${Math.random()}`,
          role: "error",
          content: warning,
          timestamp: new Date().toISOString()
        });
      }
    }
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Operaci se nepodařilo dokončit.";
    appendChatMessage({
      id: `error-${Date.now()}`,
      role: "error",
      content: message,
      timestamp: new Date().toISOString()
    });
  } finally {
    lastPreview = null;
    updatePlanState(null);
  }
}

export async function handleUndoRequest(): Promise<void> {
  const result = await performUndo();
  appendChatMessage({
    id: `undo-${Date.now()}`,
    role: result.success ? "assistant" : "error",
    content: result.message,
    timestamp: new Date().toISOString()
  });
}
