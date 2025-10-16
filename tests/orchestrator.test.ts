import { beforeEach, describe, expect, it, vi } from "vitest";
import { IntentType, IntentPreview } from "../src/backend/intents/types";

vi.mock("../src/backend/intents/parser", () => {
  return {
    parseCzechRequest: vi.fn()
  };
});

vi.mock("../src/backend/workbook/preview", () => {
  return {
    buildPreview: vi.fn()
  };
});

vi.mock("../src/backend/workbook/telemetry", () => {
  return {
    logTelemetryEvent: vi.fn()
  };
});

const { parseCzechRequest } = await import("../src/backend/intents/parser");
const { buildPreview } = await import("../src/backend/workbook/preview");
const { logTelemetryEvent } = await import("../src/backend/workbook/telemetry");

const deterministicPreview: IntentPreview = {
  intent: {
    type: IntentType.VatAdd,
    rate: 0.21,
    rateLabel: "21 %",
    columnLetter: "C",
    originalText: "",
    confidence: 1
  },
  issues: [],
  planText: "1. Testovací plán",
  applyPayload: {
    sheetName: "Sheet1",
    columnIndex: 0,
    rowIndex: 0,
    rowCount: 2,
    hasHeader: true,
    rate: 0.21,
    rateLabel: "21 %"
  }
};

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(parseCzechRequest).mockImplementation((text: string) => {
    if (text.includes("DPH")) {
      return {
        intent: deterministicPreview.intent,
        issues: []
      };
    }
    return null;
  });
  vi.mocked(buildPreview).mockResolvedValue(deterministicPreview);
});

describe("ChatOrchestrator", () => {
  it("returns deterministic intent preview and logs telemetry", async () => {
    const backendStub = { generateReply: vi.fn() } as unknown as import("../src/backend/chat/backend").ChatBackend;
    const { ChatOrchestrator } = await import("../src/backend/chat/orchestrator");
    const orchestrator = new ChatOrchestrator(backendStub);

    const outcome = await orchestrator.handleUserMessage("Přidej DPH 21 % do sloupce C");

    expect(outcome.kind).toBe("intent-preview");
    expect(vi.mocked(logTelemetryEvent)).toHaveBeenCalledWith(
      expect.objectContaining({ event: "preview", intent: IntentType.VatAdd, detail: "deterministic" })
    );
    expect(backendStub.generateReply).not.toHaveBeenCalled();
  });

  it("handles LLM follow-up intent and logs telemetry", async () => {
    vi.mocked(parseCzechRequest).mockImplementation((text: string) => {
      if (text.includes("follow")) {
        return {
          intent: deterministicPreview.intent,
          issues: []
        };
      }
      return null;
    });

    const backendStub = {
      generateReply: vi.fn(async () => ({
        role: "assistant" as const,
        content: "Navrhuji připravit plán",
        metadata: { followUpIntent: "follow intent" }
      }))
    } as unknown as import("../src/backend/chat/backend").ChatBackend;

    const { ChatOrchestrator } = await import("../src/backend/chat/orchestrator");
    const orchestrator = new ChatOrchestrator(backendStub);

    const outcome = await orchestrator.handleUserMessage("Potřebuji pomoc");

    expect(outcome.kind).toBe("intent-preview");
    expect(vi.mocked(logTelemetryEvent)).toHaveBeenCalledWith(
      expect.objectContaining({ event: "preview", intent: IntentType.VatAdd, detail: "llm" })
    );
  });

  it("logs fallback when chat response has no intent", async () => {
    vi.mocked(parseCzechRequest).mockReturnValue(null);
    const backendStub = {
      generateReply: vi.fn(async () => ({
        role: "assistant" as const,
        content: "Odpověď",
        metadata: {}
      }))
    } as unknown as import("../src/backend/chat/backend").ChatBackend;

    const { ChatOrchestrator } = await import("../src/backend/chat/orchestrator");
    const orchestrator = new ChatOrchestrator(backendStub);

    const outcome = await orchestrator.handleUserMessage("Napiš něco");

    expect(outcome.kind).toBe("assistant-message");
    expect(vi.mocked(logTelemetryEvent)).toHaveBeenCalledWith(
      expect.objectContaining({ event: "chat_fallback", detail: "conversation" })
    );
  });
});
