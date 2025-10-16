import { parseCzechRequest } from "../intents/parser";
import { buildPreview } from "../workbook/preview";
import { applyIntent } from "../workbook/apply";
import {
  appendLogEntry,
  handlePreviewResult,
  setApplyEnabled,
  showPreviewLoading,
  showUserMessage
} from "../ui/state";
import { IntentPreview, isFailedPreview } from "../intents/types";
import { performUndo } from "../workbook/undo";

let lastPreview: IntentPreview | null = null;

export async function handlePreviewRequest(request: string): Promise<void> {
  const trimmed = request.trim();
  if (!trimmed) {
    handlePreviewResult({ error: "Zadej větu s požadavkem, např. „Přidej DPH 21 % do sloupce C“." });
    return;
  }

  showPreviewLoading();
  try {
    const parsed = parseCzechRequest(trimmed);
    if (!parsed) {
      handlePreviewResult({
        error: "Požadavek se nepodařilo rozpoznat.",
        issues: ["Zkus použít klíčová slova např. „DPH 21 %“ nebo „Formát CZK“."]
      });
      appendLogEntry("Nepodařilo se rozpoznat zadaný požadavek.", "error");
      lastPreview = null;
      return;
    }

    const preview = await buildPreview(parsed.intent);
    handlePreviewResult(preview);

    if (isFailedPreview(preview)) {
      appendLogEntry(preview.error, "error");
      lastPreview = null;
    } else {
      appendLogEntry(`Náhled připraven: ${parsed.intent.type}`, "info");
      lastPreview = preview;
    }
  } catch (error) {
    console.error("Preview failed", error);
    handlePreviewResult({
      error: "Během přípravy náhledu došlo k chybě.",
      issues: ["Zkontroluj, že máš vybraný správný rozsah a zkus to znovu."]
    });
    appendLogEntry("Chyba při vytváření náhledu.", "error");
    lastPreview = null;
  }
}

export async function handleApplyRequest(): Promise<void> {
  if (!lastPreview) {
    showUserMessage("Nejprve spusť náhled, aby byl plán potvrzen.", "error");
    setApplyEnabled(false);
    return;
  }

  try {
    const result = await applyIntent(lastPreview);
    showUserMessage(result.message, "info");
    if (result.warnings) {
      for (const warning of result.warnings) {
        showUserMessage(warning, "error");
      }
    }
    showUserMessage("Akce dokončena. Pro další krok spusť nový náhled.");
    lastPreview = null;
    setApplyEnabled(false);
  } catch (error) {
    console.error("Apply failed", error);
    showUserMessage("Operaci se nepodařilo dokončit.", "error");
    showUserMessage("Akci se nepodařilo provést. Zkus náhled znovu.", "error");
  }
}

export async function handleUndoRequest(): Promise<void> {
  const result = await performUndo();
  if (result.success) {
    showUserMessage(result.message, "info");
  } else {
    showUserMessage(result.message, "error");
  }
}
