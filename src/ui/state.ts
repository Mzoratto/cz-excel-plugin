import { IntentPreview, SampleTable, FailedPreview, isFailedPreview } from "../intents/types";

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => {
    switch (char) {
      case "&":
        return "&amp;";
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case '"':
        return "&quot;";
      case "'":
        return "&#39;";
      default:
        return char;
    }
  });
}

function ensurePane(selector: string): HTMLElement {
  const element = document.querySelector<HTMLElement>(selector);
  if (!element) {
    throw new Error(`Element ${selector} nebyl nalezen.`);
  }
  return element;
}

function renderSampleTable(sample: SampleTable): string {
  const headers = sample.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("");
  const rows = sample.rows
    .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`)
    .join("");
  return `<table class="preview-table"><thead><tr>${headers}</tr></thead><tbody>${rows}</tbody></table>`;
}

function renderIssues(issues?: string[]): string {
  if (!issues || issues.length === 0) {
    return "";
  }
  const items = issues.map((issue) => `<li>${escapeHtml(issue)}</li>`).join("");
  return `<div class="alert alert--warn"><strong>Kontrola:</strong><ul>${items}</ul></div>`;
}

export function setApplyEnabled(isEnabled: boolean): void {
  const button = document.querySelector<HTMLButtonElement>("[data-role='apply']");
  if (button) {
    button.disabled = !isEnabled;
  }
}

export function showPreviewLoading(): void {
  setApplyEnabled(false);
  const pane = ensurePane("#preview-pane");
  pane.innerHTML = `<p class="preview-pane__placeholder">Připravuji náhled…</p>`;
}

export function showPreviewSuccess(preview: IntentPreview): void {
  setApplyEnabled(true);
  const pane = ensurePane("#preview-pane");
  const planHtml = `<div class="preview-plan"><pre>${escapeHtml(preview.planText)}</pre></div>`;
  const issuesHtml = renderIssues(preview.issues);
  const sampleHtml = preview.sample ? `<div class="preview-sample">${renderSampleTable(preview.sample)}</div>` : "";
  pane.innerHTML = `${planHtml}${issuesHtml}${sampleHtml}`;
}

export function showPreviewFailure(preview: FailedPreview): void {
  setApplyEnabled(false);
  const pane = ensurePane("#preview-pane");
  const issuesHtml = renderIssues(preview.issues);
  pane.innerHTML = `<div class="alert alert--error"><strong>Nelze připravit náhled.</strong><p>${escapeHtml(
    preview.error
  )}</p>${issuesHtml}</div>`;
}

export function appendLogEntry(message: string, variant: "info" | "error" = "info"): void {
  const pane = ensurePane("#log-pane");
  const placeholder = pane.querySelector(".log-pane__placeholder");
  if (placeholder) {
    placeholder.remove();
  }
  const now = new Date();
  const timestamp = now.toLocaleTimeString("cs-CZ", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  const entry = document.createElement("div");
  entry.className = `log-entry log-entry--${variant}`;
  entry.innerHTML = `<span class="log-entry__time">${escapeHtml(timestamp)}</span><span class="log-entry__message">${escapeHtml(
    message
  )}</span>`;
  pane.prepend(entry);
}

export function showUserMessage(message: string, variant: "info" | "error" = "info"): void {
  appendLogEntry(message, variant);
}

export function handlePreviewResult(result: IntentPreview | FailedPreview): void {
  if (isFailedPreview(result)) {
    showPreviewFailure(result);
  } else {
    showPreviewSuccess(result);
  }
}
