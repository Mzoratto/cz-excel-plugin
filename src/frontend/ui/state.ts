import { IntentPreview, FailedPreview, SampleTable } from "../../backend/intents/types";
import { ChatMessage } from "../../backend/chat/types";

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

function ensureElement<T extends HTMLElement>(selector: string): T {
  const element = document.querySelector<T>(selector);
  if (!element) {
    throw new Error(`Element ${selector} nebyl nalezen.`);
  }
  return element;
}

function createMessageElement(message: ChatMessage): HTMLElement {
  const wrapper = document.createElement("div");
  wrapper.className = `chat-message chat-message--${message.role}`;
  wrapper.innerHTML = `
    <div class="chat-message__content">${escapeHtml(message.content)}</div>
    <div class="chat-message__meta">${new Date(message.timestamp).toLocaleTimeString("cs-CZ", {
      hour: "2-digit",
      minute: "2-digit"
    })}</div>
  `;
  return wrapper;
}

function renderSampleTable(sample: SampleTable): string {
  const headers = sample.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("");
  const rows = sample.rows
    .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`)
    .join("");
  return `<table class="preview-table"><thead><tr>${headers}</tr></thead><tbody>${rows}</tbody></table>`;
}

function renderIssues(issues: string[]): string {
  if (!issues || issues.length === 0) {
    return "";
  }
  const items = issues.map((issue) => `<li>${escapeHtml(issue)}</li>`).join("");
  return `<div class="alert alert--warn"><strong>Doporučení:</strong><ul>${items}</ul></div>`;
}

export function appendChatMessage(message: ChatMessage): void {
  const thread = ensureElement<HTMLDivElement>("#chat-thread");
  const placeholder = thread.querySelector(".chat-thread__placeholder");
  if (placeholder) {
    placeholder.remove();
  }
  thread.appendChild(createMessageElement(message));
  thread.scrollTo({ top: thread.scrollHeight, behavior: "smooth" });
}

export function clearPlan(): void {
  const pane = ensureElement<HTMLDivElement>("#plan-pane");
  pane.innerHTML = `<p class="plan-pane__placeholder">Plán se zobrazí po rozpoznání konkrétní akce.</p>`;
}

export function showPlan(preview: IntentPreview): void {
  const pane = ensureElement<HTMLDivElement>("#plan-pane");
  const planHtml = `<div class="preview-plan"><pre>${escapeHtml(preview.planText)}</pre></div>`;
  const issuesHtml = renderIssues(preview.issues);
  const sampleHtml = preview.sample ? `<div class="preview-sample">${renderSampleTable(preview.sample)}</div>` : "";
  pane.innerHTML = `${planHtml}${issuesHtml}${sampleHtml}`;
}

export function showPlanFailure(preview: FailedPreview): void {
  const pane = ensureElement<HTMLDivElement>("#plan-pane");
  const issuesHtml = preview.issues ? renderIssues(preview.issues) : "";
  pane.innerHTML = `<div class="alert alert--error"><strong>Plán se nepodařilo připravit.</strong><p>${escapeHtml(
    preview.error
  )}</p>${issuesHtml}</div>`;
}

export function setApplyEnabled(isEnabled: boolean): void {
  const button = document.querySelector<HTMLButtonElement>("[data-role='apply']");
  if (button) {
    button.disabled = !isEnabled;
  }
}
