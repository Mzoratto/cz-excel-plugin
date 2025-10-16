import { handleApplyRequest, handlePreviewRequest, handleUndoRequest } from "../core/controller";

const QUICK_ACTIONS = [
  { label: "DPH 21 %", prompt: "Přidej DPH 21 % do sloupce C" },
  { label: "Formát CZK", prompt: "Nastav formát CZK pro sloupec B" },
  { label: "Odeber duplicity", prompt: "Odeber duplicity ve sloupci C" },
  { label: "Kurz ČNB", prompt: "Kurz ČNB EUR dnes" },
  { label: "Svátky", prompt: "Vlož svátky 2026" },
  { label: "SLA +5", prompt: "Spočítej termín za 5 pracovních dní od dnes" }
];

function bindQuickActions(input: HTMLTextAreaElement) {
  const chips = document.querySelectorAll<HTMLButtonElement>("[data-role='quick-chip']");
  chips.forEach((chip) => {
    chip.addEventListener("click", () => {
      input.value = chip.dataset.prompt ?? "";
      input.focus();
    });
  });
}

function bindFormHandlers(input: HTMLTextAreaElement) {
  const form = document.querySelector<HTMLFormElement>("[data-role='request-form']");
  const previewButton = document.querySelector<HTMLButtonElement>("[data-role='preview']");
  const applyButton = document.querySelector<HTMLButtonElement>("[data-role='apply']");
  const undoButton = document.querySelector<HTMLButtonElement>("[data-role='undo']");

  if (!form || !previewButton || !applyButton || !undoButton) {
    return;
  }

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    previewButton.click();
  });

  previewButton.addEventListener("click", async () => {
    const request = input.value;
    await handlePreviewRequest(request);
  });

  applyButton.addEventListener("click", async () => {
    await handleApplyRequest();
  });

  undoButton.addEventListener("click", () => {
    handleUndoRequest();
  });
}

export function renderApp() {
  const root = document.querySelector<HTMLDivElement>("#root");
  if (!root) {
    throw new Error("Missing root element.");
  }

  const quickActionsMarkup = QUICK_ACTIONS.map(
    (action) =>
      `<button class="chip" data-role="quick-chip" data-prompt="${action.prompt}">${action.label}</button>`
  ).join("");

  root.innerHTML = `
    <div class="pane">
      <header class="pane__header">
        <h1 class="pane__title">CZ Excel Copilot</h1>
        <p class="pane__subtitle">Zadej pokyn v češtině, zobraz náhled a aplikuj změny bezpečně.</p>
      </header>
      <section class="pane__section">
        <form data-role="request-form" class="pane__form">
          <label for="request-input" class="pane__label">Tvůj požadavek</label>
          <textarea id="request-input" class="pane__textarea" rows="4" placeholder="Např. Přidej DPH 21 % do sloupce C"></textarea>
          <div class="pane__chips">
            ${quickActionsMarkup}
          </div>
          <div class="pane__actions">
            <button type="button" data-role="preview" class="button button--primary">Náhled</button>
            <button type="button" data-role="apply" class="button button--accent" disabled>Provést</button>
            <button type="button" data-role="undo" class="button button--ghost">Zpět</button>
          </div>
        </form>
      </section>
      <section class="pane__section">
        <h2 class="pane__section-title">Náhled</h2>
        <div id="preview-pane" class="preview-pane">
          <p class="preview-pane__placeholder">Ve vývoji: zde se zobrazí plán a vzorová data.</p>
        </div>
      </section>
      <section class="pane__section">
        <h2 class="pane__section-title">Log</h2>
        <div id="log-pane" class="log-pane">
          <p class="log-pane__placeholder">Log akcí bude čitelný i v listu _Audit.</p>
        </div>
      </section>
    </div>
  `;

  const input = document.querySelector<HTMLTextAreaElement>("#request-input");
  if (!input) {
    throw new Error("Missing request input element.");
  }

  bindQuickActions(input);
  bindFormHandlers(input);
}
