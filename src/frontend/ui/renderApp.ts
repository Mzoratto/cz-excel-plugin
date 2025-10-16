import { handleApplyRequest, handleUndoRequest, handleUserRequest } from "../core/controller";
import { clearPlan } from "./state";

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
  const sendButton = document.querySelector<HTMLButtonElement>("[data-role='send']");
  const applyButton = document.querySelector<HTMLButtonElement>("[data-role='apply']");
  const undoButton = document.querySelector<HTMLButtonElement>("[data-role='undo']");

  if (!form || !sendButton || !applyButton || !undoButton) {
    return;
  }

  const submitHandler = async () => {
    const request = input.value;
    input.value = "";
    await handleUserRequest(request);
    input.focus();
  };

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    submitHandler().catch((error) => {
      console.error("Chat handling failed", error);
    });
  });

  sendButton.addEventListener("click", () => {
    submitHandler().catch((error) => {
      console.error("Chat handling failed", error);
    });
  });

  applyButton.addEventListener("click", () => {
    handleApplyRequest().catch((error) => {
      console.error("Apply failed", error);
    });
  });

  undoButton.addEventListener("click", () => {
    handleUndoRequest().catch((error) => {
      console.error("Undo failed", error);
    });
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
        <p class="pane__subtitle">Chatbot, který rozumí českým finančním požadavkům v Excelu.</p>
      </header>
      <section class="pane__section pane__section--chat">
        <div id="chat-thread" class="chat-thread">
          <p class="chat-thread__placeholder">Zeptej se například: „Přidej DPH 21 % do sloupce C“.</p>
        </div>
      </section>
      <section class="pane__section">
        <form data-role="request-form" class="pane__form">
          <label for="request-input" class="pane__label">Tvůj požadavek</label>
          <textarea id="request-input" class="pane__textarea" rows="3" placeholder="Např. Přepočítej USD na CZK k 5.1.2024"></textarea>
          <div class="pane__chips">
            ${quickActionsMarkup}
          </div>
          <div class="pane__actions">
            <button type="button" data-role="send" class="button button--primary">Odeslat</button>
            <button type="button" data-role="apply" class="button button--accent" disabled>Provést plán</button>
            <button type="button" data-role="undo" class="button button--ghost">Zpět</button>
          </div>
        </form>
      </section>
      <section class="pane__section">
        <h2 class="pane__section-title">Plán</h2>
        <div id="plan-pane" class="plan-pane">
          <p class="plan-pane__placeholder">Plán se zobrazí po rozpoznání konkrétní akce.</p>
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
  clearPlan();
  input.focus();
}
