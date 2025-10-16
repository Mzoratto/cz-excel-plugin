# CZ Excel Copilot Add-in

Offline-first český chatbot/copilot panel pro Excel. Add-in čte přirozené požadavky (DPH, kurz ČNB, deduplikace, pracovní termíny), zobrazí plán v postranním panelu a po potvrzení provede změny v sešitu včetně undo/audit vrstvy.

## Struktura projektu

- `manifest.xml` – manifest pro sideload add-inu (`https://localhost:5173/taskpane.html`).
- `taskpane.html` – HTML vstup Vite dev serveru.
- `src/frontend` – Office.js bootstrap (`main.ts`), chat UI, ovládání tlačítek a stylování.
- `src/backend` – deterministické intent služby, Excel workbook operace, chat orchestrátor a fallbacky na MCP/LLM.
- `src/backend/workbook/artifacts.ts` – zajišťuje vytvoření listů `_Audit`, `_UndoIndex`, `_UndoData`, `_FX_CNB`, `_HOLIDAYS_CZ`, `_Settings`.
- `src/backend/workbook/cnb.ts` – cache ČNB kurzů s fallbackem na API.
- `src/backend/workbook/holidays.ts` – generování českých svátků + výpočty pracovních dní.
- `docs/` – instrukce pro různé agenty (Codex, Claude, atd.).
- `public/assets` – dočasné ikonky pro manifest (nahraď vlastní grafikou).

## Lokální běh

```
npm install
npm run dev
```

V Excelu (Desktop/Web) sideload manifest: `manifest.xml`. Dev server běží na `https://localhost:5173`.

## Automatizované releasy

- `npm run package` spustí produkční build a zabalí `manifest.xml` + `dist/` do `release/cz-excel-copilot-<timestamp>.zip` – balíček připravený k nasazení nebo k nahrání do admin konzole.
- CI workflow `ci.yml` běží na každý push/PR a ukládá artefakt `cz-excel-copilot.zip` s aktuálním balíčkem.
- Push tagu ve formátu `v*` (nebo ruční spuštění `Release` workflow) vyrobí produkční build, vytvoří GitHub Release a připojí ZIP. Release notes jsou generovány automaticky.
- V Excelu lze ZIP rozbalit a `manifest.xml` sideloadovat; složka `dist/` obsahuje hotová statická aktiva pro produkci.
- `npm run manifest:build -- --host=https://cdn.example.com --output=release/manifest-prod.xml` vygeneruje manifest z šablony `manifest.template.xml` s URL pro produkční hostitele.

## Chat backend & Byterover MCP

- Add-in nejprve zkouší deterministické intent-parsování; pokud výraz nelze rozpoznat, odešle historii do chat backendu (např. Byterover MCP).
- Endpoint a token lze poskytnout přes globální proměnné `window.__BYTEROVER_CHAT_ENDPOINT__` a `window.__BYTEROVER_API_KEY__` (např. injektované skriptem nebo nastavené v host aplikaci).
- Backend posílá LLM systémovou instrukci; pokud vrátíš JSON s `reply` + `follow_up_intent`, parser automaticky připraví plán a zobrazí ho v sekci **Plán**.
- Knihovna udržuje konverzační historii na straně klienta (`ChatSession`) a do UI zapisuje chat transcript.
- Pokud je CLI/bytterover tooling k dispozici lokálně, používejte `byterover-retrieve-knowledge` a `byterover-store-knowledge` dle instrukcí v `docs/` pro průběžné učení asistenta.

## Telemetrie & logování

- `_Telemetry` tabulka sbírá anonymní eventy (preview/apply, fallback do chatu). Záznam obsahuje ISO čas, název eventu, intent a detail.
- Fallback do chatu (`chat_fallback`) se zapisuje i při odpovědi LLM bez deterministické akce, takže lze sledovat, kdy uživatelé kladou otázky mimo podporované scénáře.

## Další implementace (navazuje na PRD)

1. **Parser + intents (VAT & CZK formátování)**
   - ✅ Rozpoznání `vat.add` a `format.currency` z českých frází (`Přidej DPH 21 % …`, `Nastav formát CZK …`).
   - ✅ Náhled generuje plán + ukázkové hodnoty a hlásí problémy s výběrem.
   - ✅ Apply používá stejný plán (DPH zapisuje vedlejší sloupec, formát nastavuje měnu CZK).
   - ✅ Nově: `vat.remove` (výpočet základu bez DPH), `sheet.sort_column` (řazení vzestupně/sestupně) a `finance.dedupe`.
   - ✅ Přidané `sheet.highlight_negative` (podmíněné formátování záporných hodnot), `sheet.sum_column` (součet vybraného sloupce) a `analysis.monthly_runrate` (run-rate z posledních N měsíců).
2. **Undo & Audit vrstvy**
   - ✅ In-memory zásobník + perzistentní snapshoty (do 2 000 buněk) v `_UndoIndex` / `_UndoData`.
   - ✅ `Zpět` tlačítko vrací poslední akci; chyby a velké operace logují varování.
   - ✅ Audit zapisuje ISO čas, intent, args, rozsah a poznámku po každém apply.
3. **ČNB a kalendář**
   - ✅ `_FX_CNB`: cache + fetch `api.cnb.cz` s fallbackem a auditní stopou.
   - ✅ `seed_holidays(year)` pro `_HOLIDAYS_CZ` včetně Velkého pátku / Velikonočního pondělí.
   - ✅ Výpočet termínu `networkdays_due` využívající svátky + víkendy.
4. **Bezpečné provádění**
   - Transakce: před provedením snímek pro undo, poté apply.
   - Preview → Apply musí používat stejný hash/plán.
   - Hlásiť chyby v češtině s návrhem další akce.

## Poznámky

- Všechny budoucí cloudové volání drž na přepínači (offline default).
- Udržuj `ctx.sync()` na minimum: pracuj s obdélníkovými rozsahy, batchej změny.
- Maskuj zjevné PII v `_Audit` a logu.
