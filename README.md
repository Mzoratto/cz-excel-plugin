# CZ Excel Copilot Add-in

Offline-first český chatbot/copilot panel pro Excel. Add-in čte přirozené požadavky (DPH, kurz ČNB, deduplikace, pracovní termíny), zobrazí plán v postranním panelu a po potvrzení provede změny v sešitu včetně undo/audit vrstvy.

## Struktura projektu

- `manifest.xml` – manifest pro sideload add-inu (`https://localhost:5173/taskpane.html`).
- `taskpane.html` – HTML vstup Vite dev serveru.
- `src/main.ts` – bootstrap Office.js + UI + provisioning skrytých listů.
- `src/ui/*` – jednoduché UI v TypeScriptu a CSS.
- `src/workbook/artifacts.ts` – zajišťuje vytvoření listů `_Audit`, `_UndoIndex`, `_UndoData`, `_FX_CNB`, `_HOLIDAYS_CZ`, `_Settings`.
- `src/workbook/cnb.ts` – cache ČNB kurzů s fallbackem na API.
- `src/workbook/holidays.ts` – generování českých svátků + výpočty pracovních dní.
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

## Další implementace (navazuje na PRD)

1. **Parser + intents (VAT & CZK formátování)**
   - ✅ Rozpoznání `vat.add` a `format.currency` z českých frází (`Přidej DPH 21 % …`, `Nastav formát CZK …`).
   - ✅ Náhled generuje plán + ukázkové hodnoty a hlásí problémy s výběrem.
   - ✅ Apply používá stejný plán (DPH zapisuje vedlejší sloupec, formát nastavuje měnu CZK).
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
