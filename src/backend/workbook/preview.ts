import {
  ApplyPayload,
  FormatCzkApplyPayload,
  IntentType,
  PreviewResult,
  SampleTable,
  VatAddApplyPayload,
  VatAddIntent,
  FormatCzkIntent,
  SupportedIntent,
  FetchCnbRateIntent,
  FetchCnbRateApplyPayload,
  FxConvertCnbIntent,
  FxConvertCnbApplyPayload,
  FinanceDedupeIntent,
  FinanceDedupeApplyPayload,
  SortColumnIntent,
  SortColumnApplyPayload,
  VatRemoveIntent,
  VatRemoveApplyPayload,
  HighlightNegativeIntent,
  HighlightNegativeApplyPayload,
  SumColumnIntent,
  SumColumnApplyPayload,
  MonthlyRunRateIntent,
  MonthlyRunRateApplyPayload,
  SeedHolidaysIntent,
  SeedHolidaysApplyPayload,
  NetworkdaysDueIntent,
  NetworkdaysDueApplyPayload
} from "../intents/types";
import { parseCzechNumeric, formatCzk } from "../utils/numbers";
import { buildPlanList, columnLetterFromIndex, columnIndexFromLetter } from "../utils/excel";
import { getCachedRate } from "./cnb";
import { listCzechHolidays, loadHolidaySet, calculateBusinessDueDate } from "./holidays";
import { formatISODate } from "../utils/date";

interface SelectionSnapshot {
  sheetName: string;
  columnIndex: number;
  rowIndex: number;
  rowCount: number;
  columnCount: number;
  sampleValues: unknown[][];
}

async function captureSelection(context: Excel.RequestContext, sampleRowLimit = 6): Promise<
  | SelectionSnapshot
  | {
      error: string;
    }
> {
  const range = context.workbook.getSelectedRange();
  range.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
  const worksheet = range.worksheet;
  worksheet.load("name");
  await context.sync();

  if (range.rowCount === 0 || range.columnCount === 0) {
    return { error: "Není vybrán žádný rozsah." };
  }

  const sampleRowCount = Math.min(sampleRowLimit, range.rowCount);
  const sampleColumnCount = Math.min(range.columnCount, 3);
  const sampleRange = range.getCell(0, 0).getResizedRange(sampleRowCount - 1, sampleColumnCount - 1);
  sampleRange.load("values");
  await context.sync();

  return {
    sheetName: worksheet.name,
    columnIndex: range.columnIndex,
    rowIndex: range.rowIndex,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    sampleValues: sampleRange.values
  };
}

function ensureSingleColumn(selection: SelectionSnapshot): string | undefined {
  if (selection.columnCount !== 1) {
    return "Vyber přesně jeden sloupec včetně hlavičky.";
  }
  return undefined;
}

function validateColumnMatch(
  intendedColumn: string | undefined,
  selection: SelectionSnapshot
): string | undefined {
  if (!intendedColumn) {
    return undefined;
  }

  const selectedLetter = columnLetterFromIndex(selection.columnIndex);
  if (intendedColumn.toUpperCase() !== selectedLetter) {
    return `Během náhledu je aktivní sloupec ${selectedLetter}, ale požadavek odkazuje na sloupec ${intendedColumn}.`;
  }

  return undefined;
}

function detectHeader(selection: SelectionSnapshot): boolean {
  if (selection.rowCount <= 1) {
    return false;
  }

  const firstCell = selection.sampleValues[0]?.[0];
  const secondCell = selection.sampleValues[1]?.[0];

  const firstNumeric = parseCzechNumeric(firstCell);
  const secondNumeric = parseCzechNumeric(secondCell);

  if (firstNumeric === null && secondNumeric !== null) {
    return true;
  }

  return false;
}

function buildVatSample(intent: VatAddIntent, selection: SelectionSnapshot, hasHeader: boolean): SampleTable {
  const rows: string[][] = [];
  const dataRows = selection.sampleValues.slice(hasHeader ? 1 : 0);
  const sampleCount = Math.min(5, dataRows.length);

  for (let i = 0; i < sampleCount; i += 1) {
    const value = parseCzechNumeric(dataRows[i]?.[0]);
    if (value === null) {
      continue;
    }
    const vatValue = value * intent.rate;
    const total = value + vatValue;
    rows.push([formatCzk(value), formatCzk(vatValue), formatCzk(total)]);
  }

  if (rows.length === 0) {
    rows.push(["(žádná numerická data)", "-", "-"]);
  }

  return {
    headers: ["Základ", `DPH ${intent.rateLabel}`, "S DPH"],
    rows
  };
}

function buildSumSample(selection: SelectionSnapshot, hasHeader: boolean): SampleTable {
  const dataRows = selection.sampleValues.slice(hasHeader ? 1 : 0);
  const numeric = dataRows
    .map((row) => parseCzechNumeric(row?.[0]))
    .filter((value): value is number => value !== null);
  const previewSum = numeric.slice(0, 5).reduce((acc, value) => acc + value, 0);

  return {
    headers: ["Náhled součtu"],
    rows: [[numeric.length > 0 ? formatCzk(previewSum) : "(není co sčítat)"]]
  };
}

function buildVatRemoveSample(intent: VatRemoveIntent, selection: SelectionSnapshot, hasHeader: boolean): SampleTable {
  const rows: string[][] = [];
  const dataRows = selection.sampleValues.slice(hasHeader ? 1 : 0);
  const sampleCount = Math.min(5, dataRows.length);

  for (let i = 0; i < sampleCount; i += 1) {
    const value = parseCzechNumeric(dataRows[i]?.[0]);
    if (value === null) {
      rows.push(["(nenumerické)", "-", "-"]);
      continue;
    }
    const base = value / (1 + intent.rate);
    const vatAmount = value - base;
    rows.push([formatCzk(value), formatCzk(base), formatCzk(vatAmount)]);
  }

  if (rows.length === 0) {
    rows.push(["(žádná numerická data)", "-", "-"]);
  }

  return {
    headers: ["S DPH", "Bez DPH", `DPH ${intent.rateLabel}`],
    rows
  };
}

function buildFormatSample(selection: SelectionSnapshot, hasHeader: boolean): SampleTable {
  const rows: string[][] = [];
  const dataRows = selection.sampleValues.slice(hasHeader ? 1 : 0);
  const sampleCount = Math.min(5, dataRows.length);

  for (let i = 0; i < sampleCount; i += 1) {
    const raw = dataRows[i]?.[0];
    const numeric = parseCzechNumeric(raw);
    rows.push([
      typeof raw === "string" ? raw : `${raw}`,
      numeric !== null ? formatCzk(numeric) : "(nenumerické)"
    ]);
  }

  if (rows.length === 0) {
    rows.push(["(prázdný řádek)", "-"]);
  }

  return {
    headers: ["Původní hodnota", "Formát CZK"],
    rows
  };
}

function formatDedupeValue(value: unknown): string {
  if (value === null || value === undefined || value === "") {
    return "(prázdné)";
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return value.toString();
  }
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  return `${value}`;
}

function buildDedupeSample(selection: SelectionSnapshot, hasHeader: boolean): {
  sample: SampleTable;
  duplicateCount: number;
} {
  const dataRows = selection.sampleValues.slice(hasHeader ? 1 : 0);
  const seen = new Map<string, number>();
  const duplicates: Array<{ firstRow: number; duplicateRow: number; values: unknown[] }> = [];

  dataRows.forEach((row, index) => {
    const normalizedRow = JSON.stringify(row ?? []);
    if (seen.has(normalizedRow)) {
      duplicates.push({
        firstRow: seen.get(normalizedRow)!,
        duplicateRow: index,
        values: Array.isArray(row) ? row : []
      });
    } else {
      seen.set(normalizedRow, index);
    }
  });

  const rows: string[][] = [];
  const displayLimit = Math.min(5, duplicates.length);
  const rowOffset = hasHeader ? 2 : 1;

  for (let i = 0; i < displayLimit; i += 1) {
    const item = duplicates[i]!;
    const first = item.firstRow + rowOffset;
    const duplicate = item.duplicateRow + rowOffset;
    const valuePreview = item.values.map((cell) => formatDedupeValue(cell)).join(" | ") || "(prázdné)";
    rows.push([`${first} ↔ ${duplicate}`, valuePreview]);
  }

  if (rows.length === 0) {
    rows.push(["-", "Ve vzorku nebyly nalezeny duplicitní řádky."]);
  }

  return {
    sample: {
      headers: ["Řádky", "Hodnoty"],
      rows
    },
    duplicateCount: duplicates.length
  };
}

async function buildVatPreview(intent: VatAddIntent): Promise<PreviewResult<VatAddApplyPayload>> {
  return Excel.run(async (context) => {
    const selection = await captureSelection(context);
    if ("error" in selection) {
      return { error: selection.error };
    }

    const issues: string[] = [];
    const singleColumnIssue = ensureSingleColumn(selection);
    if (singleColumnIssue) {
      issues.push(singleColumnIssue);
    }

    const columnConflict = validateColumnMatch(intent.columnLetter, selection);
    if (columnConflict) {
      issues.push(columnConflict);
    }

    if (selection.rowCount <= 1) {
      issues.push("Rozsah musí obsahovat alespoň jeden řádek s daty pod hlavičkou.");
    }

    if (issues.length > 0) {
      return { error: "Nelze připravit náhled pro DPH.", issues };
    }

    const hasHeader = detectHeader(selection);
    const selectedLetter = columnLetterFromIndex(selection.columnIndex);
    const targetLetter = columnLetterFromIndex(selection.columnIndex + 1);

    const sample = buildVatSample(intent, selection, hasHeader);

    const planItems = [
      `Vypočítat DPH ${intent.rateLabel} pro hodnoty ve sloupci ${selectedLetter}.`,
      `Vyplnit sloupec ${targetLetter} výsledky (jen data, hlavička "${hasHeader ? "DPH " + intent.rateLabel : "DPH"}").`,
      "Nastavit formát měny CZK pro nový sloupec."
    ];

    const applyPayload: VatAddApplyPayload = {
      sheetName: selection.sheetName,
      columnIndex: selection.columnIndex,
      rowIndex: selection.rowIndex,
      rowCount: selection.rowCount,
      hasHeader,
      rate: intent.rate,
      rateLabel: intent.rateLabel
    };

    return {
      intent,
      issues: [],
      planText: buildPlanList(planItems),
      sample,
      applyPayload
    };
  });
}

async function buildFormatPreview(
  intent: FormatCzkIntent
): Promise<PreviewResult<FormatCzkApplyPayload>> {
  return Excel.run(async (context) => {
    const selection = await captureSelection(context);
    if ("error" in selection) {
      return { error: selection.error };
    }

    const issues: string[] = [];

    const singleColumnIssue = ensureSingleColumn(selection);
    if (singleColumnIssue) {
      issues.push(singleColumnIssue);
    }

    const columnConflict = validateColumnMatch(intent.columnLetter, selection);
    if (columnConflict) {
      issues.push(columnConflict);
    }

    const hasHeader = detectHeader(selection);
    const selectedLetter = columnLetterFromIndex(selection.columnIndex);
    const sample = buildFormatSample(selection, hasHeader);

    if (issues.length > 0) {
      return { error: "Nelze připravit náhled pro formátování CZK.", issues };
    }

    const planItems = [
      `Nastavit formát CZK pro vybraný sloupec ${selectedLetter}.`,
      "Zachovat původní hodnoty buněk, změnit pouze číslený formát."
    ];

    const applyPayload: FormatCzkApplyPayload = {
      sheetName: selection.sheetName,
      columnIndex: selection.columnIndex,
      rowIndex: selection.rowIndex,
      rowCount: selection.rowCount
    };

    return {
      intent,
      issues: [],
      planText: buildPlanList(planItems),
      sample,
      applyPayload
    };
  });
}

async function buildDedupePreview(
  intent: FinanceDedupeIntent
): Promise<PreviewResult<FinanceDedupeApplyPayload>> {
  return Excel.run(async (context) => {
    const selection = await captureSelection(context);
    if ("error" in selection) {
      return { error: selection.error };
    }

    const blockingIssues: string[] = [];

    const singleColumnIssue = ensureSingleColumn(selection);
    if (singleColumnIssue) {
      blockingIssues.push(singleColumnIssue);
    }

    const columnConflict = validateColumnMatch(intent.columnLetter, selection);
    if (columnConflict) {
      blockingIssues.push(columnConflict);
    }

    if (selection.rowCount <= 1) {
      blockingIssues.push("Rozsah musí obsahovat alespoň dva řádky.");
    }

    if (blockingIssues.length > 0) {
      return { error: "Nelze připravit náhled pro odebrání duplicit.", issues: blockingIssues };
    }

    const hasHeader = detectHeader(selection);
    const { sample, duplicateCount } = buildDedupeSample(selection, hasHeader);

    const startLetter = columnLetterFromIndex(selection.columnIndex);
    const endLetter = columnLetterFromIndex(selection.columnIndex + selection.columnCount - 1);
    const startRow = selection.rowIndex + 1;
    const endRow = selection.rowIndex + selection.rowCount;
    const rangeLabel =
      selection.columnCount === 1
        ? `${startLetter}${startRow}:${startLetter}${endRow}`
        : `${startLetter}${startRow}:${endLetter}${endRow}`;

    const planItems = [
      `Analyzovat rozsah ${rangeLabel} a identifikovat duplicitní hodnoty ${hasHeader ? "(bez hlavičky)" : ""}.`,
      "Odebrat duplicitní řádky a ponechat první výskyt každé hodnoty.",
      "Zapsat výsledek a zaznamenat akci do auditu."
    ];

    const informationalIssues =
      duplicateCount === 0
        ? ["Ve vzorku nebyly nalezeny duplicitní řádky. Operace proběhne pro jistotu na celém rozsahu."]
        : [];

    const applyPayload: FinanceDedupeApplyPayload = {
      sheetName: selection.sheetName,
      rowIndex: selection.rowIndex,
      columnIndex: selection.columnIndex,
      rowCount: selection.rowCount,
      columnCount: selection.columnCount,
      hasHeader
    };

    return {
      intent,
      issues: informationalIssues,
      planText: buildPlanList(planItems),
      sample,
      applyPayload
    };
  });
}

async function buildSortPreview(intent: SortColumnIntent): Promise<PreviewResult<SortColumnApplyPayload>> {
  return Excel.run(async (context) => {
    const selection = await captureSelection(context);
    if ("error" in selection) {
      return { error: selection.error };
    }

    const issues: string[] = [];
    const singleColumnIssue = ensureSingleColumn(selection);
    if (singleColumnIssue) {
      issues.push(singleColumnIssue);
    }

    const columnConflict = validateColumnMatch(intent.columnLetter, selection);
    if (columnConflict) {
      issues.push(columnConflict);
    }

    if (selection.rowCount <= 1) {
      issues.push("Rozsah musí obsahovat alespoň dva řádky.");
    }

    if (issues.length > 0) {
      return { error: "Nelze připravit náhled pro seřazení.", issues };
    }

    const hasHeader = detectHeader(selection);
    const selectedLetter = columnLetterFromIndex(selection.columnIndex);
    const directionLabel = intent.direction === "asc" ? "vzestupně" : "sestupně";

    const planItems = [
      `Seřadit hodnoty ve sloupci ${selectedLetter} ${directionLabel}.`,
      hasHeader ? "Zachovat hlavičku mimo řazení." : "Řadit všechny řádky včetně prvního.",
      "Zapsat informaci o akci do auditu."
    ];

    const applyPayload: SortColumnApplyPayload = {
      sheetName: selection.sheetName,
      rowIndex: selection.rowIndex,
      columnIndex: selection.columnIndex,
      rowCount: selection.rowCount,
      columnCount: selection.columnCount,
      hasHeader,
      ascending: intent.direction === "asc"
    };

    return {
      intent,
      issues: [],
      planText: buildPlanList(planItems),
      sample: {
        headers: ["Poznámka"],
        rows: [[`Ukázka po seřazení se zobrazí až po provedení akce.`]]
      },
      applyPayload
    };
  });
}

async function buildVatRemovePreview(intent: VatRemoveIntent): Promise<PreviewResult<VatRemoveApplyPayload>> {
  return Excel.run(async (context) => {
    const selection = await captureSelection(context);
    if ("error" in selection) {
      return { error: selection.error };
    }

    const issues: string[] = [];
    const singleColumnIssue = ensureSingleColumn(selection);
    if (singleColumnIssue) {
      issues.push(singleColumnIssue);
    }

    const columnConflict = validateColumnMatch(intent.columnLetter, selection);
    if (columnConflict) {
      issues.push(columnConflict);
    }

    if (selection.rowCount <= 1) {
      issues.push("Rozsah musí obsahovat alespoň jeden řádek s daty pod hlavičkou.");
    }

    if (issues.length > 0) {
      return { error: "Nelze připravit náhled pro odebrání DPH.", issues };
    }

    const hasHeader = detectHeader(selection);
    const selectedLetter = columnLetterFromIndex(selection.columnIndex);
    const baseLetter = columnLetterFromIndex(selection.columnIndex + 1);
    const vatLetter = columnLetterFromIndex(selection.columnIndex + 2);

    const sample = buildVatRemoveSample(intent, selection, hasHeader);

    const planItems = [
      `Spočítat základ bez DPH ${intent.rateLabel} z hodnot ve sloupci ${selectedLetter}.`,
      `Vyplnit sloupec ${baseLetter} hodnotami bez DPH a sloupec ${vatLetter} výší DPH.`,
      "Nastavit formát měny CZK a zapsat akci do auditu."
    ];

    const applyPayload: VatRemoveApplyPayload = {
      sheetName: selection.sheetName,
      columnIndex: selection.columnIndex,
      rowIndex: selection.rowIndex,
      rowCount: selection.rowCount,
      hasHeader,
      rate: intent.rate,
      rateLabel: intent.rateLabel
    };

    return {
      intent,
      issues: [],
      planText: buildPlanList(planItems),
      sample,
      applyPayload
    };
  });
}

async function buildHighlightNegativePreview(
  intent: HighlightNegativeIntent
): Promise<PreviewResult<HighlightNegativeApplyPayload>> {
  return Excel.run(async (context) => {
    const selection = await captureSelection(context);
    if ("error" in selection) {
      return { error: selection.error };
    }

    const issues: string[] = [];
    const singleColumnIssue = ensureSingleColumn(selection);
    if (singleColumnIssue) {
      issues.push(singleColumnIssue);
    }

    const columnConflict = validateColumnMatch(intent.columnLetter, selection);
    if (columnConflict) {
      issues.push(columnConflict);
    }

    if (issues.length > 0) {
      return { error: "Nelze připravit zvýraznění záporných hodnot.", issues };
    }

    const hasHeader = detectHeader(selection);
    const selectedLetter = columnLetterFromIndex(selection.columnIndex);

    const planItems = [
      `Přidat podmíněné formátování pro záporné hodnoty ve sloupci ${selectedLetter}.`,
      "Zvýraznit buňky červeným pozadím a tmavým písmem.",
      "Zapsat akci do auditu."
    ];

    const sample: SampleTable = {
      headers: ["Očekávaný efekt"],
      rows: [["Záporné hodnoty získají červené pozadí."]]
    };

    const applyPayload: HighlightNegativeApplyPayload = {
      sheetName: selection.sheetName,
      rowIndex: selection.rowIndex,
      columnIndex: selection.columnIndex,
      rowCount: selection.rowCount,
      hasHeader
    };

    return {
      intent,
      issues: [],
      planText: buildPlanList(planItems),
      sample,
      applyPayload
    };
  });
}

async function buildSumColumnPreview(intent: SumColumnIntent): Promise<PreviewResult<SumColumnApplyPayload>> {
  return Excel.run(async (context) => {
    const selection = await captureSelection(context);
    if ("error" in selection) {
      return { error: selection.error };
    }

    const issues: string[] = [];
    const singleColumnIssue = ensureSingleColumn(selection);
    if (singleColumnIssue) {
      issues.push(singleColumnIssue);
    }

    const columnConflict = validateColumnMatch(intent.columnLetter, selection);
    if (columnConflict) {
      issues.push(columnConflict);
    }

    if (selection.rowCount <= 1) {
      issues.push("Rozsah musí obsahovat alespoň jeden řádek s daty.");
    }

    if (issues.length > 0) {
      return { error: "Nelze připravit součet sloupce.", issues };
    }

    const hasHeader = detectHeader(selection);
    const selectedLetter = columnLetterFromIndex(selection.columnIndex);
    const sample = buildSumSample(selection, hasHeader);

    const planItems = [
      `Spočítat součet všech hodnot ve sloupci ${selectedLetter}.`,
      "Výsledek zapsat do buňky pod aktuálním výběrem.",
      "Součet formátovat jako číslo a zapsat akci do auditu."
    ];

    const applyPayload: SumColumnApplyPayload = {
      sheetName: selection.sheetName,
      rowIndex: selection.rowIndex,
      columnIndex: selection.columnIndex,
      rowCount: selection.rowCount,
      hasHeader
    };

    return {
      intent,
      issues: [],
      planText: buildPlanList(planItems),
      sample,
      applyPayload
    };
  });
}

async function buildMonthlyRunRatePreview(
  intent: MonthlyRunRateIntent
): Promise<PreviewResult<MonthlyRunRateApplyPayload>> {
  return Excel.run(async (context) => {
    const selection = await captureSelection(context);
    if ("error" in selection) {
      return { error: selection.error };
    }

    const issues: string[] = [];

    if (selection.columnCount < 2) {
      issues.push("Vyber alespoň dva sloupce: datumy a částky.");
    }

    const hasHeader = detectHeader(selection);

    const dateLetter = intent.dateColumn ?? columnLetterFromIndex(selection.columnIndex);
    const amountLetter =
      intent.amountColumn ?? columnLetterFromIndex(selection.columnIndex + Math.min(1, selection.columnCount - 1));

    const dateAbsolute = columnIndexFromLetter(dateLetter ?? "");
    const amountAbsolute = columnIndexFromLetter(amountLetter ?? "");

    if (dateAbsolute === null || amountAbsolute === null) {
      issues.push("Nepodařilo se určit, které sloupce obsahují datum a částku.");
    } else {
      const dateOffset = dateAbsolute - selection.columnIndex;
      const amountOffset = amountAbsolute - selection.columnIndex;
      if (dateOffset < 0 || dateOffset >= selection.columnCount) {
        issues.push(`Sloupec s daty (${dateLetter}) není součástí výběru.`);
      }
      if (amountOffset < 0 || amountOffset >= selection.columnCount) {
        issues.push(`Sloupec s částkami (${amountLetter}) není součástí výběru.`);
      }
    }

    if (issues.length > 0) {
      return { error: "Nelze připravit run-rate náhled.", issues };
    }

    const months = intent.months ?? 3;

    const planItems = [
      `Vzít data ve sloupci ${amountLetter} a seřadit je podle měsíců z ${dateLetter}.`,
      `Spočítat průměr za posledních ${months} měsíců a annualizovat ×12.`,
      "Výsledky zapsat do listu _RunRate včetně poznámky a auditního záznamu."
    ];

    const sample: SampleTable = {
      headers: ["Výstup"],
      rows: [[`Roční run-rate na základě ${months} měsíců.`]]
    };

    const applyPayload: MonthlyRunRateApplyPayload = {
      sheetName: selection.sheetName,
      rowIndex: selection.rowIndex,
      columnIndex: selection.columnIndex,
      rowCount: selection.rowCount,
      columnCount: selection.columnCount,
      hasHeader,
      amountColumnLetter: amountLetter!,
      dateColumnLetter: dateLetter!,
      months
    };

    return {
      intent,
      issues: [],
      planText: buildPlanList(planItems),
      sample,
      applyPayload
    };
  });
}

async function buildFetchCnbPreview(
  intent: FetchCnbRateIntent
): Promise<PreviewResult<FetchCnbRateApplyPayload>> {
  return Excel.run(async (context) => {
    const cachedRate = await getCachedRate(context, intent.currency, intent.targetDate);

    const planItems = [
      `Zkontrolovat tabulku _FX_CNB pro ${intent.currency} k datu ${intent.targetDate}.`,
      "Pokud není k dispozici, stáhnout kurz z api.cnb.cz (denní kurzy).",
      "Zapsat kurz do _FX_CNB a uvést výsledek v panelu."
    ];

    const sampleRows =
      cachedRate !== null
        ? [[intent.currency, intent.targetDate, formatCzk(cachedRate)]]
        : [[intent.currency, intent.targetDate, "— (zatím není v cache)"]];

    const sample: SampleTable = {
      headers: ["Měna", "Datum", "Kurz CZK"],
      rows: sampleRows
    };

    return {
      intent,
      issues: [],
      planText: buildPlanList(planItems),
      sample,
      applyPayload: {
        currency: intent.currency,
        targetDate: intent.targetDate
      }
    };
  });
}

function buildFxSample(
  selection: SelectionSnapshot,
  hasHeader: boolean,
  rate: number | null
): SampleTable {
  const rows: string[][] = [];
  const dataRows = selection.sampleValues.slice(hasHeader ? 1 : 0);
  const sampleCount = Math.min(5, dataRows.length);

  for (let i = 0; i < sampleCount; i += 1) {
    const raw = dataRows[i]?.[0];
    const numeric = parseCzechNumeric(raw);
    if (numeric === null || rate === null) {
      rows.push([typeof raw === "string" ? raw : `${raw}`, "(nelze spočítat)"]);
    } else {
      rows.push([formatCzk(numeric), formatCzk(numeric * rate)]);
    }
  }

  if (rows.length === 0) {
    rows.push(["(prázdný řádek)", "-"]);
  }

  return {
    headers: ["Původní částka", "CZK podle ČNB"],
    rows
  };
}

async function buildFxConvertPreview(
  intent: FxConvertCnbIntent
): Promise<PreviewResult<FxConvertCnbApplyPayload>> {
  return Excel.run(async (context) => {
    const selection = await captureSelection(context);
    if ("error" in selection) {
      return { error: selection.error };
    }

    const issues: string[] = [];

    const singleColumnIssue = ensureSingleColumn(selection);
    if (singleColumnIssue) {
      issues.push(singleColumnIssue);
    }

    const columnConflict = validateColumnMatch(intent.columnLetter, selection);
    if (columnConflict) {
      issues.push(columnConflict);
    }

    const hasHeader = detectHeader(selection);
    if (selection.rowCount <= 1) {
      issues.push("Rozsah musí obsahovat alespoň jeden řádek s daty pod hlavičkou.");
    }

    const cachedRate = await getCachedRate(context, intent.currency, intent.targetDate);

    const selectedLetter = columnLetterFromIndex(selection.columnIndex);
    const targetLetter = columnLetterFromIndex(selection.columnIndex + 1);

    const planItems = [
      `Zjistit kurz ČNB pro ${intent.currency} k ${intent.targetDate} (použít cache, jinak stáhnout).`,
      `Vyplnit sloupec ${targetLetter} přepočtenými hodnotami z ${selectedLetter}.`,
      "Nastavit formát CZK a zapsat auditní stopu."
    ];

    const sample = buildFxSample(selection, hasHeader, cachedRate);

    if (issues.length > 0) {
      return { error: "Nelze připravit náhled pro přepočet pomocí ČNB.", issues };
    }

    return {
      intent,
      issues: cachedRate === null ? ["Kurz není v cache, bude nutné online stažení."] : [],
      planText: buildPlanList(planItems),
      sample,
      applyPayload: {
        sheetName: selection.sheetName,
        rowIndex: selection.rowIndex,
        rowCount: selection.rowCount,
        columnIndex: selection.columnIndex,
        currency: intent.currency,
        targetDate: intent.targetDate,
        hasHeader
      }
    };
  });
}

async function buildSeedHolidaysPreview(
  intent: SeedHolidaysIntent
): Promise<PreviewResult<SeedHolidaysApplyPayload>> {
  const entries = listCzechHolidays(intent.year);
  const sampleRows = entries.slice(0, 5).map((entry) => [entry.date, entry.name]);

  const planItems = [
    `Odstranit existující záznamy roku ${intent.year} v _HOLIDAYS_CZ.`,
    "Zapsat nové záznamy včetně Velkého pátku a Velikonočního pondělí.",
    "Zpřístupnit je pro výpočty pracovních dní."
  ];

  return {
    intent,
    issues: [],
    planText: buildPlanList(planItems),
    sample: {
      headers: ["Datum", "Název"],
      rows: sampleRows
    },
    applyPayload: {
      year: intent.year
    }
  };
}

async function buildNetworkdaysPreview(
  intent: NetworkdaysDueIntent
): Promise<PreviewResult<NetworkdaysDueApplyPayload>> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.getActiveWorksheet();
    worksheet.load("name");
    await context.sync();

    const holidaySet = await loadHolidaySet(context);
    const startDate = intent.startDate ? new Date(intent.startDate) : new Date();
    let dueDate: Date;

    try {
      dueDate = calculateBusinessDueDate(startDate, intent.businessDays, holidaySet);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Chyba při výpočtu termínu.";
      return { error: message };
    }

    const startDateISO = formatISODate(startDate);
    const dueDateISO = formatISODate(dueDate);

    const planItems = [
      `Spočítat termín od ${startDateISO} posunutý o ${intent.businessDays} pracovních dní.`,
      "Využít zapsané svátky v _HOLIDAYS_CZ a vynechat víkendy.",
      "Zapsat přehled do buněk H1:I3 na aktuálním listu."
    ];

    const sample: SampleTable = {
      headers: ["Popis", "Hodnota"],
      rows: [
        ["Start", startDateISO],
        ["Pracovní dny", `${intent.businessDays}`],
        ["Termín", dueDateISO]
      ]
    };

    return {
      intent,
      issues: holidaySet.size === 0 ? ["Varování: tabulka svátků je prázdná."] : [],
      planText: buildPlanList(planItems),
      sample,
      applyPayload: {
        sheetName: worksheet.name,
        startCell: "H1",
        businessDays: intent.businessDays,
        startDateISO
      }
    };
  });
}

export async function buildPreview(intent: SupportedIntent): Promise<PreviewResult<ApplyPayload>> {
  switch (intent.type) {
    case IntentType.VatAdd:
      return buildVatPreview(intent);
    case IntentType.FormatCzk:
      return buildFormatPreview(intent);
    case IntentType.FinanceDedupe:
      return buildDedupePreview(intent);
    case IntentType.SortColumn:
      return buildSortPreview(intent);
    case IntentType.VatRemove:
      return buildVatRemovePreview(intent);
    case IntentType.HighlightNegative:
      return buildHighlightNegativePreview(intent);
    case IntentType.SumColumn:
      return buildSumColumnPreview(intent);
    case IntentType.MonthlyRunRate:
      return buildMonthlyRunRatePreview(intent);
    case IntentType.FetchCnbRate:
      return buildFetchCnbPreview(intent);
    case IntentType.FxConvertCnb:
      return buildFxConvertPreview(intent);
    case IntentType.SeedHolidays:
      return buildSeedHolidaysPreview(intent);
    case IntentType.NetworkdaysDue:
      return buildNetworkdaysPreview(intent);
    default:
      return { error: "Náhled pro tento typ požadavku zatím není podporován." };
  }
}
