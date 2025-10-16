import {
  IntentPreview,
  IntentType,
  VatAddApplyPayload,
  FormatCzkApplyPayload,
  ApplyResult,
  FetchCnbRateApplyPayload,
  FxConvertCnbApplyPayload,
  FinanceDedupeApplyPayload,
  SortColumnApplyPayload,
  VatRemoveApplyPayload,
  HighlightNegativeApplyPayload,
  SumColumnApplyPayload,
  MonthlyRunRateApplyPayload,
  PeriodComparisonApplyPayload,
  SeedHolidaysApplyPayload,
  NetworkdaysDueApplyPayload
} from "../intents/types";
import { columnLetterFromIndex, columnIndexFromLetter } from "../utils/excel";
import { captureUndoSnapshot } from "./undo";
import { recordAuditEntry } from "./audit";
import { ensureCnbRate } from "./cnb";
import { seedCzechHolidays, loadHolidaySet, calculateBusinessDueDate } from "./holidays";
import { recordTelemetryEvent } from "./telemetry";
import { formatISODate } from "../utils/date";
import { parseCzechNumeric } from "../utils/numbers";

const CZK_NUMBER_FORMAT = '[$-cs-CZ]#,##0.00 "Kč"';
const RATE_FORMATTER = new Intl.NumberFormat("cs-CZ", {
  minimumFractionDigits: 4,
  maximumFractionDigits: 4
});

async function applyVat(preview: IntentPreview<VatAddApplyPayload>): Promise<ApplyResult> {
  const { applyPayload } = preview;
  const sourceLetter = columnLetterFromIndex(applyPayload.columnIndex);
  const targetLetter = columnLetterFromIndex(applyPayload.columnIndex + 1);
  const note = `DPH ${applyPayload.rateLabel} pro ${targetLetter}`;

  const snapshotOutcome = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(applyPayload.sheetName);
    const targetRange = sheet.getRangeByIndexes(applyPayload.rowIndex, applyPayload.columnIndex + 1, applyPayload.rowCount, 1);
    const snapshot = await captureUndoSnapshot(context, {
      sheetName: applyPayload.sheetName,
      rowIndex: applyPayload.rowIndex,
      columnIndex: applyPayload.columnIndex + 1,
      rowCount: applyPayload.rowCount,
      columnCount: 1,
      note
    });

    if (applyPayload.hasHeader && applyPayload.rowCount > 0) {
      const headerCell = targetRange.getCell(0, 0);
      headerCell.values = [[`DPH ${applyPayload.rateLabel}`]];
    }

    const dataRowCount = applyPayload.hasHeader ? applyPayload.rowCount - 1 : applyPayload.rowCount;
    if (dataRowCount > 0) {
      const dataStartRowOffset = applyPayload.hasHeader ? 1 : 0;
      const dataRange =
        dataRowCount === applyPayload.rowCount
          ? targetRange
          : targetRange.getCell(dataStartRowOffset, 0).getResizedRange(dataRowCount - 1, 0);

      const formulas = Array.from({ length: dataRowCount }, () => [`=RC[-1]*${applyPayload.rate}`]);
      const formats = Array.from({ length: dataRowCount }, () => [CZK_NUMBER_FORMAT]);

      dataRange.formulasR1C1 = formulas;
      dataRange.numberFormat = formats;
    }

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: {
        rate: applyPayload.rate,
        rateLabel: applyPayload.rateLabel,
        sourceColumn: sourceLetter,
        targetColumn: targetLetter
      },
      rangeAddress: snapshot.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });

    await context.sync();
    return snapshot;
  });

  const warnings = snapshotOutcome.persisted
    ? undefined
    : [
        "Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět."
      ];

  return {
    message: `DPH ${applyPayload.rateLabel} aplikováno: ${sourceLetter} → ${targetLetter}`,
    warnings
  };
}

async function applyFormatCzk(preview: IntentPreview<FormatCzkApplyPayload>): Promise<ApplyResult> {
  const letter = columnLetterFromIndex(preview.applyPayload.columnIndex);
  const note = `Formát CZK pro ${letter}`;

  const snapshotOutcome = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(preview.applyPayload.sheetName);
    const range = sheet.getRangeByIndexes(
      preview.applyPayload.rowIndex,
      preview.applyPayload.columnIndex,
      preview.applyPayload.rowCount,
      1
    );
    const snapshot = await captureUndoSnapshot(context, {
      sheetName: preview.applyPayload.sheetName,
      rowIndex: preview.applyPayload.rowIndex,
      columnIndex: preview.applyPayload.columnIndex,
      rowCount: preview.applyPayload.rowCount,
      columnCount: 1,
      note
    });

    const formats = Array.from({ length: preview.applyPayload.rowCount }, () => [CZK_NUMBER_FORMAT]);
    range.numberFormat = formats;

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: {
        column: letter
      },
      rangeAddress: snapshot.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });

    await context.sync();
    return snapshot;
  });

  const warnings = snapshotOutcome.persisted
    ? undefined
    : [
        "Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět."
      ];

  return {
    message: `Formát CZK nastaven pro sloupec ${letter}`,
    warnings
  };
}

async function applyFinanceDedupe(preview: IntentPreview<FinanceDedupeApplyPayload>): Promise<ApplyResult> {
  const payload = preview.applyPayload;
  const columnLetter = columnLetterFromIndex(payload.columnIndex);
  const note = `Odebrat duplicity ${columnLetter}`;

  const { snapshot, removed } = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(payload.sheetName);
    const range = sheet.getRangeByIndexes(payload.rowIndex, payload.columnIndex, payload.rowCount, payload.columnCount);
    range.load("address");
    await context.sync();

    const undoSnapshot = await captureUndoSnapshot(context, {
      sheetName: payload.sheetName,
      rowIndex: payload.rowIndex,
      columnIndex: payload.columnIndex,
      rowCount: payload.rowCount,
      columnCount: payload.columnCount,
      note
    });

    const columnIndexes = Array.from({ length: payload.columnCount }, (_, index) => index);
    const result = range.removeDuplicates(columnIndexes, payload.hasHeader);
    await context.sync();

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: {
        removed: result.removed,
        uniqueRemaining: result.uniqueRemaining,
        columnCount: payload.columnCount,
        hasHeader: payload.hasHeader
      },
      rangeAddress: range.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });
    await context.sync();

    return { snapshot: undoSnapshot, removed: result.removed };
  });

  const warnings = snapshot.persisted
    ? undefined
    : ["Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět."];

  const message =
    removed > 0
      ? `Odebráno ${removed} duplicitních řádků ve sloupci ${columnLetter}.`
      : `Ve sloupci ${columnLetter} nebyly nalezeny žádné duplicity.`;

  return {
    message,
    warnings
  };
}

async function applySortColumn(preview: IntentPreview<SortColumnApplyPayload>): Promise<ApplyResult> {
  const payload = preview.applyPayload;
  const letter = columnLetterFromIndex(payload.columnIndex);
  const note = `Seřadit ${letter} ${payload.ascending ? "vzestupně" : "sestupně"}`;

  const { snapshot } = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(payload.sheetName);
    const range = sheet.getRangeByIndexes(payload.rowIndex, payload.columnIndex, payload.rowCount, payload.columnCount);

    const undoSnapshot = await captureUndoSnapshot(context, {
      sheetName: payload.sheetName,
      rowIndex: payload.rowIndex,
      columnIndex: payload.columnIndex,
      rowCount: payload.rowCount,
      columnCount: payload.columnCount,
      note
    });

    const sort = range.getSort();
    sort.apply([
      {
        key: 0,
        ascending: payload.ascending,
        sortOn: Excel.SortOn.value
      }
    ], false, payload.hasHeader);

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: {
        column: letter,
        ascending: payload.ascending
      },
      rangeAddress: range.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });
    await context.sync();

    return { snapshot: undoSnapshot };
  });

  const warnings = snapshot.persisted
    ? undefined
    : ["Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět."];

  return {
    message: `Sloupec ${letter} seřazen ${payload.ascending ? "vzestupně" : "sestupně"}.`,
    warnings
  };
}

async function applyVatRemove(preview: IntentPreview<VatRemoveApplyPayload>): Promise<ApplyResult> {
  const payload = preview.applyPayload;
  const sourceLetter = columnLetterFromIndex(payload.columnIndex);
  const baseLetter = columnLetterFromIndex(payload.columnIndex + 1);
  const vatLetter = columnLetterFromIndex(payload.columnIndex + 2);
  const note = `DPH ${payload.rateLabel} z ${sourceLetter}`;

  const { snapshot } = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(payload.sheetName);
    const baseRange = sheet.getRangeByIndexes(payload.rowIndex, payload.columnIndex + 1, payload.rowCount, 1);
    const vatRange = sheet.getRangeByIndexes(payload.rowIndex, payload.columnIndex + 2, payload.rowCount, 1);

    const undoSnapshot = await captureUndoSnapshot(context, {
      sheetName: payload.sheetName,
      rowIndex: payload.rowIndex,
      columnIndex: payload.columnIndex + 1,
      rowCount: payload.rowCount,
      columnCount: 2,
      note
    });

    if (payload.hasHeader && payload.rowCount > 0) {
      baseRange.getCell(0, 0).values = [[`Bez DPH (${payload.rateLabel})`]];
      vatRange.getCell(0, 0).values = [[`DPH ${payload.rateLabel}`]];
    }

    const dataRowCount = payload.hasHeader ? payload.rowCount - 1 : payload.rowCount;
    if (dataRowCount > 0) {
      const startOffset = payload.hasHeader ? 1 : 0;
      const baseDataRange =
        dataRowCount === payload.rowCount
          ? baseRange
          : baseRange.getCell(startOffset, 0).getResizedRange(dataRowCount - 1, 0);
      const vatDataRange =
        dataRowCount === payload.rowCount
          ? vatRange
          : vatRange.getCell(startOffset, 0).getResizedRange(dataRowCount - 1, 0);

      const rateFactor = 1 + payload.rate;
      const baseFormula = `=RC[-1]/${rateFactor}`;
      const vatFormula = "=RC[-2]-RC[-1]";

      baseDataRange.formulasR1C1 = Array.from({ length: dataRowCount }, () => [baseFormula]);
      vatDataRange.formulasR1C1 = Array.from({ length: dataRowCount }, () => [vatFormula]);

      const currencyFormat = Array.from({ length: dataRowCount }, () => [CZK_NUMBER_FORMAT]);
      baseDataRange.numberFormat = currencyFormat;
      vatDataRange.numberFormat = currencyFormat;
    }

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: {
        sourceColumn: sourceLetter,
        baseColumn: baseLetter,
        vatColumn: vatLetter,
        rate: payload.rate,
        rateLabel: payload.rateLabel
      },
      rangeAddress: baseRange.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });
    await context.sync();

    return { snapshot: undoSnapshot };
  });

  const warnings = snapshot.persisted
    ? undefined
    : ["Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět."];

  return {
    message: `Vypočítán základ bez DPH a částka DPH ze sloupce ${sourceLetter}.`,
    warnings
  };
}

async function applyHighlightNegative(
  preview: IntentPreview<HighlightNegativeApplyPayload>
): Promise<ApplyResult> {
  const payload = preview.applyPayload;
  const letter = columnLetterFromIndex(payload.columnIndex);
  const note = `Zvýraznit záporné hodnoty ${letter}`;

  const { snapshot } = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(payload.sheetName);
    const range = sheet.getRangeByIndexes(payload.rowIndex, payload.columnIndex, payload.rowCount, 1);

    const undoSnapshot = await captureUndoSnapshot(context, {
      sheetName: payload.sheetName,
      rowIndex: payload.rowIndex,
      columnIndex: payload.columnIndex,
      rowCount: payload.rowCount,
      columnCount: 1,
      note
    });

	let targetRange = range;
	const dataStartOffset = payload.hasHeader ? 1 : 0;
	const dataRowCount = payload.rowCount - dataStartOffset;
	if (dataRowCount > 0 && dataRowCount !== payload.rowCount) {
	  targetRange = range
	    .getCell(dataStartOffset, 0)
	    .getResizedRange(dataRowCount - 1, 0);
	}

	const conditionalFormat = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
	conditionalFormat.cellValue.rule = {
	  operator: Excel.ConditionalCellValueOperator.lessThan,
	  formula1: "0"
	};
    conditionalFormat.cellValue.format.fill.color = "#fdecea";
    conditionalFormat.cellValue.format.font.color = "#842029";

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: { column: letter },
      rangeAddress: targetRange.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });
    await context.sync();

    return { snapshot: undoSnapshot };
  });

  const warnings = snapshot.persisted
    ? undefined
    : ["Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět."];

  return {
	message: `Záporné hodnoty ve sloupci ${letter} jsou zvýrazněny.`,
    warnings
  };
}

async function applySumColumn(preview: IntentPreview<SumColumnApplyPayload>): Promise<ApplyResult> {
  const payload = preview.applyPayload;
  const letter = columnLetterFromIndex(payload.columnIndex);
  const note = `Součet ve sloupci ${letter}`;

  const result = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(payload.sheetName);
    const totalRowIndex = payload.rowIndex + payload.rowCount;
    const totalCell = sheet.getRangeByIndexes(totalRowIndex, payload.columnIndex, 1, 1);

    const undoSnapshot = await captureUndoSnapshot(context, {
      sheetName: payload.sheetName,
      rowIndex: totalRowIndex,
      columnIndex: payload.columnIndex,
      rowCount: 1,
      columnCount: 1,
      note
    });

    const dataStartRow = payload.rowIndex + (payload.hasHeader ? 1 : 0) + 1;
    const dataEndRow = payload.rowIndex + payload.rowCount;
    const address = `${letter}${dataStartRow}:${letter}${dataEndRow}`;
    totalCell.formulas = [[`=SUM(${address})`]];
    totalCell.numberFormat = [["0.00"]];

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: {
        column: letter,
        range: address
      },
      rangeAddress: totalCell.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });
    await context.sync();

    return { snapshot: undoSnapshot, totalAddress: totalCell.address };
  });

  const warnings = result.snapshot.persisted
    ? undefined
    : ["Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět."];

  return {
    message: `Součet sloupce ${letter} byl zapsán do ${result.totalAddress}.`,
    warnings
  };
}

async function applyMonthlyRunRate(preview: IntentPreview<MonthlyRunRateApplyPayload>): Promise<ApplyResult> {
  const payload = preview.applyPayload;
  const amountAbsolute = columnIndexFromLetter(payload.amountColumnLetter);
  const dateAbsolute = columnIndexFromLetter(payload.dateColumnLetter);
  if (amountAbsolute === null || dateAbsolute === null) {
    throw new Error("Nelze určit sloupce pro výpočet run-rate.");
  }

  const { worksheetName, snapshot } = await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(payload.sheetName);
      const range = sheet.getRangeByIndexes(payload.rowIndex, payload.columnIndex, payload.rowCount, payload.columnCount);
      range.load("values");
      await context.sync();

      const data = range.values as (string | number | boolean)[][];
      const dataStart = payload.hasHeader ? 1 : 0;
      const amountOffset = amountAbsolute - payload.columnIndex;
      const dateOffset = dateAbsolute - payload.columnIndex;

      if (amountOffset < 0 || amountOffset >= payload.columnCount || dateOffset < 0 || dateOffset >= payload.columnCount) {
        throw new Error("Výběr neobsahuje očekávané sloupce pro run-rate.");
      }

      const monthBuckets = new Map<string, { total: number; label: string }>();
      let latestDate: Date | null = null;

      for (let i = dataStart; i < data.length; i += 1) {
        const row = data[i];
        if (!row) {
          continue;
        }
        const dateCell = row[dateOffset];
        const amount = parseCzechNumeric(row[amountOffset]);
        if (amount === null) {
          continue;
        }
        let jsDate: Date;
        if (typeof dateCell === "object" && dateCell !== null && "getTime" in (dateCell as object)) {
          jsDate = new Date((dateCell as Date).getTime());
        } else if (typeof dateCell === "string" || typeof dateCell === "number") {
          jsDate = new Date(dateCell);
        } else {
          continue;
        }
        if (Number.isNaN(jsDate.getTime())) {
          continue;
        }
        const key = `${jsDate.getUTCFullYear()}-${jsDate.getUTCMonth() + 1}`;
        const label = `${jsDate.getUTCFullYear()}-${String(jsDate.getUTCMonth() + 1).padStart(2, "0")}`;
        const bucket = monthBuckets.get(key) ?? { total: 0, label };
        bucket.total += amount;
        monthBuckets.set(key, bucket);
        if (!latestDate || jsDate > latestDate) {
          latestDate = jsDate;
        }
      }

      if (monthBuckets.size === 0) {
        throw new Error("Ve vybraném rozsahu chybí numerická data nebo datum.");
      }

      const sortedKeys = Array.from(monthBuckets.keys()).sort((a, b) => {
        const [ay, am] = a.split("-").map(Number);
        const [by, bm] = b.split("-").map(Number);
        return ay === by ? am - bm : ay - by;
      });

      const windowSize = payload.months;
      const windowKeys = sortedKeys.slice(-windowSize);
      const monthlyTotals = windowKeys.map((key) => monthBuckets.get(key)!);
      const sumTotals = monthlyTotals.reduce((acc, item) => acc + item.total, 0);
      const average = sumTotals / (windowKeys.length || 1);
      const annualizedValue = average * 12;

      const workbook = context.workbook;
      let outputSheet = workbook.worksheets.getItemOrNullObject("_RunRate");
      await context.sync();
      if (outputSheet.isNullObject) {
        outputSheet = workbook.worksheets.add("_RunRate");
      }

      const usedRange = outputSheet.getUsedRangeOrNullObject();
      usedRange.load("address");
      await context.sync();
      if (!usedRange.isNullObject) {
        usedRange.clear();
      }

      const summaryRows = [
        ["Období", `${windowKeys.length} měsíců`],
        ["Součet za období", sumTotals],
        ["Průměr měsíčně", average],
        ["Roční run-rate", annualizedValue],
        [
          "Poznámka",
          latestDate
            ? `Data k ${latestDate.toLocaleDateString("cs-CZ")} z ${payload.sheetName}!${payload.amountColumnLetter}`
            : `Data z listu ${payload.sheetName}`
        ]
      ];

      const outputRange = outputSheet.getRangeByIndexes(0, 0, summaryRows.length, 2);
      const undoSnapshot = await captureUndoSnapshot(context, {
        sheetName: outputSheet.name,
        rowIndex: 0,
        columnIndex: 0,
        rowCount: Math.max(summaryRows.length, monthlyTotals.length + 2),
        columnCount: 3,
        note: "Monthly run-rate"
      });

      outputRange.values = summaryRows;
      const currencyRows = outputSheet.getRangeByIndexes(1, 1, 3, 1);
      currencyRows.numberFormat = [[CZK_NUMBER_FORMAT], [CZK_NUMBER_FORMAT], [CZK_NUMBER_FORMAT]];

      if (monthlyTotals.length > 0) {
        const detailRange = outputSheet.getRangeByIndexes(summaryRows.length + 1, 0, monthlyTotals.length + 1, 2);
        const detailValues: (string | number)[][] = [["Měsíc", "Součet"], ...monthlyTotals.map((item) => [item.label, item.total])];
        detailRange.values = detailValues;
        if (monthlyTotals.length > 0) {
          const detailValuesRange = outputSheet
            .getRangeByIndexes(summaryRows.length + 2, 1, monthlyTotals.length, 1);
          detailValuesRange.numberFormat = Array.from({ length: monthlyTotals.length }, () => [CZK_NUMBER_FORMAT]);
        }
      }

      await recordAuditEntry(context, {
        intent: preview.intent.type,
        args: {
          months: windowKeys.length,
          amountColumn: payload.amountColumnLetter,
          dateColumn: payload.dateColumnLetter
        },
        rangeAddress: outputRange.address,
        note: "Run-rate summary"
      });
      await recordTelemetryEvent(context, {
        event: "apply",
        intent: preview.intent.type
      });
      await context.sync();

      return {
        worksheetName: outputSheet.name,
        snapshot: undoSnapshot
      };
    });

  const warnings = snapshot.persisted
    ? undefined
    : ["Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět."];

  return {
    message: `Run-rate připraven na listu ${worksheetName}.`,
    warnings
  };
}

async function applyPeriodComparison(
  preview: IntentPreview<PeriodComparisonApplyPayload>
): Promise<ApplyResult> {
  const payload = preview.applyPayload;
  const amountAbsolute = columnIndexFromLetter(payload.amountColumnLetter);
  const dateAbsolute = columnIndexFromLetter(payload.dateColumnLetter);
  if (amountAbsolute === null || dateAbsolute === null) {
    throw new Error("Nelze určit sloupce pro porovnání období.");
  }

  const result = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(payload.sheetName);
    const range = sheet.getRangeByIndexes(payload.rowIndex, payload.columnIndex, payload.rowCount, payload.columnCount);
    range.load("values");
    await context.sync();

    const data = range.values as (string | number | boolean | Date)[][];
    const dataStart = payload.hasHeader ? 1 : 0;
    const amountOffset = amountAbsolute - payload.columnIndex;
    const dateOffset = dateAbsolute - payload.columnIndex;

    const currentTotals = {
      month: 0,
      prevMonth: 0,
      quarter: 0,
      prevQuarter: 0,
      year: 0,
      prevYear: 0
    };

    let latestDate: Date | null = null;

    const monthTotals = new Map<string, number>();
    const quarterTotals = new Map<string, number>();
    const yearTotals = new Map<number, number>();

    for (let i = dataStart; i < data.length; i += 1) {
      const row = data[i];
      if (!row) {
        continue;
      }
      const dateCell = row[dateOffset];
      const amount = parseCzechNumeric(row[amountOffset]);
      if (amount === null) {
        continue;
      }
      let jsDate: Date;
      if (typeof dateCell === "object" && dateCell !== null && "getTime" in (dateCell as object)) {
        jsDate = new Date((dateCell as Date).getTime());
      } else if (typeof dateCell === "string" || typeof dateCell === "number") {
        jsDate = new Date(dateCell);
      } else {
        continue;
      }
      if (Number.isNaN(jsDate.getTime())) {
        continue;
      }
      const monthKey = `${jsDate.getUTCFullYear()}-${jsDate.getUTCMonth() + 1}`;
      const quarterKey = `${jsDate.getUTCFullYear()}-Q${Math.floor(jsDate.getUTCMonth() / 3) + 1}`;
      const yearKey = jsDate.getUTCFullYear();

      monthTotals.set(monthKey, (monthTotals.get(monthKey) ?? 0) + amount);
      quarterTotals.set(quarterKey, (quarterTotals.get(quarterKey) ?? 0) + amount);
      yearTotals.set(yearKey, (yearTotals.get(yearKey) ?? 0) + amount);

      if (!latestDate || jsDate > latestDate) {
        latestDate = jsDate;
      }
    }

    if (!latestDate) {
      throw new Error("Výběr neobsahuje platná data pro porovnání.");
    }

    const toMonthKey = (date: Date) => `${date.getUTCFullYear()}-${date.getUTCMonth() + 1}`;
    const toQuarterKey = (date: Date) => `${date.getUTCFullYear()}-Q${Math.floor(date.getUTCMonth() / 3) + 1}`;

    const currentMonthKey = toMonthKey(latestDate);
    const prevMonthDate = new Date(Date.UTC(latestDate.getUTCFullYear(), latestDate.getUTCMonth() - 1, 1));
    const prevMonthKey = toMonthKey(prevMonthDate);

    const currentQuarterKey = toQuarterKey(latestDate);
    const prevQuarterDate = new Date(Date.UTC(latestDate.getUTCFullYear(), latestDate.getUTCMonth() - 3, 1));
    const prevQuarterKey = toQuarterKey(prevQuarterDate);

    const currentYear = latestDate.getUTCFullYear();
    const prevYear = currentYear - 1;

    currentTotals.month = monthTotals.get(currentMonthKey) ?? 0;
    currentTotals.prevMonth = monthTotals.get(prevMonthKey) ?? 0;
    currentTotals.quarter = quarterTotals.get(currentQuarterKey) ?? 0;
    currentTotals.prevQuarter = quarterTotals.get(prevQuarterKey) ?? 0;
    currentTotals.year = yearTotals.get(currentYear) ?? 0;
    currentTotals.prevYear = yearTotals.get(prevYear) ?? 0;

    const calcDelta = (current: number, previous: number) => ({
      abs: current - previous,
      pct: previous === 0 ? null : (current - previous) / previous
    });

    const monthDelta = calcDelta(currentTotals.month, currentTotals.prevMonth);
    const quarterDelta = calcDelta(currentTotals.quarter, currentTotals.prevQuarter);
    const yearDelta = calcDelta(currentTotals.year, currentTotals.prevYear);

    const workbook = context.workbook;
    let outputSheet = workbook.worksheets.getItemOrNullObject("_PeriodCompare");
    await context.sync();
    if (outputSheet.isNullObject) {
      outputSheet = workbook.worksheets.add("_PeriodCompare");
    }

    const usedRange = outputSheet.getUsedRangeOrNullObject();
    await context.sync();
    if (!usedRange.isNullObject) {
      usedRange.clear();
    }

    const summaryRows = [
      ["Období", "Aktuální", "Předchozí", "Δ absolutní", "Δ %"],
      [
        "Měsíc",
        currentTotals.month,
        currentTotals.prevMonth,
        monthDelta.abs,
        monthDelta.pct
      ],
      [
        "Čtvrtletí",
        currentTotals.quarter,
        currentTotals.prevQuarter,
        quarterDelta.abs,
        quarterDelta.pct
      ],
      [
        "Rok",
        currentTotals.year,
        currentTotals.prevYear,
        yearDelta.abs,
        yearDelta.pct
      ]
    ];

    const outputRange = outputSheet.getRangeByIndexes(0, 0, summaryRows.length, summaryRows[0]!.length);
    const undoSnapshot = await captureUndoSnapshot(context, {
      sheetName: outputSheet.name,
      rowIndex: 0,
      columnIndex: 0,
      rowCount: summaryRows.length,
      columnCount: summaryRows[0]!.length,
      note: "Period comparison"
    });

    outputRange.values = summaryRows.map((row) =>
      row.map((value) => (typeof value === "number" && Number.isFinite(value) ? value : value ?? ""))
    );

    const valueRange = outputRange.getOffsetRange(1, 1).getResizedRange(summaryRows.length - 2, 2);
    valueRange.numberFormat = Array.from({ length: valueRange.rowCount }, () => [CZK_NUMBER_FORMAT, CZK_NUMBER_FORMAT]);

    const percentRange = outputRange.getOffsetRange(1, 4).getResizedRange(summaryRows.length - 2, 1);
    percentRange.numberFormat = Array.from({ length: percentRange.rowCount }, () => ["0.00%"]);

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: {
        monthCurrent: currentTotals.month,
        monthPrevious: currentTotals.prevMonth,
        quarterCurrent: currentTotals.quarter,
        quarterPrevious: currentTotals.prevQuarter,
        yearCurrent: currentTotals.year,
        yearPrevious: currentTotals.prevYear,
        columns: {
          amount: payload.amountColumnLetter,
          date: payload.dateColumnLetter
        }
      },
      rangeAddress: outputRange.address,
      note: "Period comparison"
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });
    await context.sync();

    return {
      worksheetName: outputSheet.name,
      snapshot: undoSnapshot
    };
  });

  const warnings = result.snapshot.persisted
    ? undefined
    : ["Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět."];

  return {
    message: `Porovnání období připraveno na listu ${result.worksheetName}.`,
    warnings
  };
}

async function applyFetchCnbRate(
  preview: IntentPreview<FetchCnbRateApplyPayload>
): Promise<ApplyResult> {
  const { currency, targetDate } = preview.applyPayload;
  const note = `Kurz ${currency} ${targetDate}`;

  const { rate, source } = await Excel.run(async (context) => {
    let outcome;
    try {
      outcome = await ensureCnbRate(context, currency, targetDate);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Nepodařilo se získat kurz ČNB.";
      throw new Error(message);
    }

    const table = context.workbook.tables.getItem("tblFxCnb");
    const range = table.getDataBodyRange();
    range.load("address");
    await context.sync();

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: { currency, targetDate, source: outcome.source, rate: outcome.rate },
      rangeAddress: range.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });
    await context.sync();

    return outcome;
  });

  const sourceLabel = source === "cache" ? "z cache" : "staženo z ČNB";
  return {
    message: `Kurz ${currency} k ${targetDate}: ${RATE_FORMATTER.format(rate)} CZK (${sourceLabel})`
  };
}

async function applyFxConvertCnb(
  preview: IntentPreview<FxConvertCnbApplyPayload>
): Promise<ApplyResult> {
  const payload = preview.applyPayload;
  const sourceLetter = columnLetterFromIndex(payload.columnIndex);
  const targetLetter = columnLetterFromIndex(payload.columnIndex + 1);
  const note = `ČNB ${payload.currency} → CZK ${targetLetter}`;

  const { snapshot, rate, source } = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(payload.sheetName);
    const targetRange = sheet.getRangeByIndexes(payload.rowIndex, payload.columnIndex + 1, payload.rowCount, 1);

    let rateOutcome;
    try {
      rateOutcome = await ensureCnbRate(context, payload.currency, payload.targetDate);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Nepodařilo se získat kurz ČNB.";
      throw new Error(message);
    }

    const undoSnapshot = await captureUndoSnapshot(context, {
      sheetName: payload.sheetName,
      rowIndex: payload.rowIndex,
      columnIndex: payload.columnIndex + 1,
      rowCount: payload.rowCount,
      columnCount: 1,
      note
    });

    if (payload.hasHeader && payload.rowCount > 0) {
      const headerCell = targetRange.getCell(0, 0);
      headerCell.values = [[`CZK (${payload.currency})`]];
    }

    const dataRowCount = payload.hasHeader ? payload.rowCount - 1 : payload.rowCount;
    if (dataRowCount > 0) {
      const dataStartOffset = payload.hasHeader ? 1 : 0;
      const dataRange =
        dataRowCount === payload.rowCount
          ? targetRange
          : targetRange.getCell(dataStartOffset, 0).getResizedRange(dataRowCount - 1, 0);
      const formulas = Array.from({ length: dataRowCount }, () => [`=RC[-1]*${rateOutcome.rate}`]);
      const formats = Array.from({ length: dataRowCount }, () => [CZK_NUMBER_FORMAT]);

      dataRange.formulasR1C1 = formulas;
      dataRange.numberFormat = formats;
    }

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: {
        currency: payload.currency,
        targetDate: payload.targetDate,
        rate: rateOutcome.rate,
        source: rateOutcome.source,
        sourceColumn: sourceLetter,
        targetColumn: targetLetter
      },
      rangeAddress: undoSnapshot.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });
    await context.sync();

    return { snapshot: undoSnapshot, rate: rateOutcome.rate, source: rateOutcome.source };
  });

  const warnings = snapshot.persisted
    ? undefined
    : [
        "Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět."
      ];

  const sourceLabel = source === "cache" ? "z cache" : "staženo z ČNB";
  return {
    message: `Sloupec ${sourceLetter} přepočten na CZK (${RATE_FORMATTER.format(rate)} CZK/${preview.applyPayload.currency}, ${sourceLabel}).`,
    warnings
  };
}

async function applySeedHolidays(
  preview: IntentPreview<SeedHolidaysApplyPayload>
): Promise<ApplyResult> {
  const year = preview.applyPayload.year;
  const note = `Svátky ${year}`;

  const { entries, snapshot } = await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem("tblHolidaysCz");
    const sheetName = "_HOLIDAYS_CZ";
    const existingRange = table.getDataBodyRangeOrNullObject();
    existingRange.load(["rowCount", "columnCount", "rowIndex", "columnIndex", "isNullObject"]);
    await context.sync();

    let undoSnapshot = null;
    if (!existingRange.isNullObject && existingRange.rowCount > 0 && existingRange.columnCount > 0) {
      undoSnapshot = await captureUndoSnapshot(context, {
        sheetName,
        rowIndex: existingRange.rowIndex,
        columnIndex: existingRange.columnIndex,
        rowCount: existingRange.rowCount,
        columnCount: existingRange.columnCount,
        note
      });
    }

    const seededEntries = await seedCzechHolidays(context, year);
    const rangeAfter = table.getDataBodyRange();
    rangeAfter.load("address");
    await context.sync();

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: { year, count: seededEntries.length },
      rangeAddress: rangeAfter.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });
    await context.sync();

    return { entries: seededEntries, snapshot: undoSnapshot };
  });

  return {
    message: `Tabulka _HOLIDAYS_CZ aktualizována pro rok ${year} (${entries.length} záznamů).`,
    warnings: snapshot && !snapshot.persisted ? ["Snapshot svátků nebyl uložen trvale."] : undefined
  };
}

async function applyNetworkdaysDue(
  preview: IntentPreview<NetworkdaysDueApplyPayload>
): Promise<ApplyResult> {
  const payload = preview.applyPayload;
  const note = `SLA ${payload.businessDays} dní`;

  const result = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(payload.sheetName);
    const startRange = sheet.getRange(payload.startCell);
    startRange.load(["rowIndex", "columnIndex"]);
    await context.sync();

    const targetRange = startRange.getResizedRange(2, 1);

    const undoSnapshot = await captureUndoSnapshot(context, {
      sheetName: payload.sheetName,
      rowIndex: startRange.rowIndex,
      columnIndex: startRange.columnIndex,
      rowCount: 3,
      columnCount: 2,
      note
    });

    const holidaySet = await loadHolidaySet(context);
    const startDate = new Date(payload.startDateISO);
    const dueDate = calculateBusinessDueDate(startDate, payload.businessDays, holidaySet);
    const dueDateISO = formatISODate(dueDate);

    const values = [
      ["Počet pracovních dní", payload.businessDays],
      ["Start", startDate],
      ["Termín", dueDate]
    ];

    targetRange.values = values;

    const valueColumn = targetRange.getColumn(1);
    valueColumn.numberFormat = [["0"], ["dd.mm.yyyy"], ["dd.mm.yyyy"]];

    await recordAuditEntry(context, {
      intent: preview.intent.type,
      args: {
        businessDays: payload.businessDays,
        startDate: payload.startDateISO,
        dueDate: dueDateISO
      },
      rangeAddress: targetRange.address,
      note
    });
    await recordTelemetryEvent(context, {
      event: "apply",
      intent: preview.intent.type
    });
    await context.sync();

    return {
      dueDateISO,
      snapshot: undoSnapshot,
      holidaySetSize: holidaySet.size
    };
  });

  const warnings: string[] = [];
  if (!result.snapshot.persisted) {
    warnings.push("Operace příliš velká pro trvalé Zpět; aktuální stav lze vrátit jen pomocí poslední akce Zpět.");
  }
  if (result.holidaySetSize === 0) {
    warnings.push("Upozornění: Tabulka svátků je prázdná, termín nemusí zohledňovat volné dny.");
  }

  return {
    message: `Termín posunutý o ${preview.applyPayload.businessDays} pracovních dní: ${result.dueDateISO}`,
    warnings: warnings.length > 0 ? warnings : undefined
  };
}

export async function applyIntent(preview: IntentPreview): Promise<ApplyResult> {
  switch (preview.intent.type) {
    case IntentType.VatAdd:
      return applyVat(preview as IntentPreview<VatAddApplyPayload>);
    case IntentType.FormatCzk:
      return applyFormatCzk(preview as IntentPreview<FormatCzkApplyPayload>);
    case IntentType.FinanceDedupe:
      return applyFinanceDedupe(preview as IntentPreview<FinanceDedupeApplyPayload>);
    case IntentType.SortColumn:
      return applySortColumn(preview as IntentPreview<SortColumnApplyPayload>);
    case IntentType.VatRemove:
      return applyVatRemove(preview as IntentPreview<VatRemoveApplyPayload>);
    case IntentType.HighlightNegative:
      return applyHighlightNegative(preview as IntentPreview<HighlightNegativeApplyPayload>);
    case IntentType.SumColumn:
      return applySumColumn(preview as IntentPreview<SumColumnApplyPayload>);
    case IntentType.MonthlyRunRate:
      return applyMonthlyRunRate(preview as IntentPreview<MonthlyRunRateApplyPayload>);
    case IntentType.PeriodComparison:
      return applyPeriodComparison(preview as IntentPreview<PeriodComparisonApplyPayload>);
    case IntentType.FetchCnbRate:
      return applyFetchCnbRate(preview as IntentPreview<FetchCnbRateApplyPayload>);
    case IntentType.FxConvertCnb:
      return applyFxConvertCnb(preview as IntentPreview<FxConvertCnbApplyPayload>);
    case IntentType.SeedHolidays:
      return applySeedHolidays(preview as IntentPreview<SeedHolidaysApplyPayload>);
    case IntentType.NetworkdaysDue:
      return applyNetworkdaysDue(preview as IntentPreview<NetworkdaysDueApplyPayload>);
    default:
      throw new Error("Tato operace zatím není podporována.");
  }
}
