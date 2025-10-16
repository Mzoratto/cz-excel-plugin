import { columnIndexFromLetter } from "../utils/excel";

const MEMORY_STACK: UndoSnapshot[] = [];
const CELL_PERSISTENCE_CAP = 2000;
const UNDO_INDEX_TABLE = "tblUndoIndex";
const UNDO_DATA_TABLE = "tblUndoData";

interface UndoSnapshot {
  id: string;
  timestamp: string;
  sheetName: string;
  address: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  values: (string | number | boolean | null)[][];
  formulasR1C1: (string | null)[][];
  numberFormat: (string | null)[][];
  note: string;
  persisted: boolean;
  cellCount: number;
}

interface CaptureOptions {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  note: string;
}

interface PersistentRow {
  snapshot: UndoSnapshot;
  persisted: boolean;
}

function clone2DArray<T>(source: T[][] | undefined | null): T[][] {
  if (!Array.isArray(source)) {
    return [];
  }
  return source.map((row) => (Array.isArray(row) ? row.map((value) => value) : []));
}

function generateSnapshotId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function buildNoteWithId(note: string, id: string): string {
  if (!note) {
    return `id:${id}`;
  }
  return `${note} ::: id:${id}`;
}

function extractNoteAndId(rawNote: string): { note: string; id?: string } {
  if (!rawNote) {
    return { note: "" };
  }
  const parts = rawNote.split(":::").map((part) => part.trim());
  if (parts.length === 1) {
    return { note: parts[0] };
  }
  const idPart = parts[parts.length - 1];
  if (idPart.startsWith("id:")) {
    const id = idPart.slice(3).trim();
    const note = parts.slice(0, -1).join(" ::: ").trim();
    return { note, id };
  }
  return { note: rawNote };
}

function parseAddress(address: string): {
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
} {
  const withoutSheet = address.includes("!") ? address.split("!")[1] : address;
  const clean = withoutSheet.replace(/\$/g, "");
  const [startPart, endPart] = clean.split(":");
  const startMatch = startPart.match(/^([A-Z]+)(\d+)$/i);
  const endMatch = (endPart ?? startPart).match(/^([A-Z]+)(\d+)$/i);

  if (!startMatch || !endMatch) {
    throw new Error(`Nelze zpracovat adresu pro undo: ${address}`);
  }

  const startColumnIndex = columnIndexFromLetter(startMatch[1]!.toUpperCase());
  const endColumnIndex = columnIndexFromLetter(endMatch[1]!.toUpperCase());
  if (startColumnIndex === null || endColumnIndex === null) {
    throw new Error(`Neplatný sloupec v adrese ${address}`);
  }

  const startRow = parseInt(startMatch[2]!, 10) - 1;
  const endRow = parseInt(endMatch[2]!, 10) - 1;

  return {
    rowIndex: startRow,
    columnIndex: startColumnIndex,
    rowCount: endRow - startRow + 1,
    columnCount: endColumnIndex - startColumnIndex + 1
  };
}

export async function captureUndoSnapshot(
  context: Excel.RequestContext,
  options: CaptureOptions
): Promise<UndoSnapshot> {
  const { sheetName, rowIndex, columnIndex, rowCount, columnCount, note } = options;

  const sheet = context.workbook.worksheets.getItem(sheetName);
  const range = sheet.getRangeByIndexes(rowIndex, columnIndex, rowCount, columnCount);
  range.load(["address", "values", "formulasR1C1", "numberFormat"]);
  await context.sync();

  const snapshot: UndoSnapshot = {
    id: generateSnapshotId(),
    timestamp: new Date().toISOString(),
    sheetName,
    address: range.address,
    rowIndex,
    columnIndex,
    rowCount,
    columnCount,
    values: clone2DArray(range.values as (string | number | boolean | null)[][]),
    formulasR1C1: clone2DArray(range.formulasR1C1 as (string | null)[][]),
    numberFormat: clone2DArray(range.numberFormat as (string | null)[][]),
    note,
    persisted: false,
    cellCount: rowCount * columnCount
  };

  if (snapshot.cellCount <= CELL_PERSISTENCE_CAP) {
    const persistedRow = await persistSnapshot(context, snapshot);
    snapshot.persisted = persistedRow.persisted;
  }

  MEMORY_STACK.push(snapshot);
  return snapshot;
}

async function persistSnapshot(
  context: Excel.RequestContext,
  snapshot: UndoSnapshot
): Promise<PersistentRow> {
  const indexTable = context.workbook.tables.getItem(UNDO_INDEX_TABLE);
  const dataTable = context.workbook.tables.getItem(UNDO_DATA_TABLE);

  const noteWithId = buildNoteWithId(snapshot.note, snapshot.id);

  indexTable.rows.add(undefined, [
    [
      snapshot.timestamp,
      snapshot.sheetName,
      snapshot.address,
      snapshot.rowCount,
      snapshot.columnCount,
      snapshot.rowIndex + 1,
      noteWithId
    ]
  ]);

  dataTable.rows.add(undefined, [
    [
      snapshot.id,
      0,
      0,
      JSON.stringify(snapshot.values),
      JSON.stringify(snapshot.formulasR1C1),
      JSON.stringify(snapshot.numberFormat)
    ]
  ]);

  return { snapshot, persisted: true };
}

function buildValueMatrixForRestore(snapshot: UndoSnapshot): (string | number | boolean | null)[][] {
  return snapshot.values.map((row, rowIdx) =>
    row.map((value, colIdx) => {
      const formula = snapshot.formulasR1C1[rowIdx]?.[colIdx];
      return formula ? null : value;
    })
  );
}

async function restoreSnapshot(snapshot: UndoSnapshot): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(snapshot.sheetName);
    const range = sheet.getRangeByIndexes(
      snapshot.rowIndex,
      snapshot.columnIndex,
      snapshot.rowCount,
      snapshot.columnCount
    );

    const normalizedFormulas = snapshot.formulasR1C1.map((row) =>
      row.map((cell) => (cell == null ? "" : cell))
    );
    const normalizedFormats = snapshot.numberFormat.map((row) =>
      row.map((cell) => (cell == null ? "" : cell))
    );

    range.formulasR1C1 = normalizedFormulas;
    range.values = buildValueMatrixForRestore(snapshot);
    range.numberFormat = normalizedFormats;

    await context.sync();
  });
}

async function removePersistentRows(snapshotId: string): Promise<void> {
  await Excel.run(async (context) => {
    const indexTable = context.workbook.tables.getItem(UNDO_INDEX_TABLE);
    const dataTable = context.workbook.tables.getItem(UNDO_DATA_TABLE);

    const indexRange = indexTable.getDataBodyRangeOrNullObject();
    indexRange.load("values, rowCount, isNullObject");
    const dataRange = dataTable.getDataBodyRangeOrNullObject();
    dataRange.load("values, rowCount, isNullObject");
    await context.sync();

    if (!indexRange.isNullObject && indexRange.rowCount > 0) {
      for (let i = indexRange.rowCount - 1; i >= 0; i -= 1) {
        const row = indexRange.values[i];
        const rawNote = `${row[6] ?? ""}`;
        const { id } = extractNoteAndId(rawNote);
        if (id === snapshotId) {
          indexTable.rows.getItemAt(i).delete();
          break;
        }
      }
    }

    if (!dataRange.isNullObject && dataRange.rowCount > 0) {
      for (let i = dataRange.rowCount - 1; i >= 0; i -= 1) {
        const row = dataRange.values[i];
        if (`${row[0]}` === snapshotId) {
          dataTable.rows.getItemAt(i).delete();
          break;
        }
      }
    }
    await context.sync();
  });
}

async function popPersistentSnapshot(): Promise<UndoSnapshot | null> {
  return Excel.run(async (context) => {
    const indexTable = context.workbook.tables.getItem(UNDO_INDEX_TABLE);
    const dataTable = context.workbook.tables.getItem(UNDO_DATA_TABLE);

    const indexRange = indexTable.getDataBodyRangeOrNullObject();
    indexRange.load(["rowCount", "values", "isNullObject"]);
    await context.sync();

    if (indexRange.isNullObject || indexRange.rowCount === 0) {
      return null;
    }

    const lastRowIndex = indexRange.rowCount - 1;
    const row = indexRange.values[lastRowIndex];
    const timestamp = `${row[0] ?? ""}`;
    const sheetName = `${row[1] ?? ""}`;
    const address = `${row[2] ?? ""}`;
    const rowCount = Number(row[3] ?? 0);
    const columnCount = Number(row[4] ?? 0);
    const noteInfo = extractNoteAndId(`${row[6] ?? ""}`);
    const snapshotId = noteInfo.id ?? `${timestamp}-${sheetName}-${address}`;

    const dataRange = dataTable.getDataBodyRangeOrNullObject();
    dataRange.load(["rowCount", "values", "isNullObject"]);
    await context.sync();

    let valuesJson = "[]";
    let formulasJson = "[]";
    let formatsJson = "[]";

    if (!dataRange.isNullObject && dataRange.rowCount > 0) {
      for (let i = dataRange.rowCount - 1; i >= 0; i -= 1) {
        const dataRow = dataRange.values[i];
        if (`${dataRow[0]}` === snapshotId) {
          valuesJson = `${dataRow[3] ?? "[]"}`;
          formulasJson = `${dataRow[4] ?? "[]"}`;
          formatsJson = `${dataRow[5] ?? "[]"}`;
          break;
        }
      }
    }

    const dimensions = parseAddress(address);

    const snapshot: UndoSnapshot = {
      id: snapshotId,
      timestamp,
      sheetName,
      address,
      rowIndex: dimensions.rowIndex,
      columnIndex: dimensions.columnIndex,
      rowCount: rowCount || dimensions.rowCount,
      columnCount: columnCount || dimensions.columnCount,
      values: JSON.parse(valuesJson),
      formulasR1C1: JSON.parse(formulasJson),
      numberFormat: JSON.parse(formatsJson),
      note: noteInfo.note,
      persisted: true,
      cellCount: (rowCount || dimensions.rowCount) * (columnCount || dimensions.columnCount)
    };

    indexTable.rows.getItemAt(lastRowIndex).delete();
    if (!dataRange.isNullObject && dataRange.rowCount > 0) {
      for (let i = dataRange.rowCount - 1; i >= 0; i -= 1) {
        const dataRow = dataRange.values[i];
        if (`${dataRow[0]}` === snapshotId) {
          dataTable.rows.getItemAt(i).delete();
          break;
        }
      }
    }
    await context.sync();

    return snapshot;
  });
}

export async function performUndo(): Promise<{ success: boolean; message: string }> {
  const snapshot = MEMORY_STACK.pop() ?? (await popPersistentSnapshot());
  if (!snapshot) {
    return { success: false, message: "Žádná akce k vrácení." };
  }

  await restoreSnapshot(snapshot);

  if (snapshot.persisted) {
    await removePersistentRows(snapshot.id);
  }

  const note = snapshot.note ? ` (${snapshot.note})` : "";
  return { success: true, message: `Vrácena poslední akce${note}.` };
}
