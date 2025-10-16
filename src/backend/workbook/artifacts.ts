const SHEET_DEFINITIONS = [
  {
    name: "_Audit",
    tableName: "tblAudit",
    headers: ["Time", "Intent", "Args", "Range", "Notes"]
  },
  {
    name: "_UndoIndex",
    tableName: "tblUndoIndex",
    headers: ["Time", "Sheet", "Address", "Rows", "Cols", "DataStartRow", "Note"]
  },
  {
    name: "_UndoData",
    tableName: "tblUndoData",
    headers: ["SnapshotId", "RowOffset", "ColOffset", "ValuesJson", "FormulasJson", "FormatsJson"]
  },
  {
    name: "_FX_CNB",
    tableName: "tblFxCnb",
    headers: ["Date", "Code", "Rate"]
  },
  {
    name: "_HOLIDAYS_CZ",
    tableName: "tblHolidaysCz",
    headers: ["Date", "Name"]
  },
  {
    name: "_Settings",
    tableName: "tblSettings",
    headers: ["Key", "Value"]
  }
];

async function ensureWorksheet(
  context: Excel.RequestContext,
  name: string
): Promise<Excel.Worksheet> {
  const sheets = context.workbook.worksheets;
  const maybeSheet = sheets.getItemOrNullObject(name);
  maybeSheet.load("name");
  await context.sync();

  if (!maybeSheet.isNullObject) {
    maybeSheet.visibility = Excel.SheetVisibility.veryHidden;
    return maybeSheet;
  }

  const sheet = sheets.add(name);
  sheet.visibility = Excel.SheetVisibility.veryHidden;
  return sheet;
}

async function ensureTableHeaders(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  tableName: string,
  headers: readonly string[]
) {
  const tables = sheet.tables;
  let table = tables.getItemOrNullObject(tableName);
  table.load("name");
  await context.sync();

  if (table.isNullObject) {
    const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
    const headerValues = Array.from(headers);
    headerRange.values = [headerValues];
    const headerAddress = headerRange.getAddress();
    table = tables.add(headerAddress, true);
    table.name = tableName;
    table.showHeaders = true;
    await context.sync();
    return;
  }

  const currentHeaders = table.getHeaderRowRange();
  currentHeaders.load("values");
  await context.sync();

  const flattenedHeaders = currentHeaders.values[0]?.map((cell) => `${cell}`.trim());
  const needsReset =
    flattenedHeaders?.length !== headers.length ||
    flattenedHeaders.some((value, index) => value !== headers[index]);

  if (needsReset) {
    const headerRange = table.getHeaderRowRange();
    const headerValues = Array.from(headers);
    headerRange.values = [headerValues];
    await context.sync();
  }
}

export async function ensureWorkbookArtifacts() {
  await Excel.run(async (context) => {
    for (const definition of SHEET_DEFINITIONS) {
      const sheet = await ensureWorksheet(context, definition.name);
      await ensureTableHeaders(context, sheet, definition.tableName, definition.headers);
    }
  });
}
