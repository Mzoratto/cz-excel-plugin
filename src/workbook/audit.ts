interface AuditEntry {
  intent: string;
  args: unknown;
  rangeAddress: string;
  note?: string;
}

const AUDIT_TABLE_NAME = "tblAudit";

export async function recordAuditEntry(
  context: Excel.RequestContext,
  entry: AuditEntry
): Promise<void> {
  const table = context.workbook.tables.getItem(AUDIT_TABLE_NAME);
  const timestamp = new Date().toISOString();
  const argsString = (() => {
    try {
      return JSON.stringify(entry.args);
    } catch (error) {
      console.warn("Failed to stringify audit args", error);
      return "";
    }
  })();

  table.rows.add(undefined, [
    [timestamp, entry.intent, argsString, entry.rangeAddress, entry.note ?? ""]
  ]);
}
