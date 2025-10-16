const TELEMETRY_TABLE_NAME = "tblTelemetry";

export interface TelemetryEvent {
  event: string;
  intent?: string;
  detail?: string;
}

export async function recordTelemetryEvent(
  context: Excel.RequestContext,
  event: TelemetryEvent
): Promise<void> {
  try {
    const table = context.workbook.tables.getItem(TELEMETRY_TABLE_NAME);
    table.rows.add(undefined, [[new Date().toISOString(), event.event, event.intent ?? "", event.detail ?? ""]]);
  } catch (error) {
    console.warn("Telemetry logging failed", error);
  }
}

export async function logTelemetryEvent(event: TelemetryEvent): Promise<void> {
  try {
    await Excel.run(async (context) => {
      await recordTelemetryEvent(context, event);
    });
  } catch (error) {
    console.warn("Telemetry logging failed", error);
  }
}
