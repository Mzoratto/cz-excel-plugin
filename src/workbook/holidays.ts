import { formatISODate } from "../utils/date";

interface HolidayEntry {
  date: string;
  name: string;
}

const HOLIDAY_TABLE_NAME = "tblHolidaysCz";
const BUSINESS_DAYS_LIMIT = 1000;

function computeEasterSunday(year: number): Date {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
}

function addDays(base: Date, offset: number): Date {
  const copy = new Date(base);
  copy.setDate(copy.getDate() + offset);
  return copy;
}

function computeHolidayEntries(year: number): HolidayEntry[] {
  const easterSunday = computeEasterSunday(year);
  const goodFriday = addDays(easterSunday, -2);
  const easterMonday = addDays(easterSunday, 1);

  const fixedDates: Array<[number, number, string]> = [
    [1, 1, "Nový rok"],
    [5, 1, "Svátek práce"],
    [5, 8, "Den vítězství"],
    [7, 5, "Den slovanských věrozvěstů"],
    [7, 6, "Den upálení mistra Jana Husa"],
    [9, 28, "Den české státnosti"],
    [10, 28, "Vznik samostatného československého státu"],
    [11, 17, "Den boje za svobodu a demokracii"],
    [12, 24, "Štědrý den"],
    [12, 25, "1. svátek vánoční"],
    [12, 26, "2. svátek vánoční"]
  ];

  const entries: HolidayEntry[] = [
    { date: formatISODate(goodFriday), name: "Velký pátek" },
    { date: formatISODate(easterMonday), name: "Velikonoční pondělí" }
  ];

  for (const [month, day, name] of fixedDates) {
    entries.push({
      date: formatISODate(new Date(year, month - 1, day)),
      name
    });
  }

  entries.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  return entries;
}

export function listCzechHolidays(year: number): HolidayEntry[] {
  return computeHolidayEntries(year);
}

export async function seedCzechHolidays(
  context: Excel.RequestContext,
  year: number
): Promise<HolidayEntry[]> {
  const entries = computeHolidayEntries(year);
  const table = context.workbook.tables.getItem(HOLIDAY_TABLE_NAME);
  const range = table.getDataBodyRangeOrNullObject();
  range.load(["values", "rowCount", "isNullObject"]);
  await context.sync();

  if (!range.isNullObject && range.rowCount > 0) {
    const rowsToDelete: number[] = [];
    for (let index = 0; index < range.rowCount; index += 1) {
      const row = range.values[index];
      const dateCell = `${row[0] ?? ""}`;
      if (dateCell.startsWith(`${year}-`)) {
        rowsToDelete.push(index);
      }
    }

    for (let i = rowsToDelete.length - 1; i >= 0; i -= 1) {
      table.rows.getItemAt(rowsToDelete[i]).delete();
    }
  }

  table.rows.add(
    undefined,
    entries.map((entry) => [entry.date, entry.name])
  );

  await context.sync();

  return entries;
}

export async function loadHolidaySet(context: Excel.RequestContext): Promise<Set<string>> {
  const table = context.workbook.tables.getItem(HOLIDAY_TABLE_NAME);
  const range = table.getDataBodyRangeOrNullObject();
  range.load(["values", "rowCount", "isNullObject"]);
  await context.sync();

  const set = new Set<string>();
  if (!range.isNullObject && range.rowCount > 0) {
    for (const row of range.values) {
      const dateCell = `${row[0] ?? ""}`;
      if (dateCell) {
        set.add(dateCell);
      }
    }
  }

  return set;
}

function isBusinessDay(date: Date, holidaySet: Set<string>): boolean {
  const day = date.getDay();
  if (day === 0 || day === 6) {
    return false;
  }

  const iso = formatISODate(date);
  return !holidaySet.has(iso);
}

export function calculateBusinessDueDate(
  startDate: Date,
  businessDays: number,
  holidaySet: Set<string>
): Date {
  const direction = businessDays >= 0 ? 1 : -1;
  let remaining = Math.abs(businessDays);

  if (remaining > BUSINESS_DAYS_LIMIT) {
    throw new Error("Počet pracovních dní je mimo povolený rozsah.");
  }

  const current = new Date(startDate);
  while (remaining > 0) {
    current.setDate(current.getDate() + direction);

    if (isBusinessDay(current, holidaySet)) {
      remaining -= 1;
    }
  }

  return current;
}
