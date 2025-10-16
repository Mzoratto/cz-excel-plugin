import { normalizeCzechText } from "./text";

export function formatISODate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function createDateOrNull(year: number, month: number, day: number): Date | null {
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  ) {
    return new Date(year, month - 1, day);
  }
  return null;
}

export function parseCzechDateExpression(source: string): Date | null {
  const normalized = normalizeCzechText(source);

  if (normalized.includes("dnes")) {
    return new Date();
  }

  if (normalized.includes("zit") || normalized.includes("zítr")) {
    const date = new Date();
    date.setDate(date.getDate() + 1);
    return date;
  }

  const isoMatch = normalized.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);
    return createDateOrNull(year, month, day);
  }

  const dottedMatch = normalized.match(/(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{2,4})?/);
  if (dottedMatch) {
    const day = Number(dottedMatch[1]);
    const month = Number(dottedMatch[2]);
    const yearToken = dottedMatch[3];
    const currentYear = new Date().getFullYear();
    const year =
      yearToken?.length === 2 ? Number(`20${yearToken}`) : Number(yearToken ?? `${currentYear}`);
    return createDateOrNull(year, month, day);
  }

  return null;
}
