const CZ_FORMATTER = new Intl.NumberFormat("cs-CZ", {
  style: "currency",
  currency: "CZK",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

export function parseCzechNumeric(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value !== "string") {
    return null;
  }

  const normalized = value.trim().replace(/\s+/g, "").replace(",", ".");

  if (!normalized) {
    return null;
  }

  const parsed = Number(normalized);
  if (Number.isNaN(parsed)) {
    return null;
  }

  return parsed;
}

export function formatCzk(value: number): string {
  return CZ_FORMATTER.format(value);
}
