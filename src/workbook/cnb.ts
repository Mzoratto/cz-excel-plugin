export interface RateResult {
  rate: number;
  source: "cache" | "api";
}

const FX_TABLE_NAME = "tblFxCnb";

function normalizeCurrency(code: string): string {
  return code.trim().toUpperCase();
}

function extractCurrencyCode(entry: Record<string, unknown>): string {
  const candidateKeys = ["currencyCode", "CurrencyCode", "code", "Code"];
  for (const key of candidateKeys) {
    const raw = entry[key];
    if (typeof raw === "string" && raw.trim().length > 0) {
      return raw;
    }
  }

  const fallback = entry["currency"];
  if (typeof fallback === "string" && fallback.trim().length === 3) {
    return fallback;
  }

  return "";
}

export async function getCachedRate(
  context: Excel.RequestContext,
  currency: string,
  targetDate: string
): Promise<number | null> {
  const table = context.workbook.tables.getItem(FX_TABLE_NAME);
  const range = table.getDataBodyRangeOrNullObject();
  range.load(["values", "rowCount", "isNullObject"]);
  await context.sync();

  if (range.isNullObject || range.rowCount === 0) {
    return null;
  }

  for (const row of range.values) {
    const [dateCell, codeCell, rateCell] = row;
    if (`${dateCell}` === targetDate && normalizeCurrency(`${codeCell}`) === currency) {
      const parsed = Number(rateCell);
      if (!Number.isNaN(parsed)) {
        return parsed;
      }
    }
  }

  return null;
}

async function storeRate(
  context: Excel.RequestContext,
  currency: string,
  targetDate: string,
  rate: number
): Promise<void> {
  const table = context.workbook.tables.getItem(FX_TABLE_NAME);
  table.rows.add(undefined, [[targetDate, currency, rate]]);
}

async function fetchCnbRateFromApi(currency: string, targetDate: string): Promise<number> {
  const url = `https://api.cnb.cz/cnbapi/exrates/daily?date=${targetDate}&lang=EN`;

  const response = await fetch(url, {
    headers: {
      Accept: "application/json"
    }
  });

  if (!response.ok) {
    throw new Error(`ČNB API odpovědělo stavem ${response.status}`);
  }

  const payload = await response.json();
  const rates: Array<Record<string, unknown>> = payload?.rates ?? payload?.data ?? [];
  if (!Array.isArray(rates)) {
    throw new Error("Neočekávaný formát odpovědi ČNB.");
  }

  const upperCurrency = normalizeCurrency(currency);
  for (const entry of rates) {
    const extractedCode = extractCurrencyCode(entry);
    const code = normalizeCurrency(extractedCode);
    if (code === upperCurrency) {
      const amount = Number(entry?.amount ?? 1);
      const rateValue = Number(entry?.rate ?? entry?.mid ?? entry?.value);
      if (!Number.isFinite(rateValue) || !Number.isFinite(amount) || amount === 0) {
        throw new Error("ČNB odpověď neobsahovala platný kurz.");
      }
      return rateValue / amount;
    }
  }

  throw new Error(`Kurz ČNB pro ${upperCurrency} k ${targetDate} nebyl nalezen.`);
}

export async function ensureCnbRate(
  context: Excel.RequestContext,
  currencyCode: string,
  targetDate: string
): Promise<RateResult> {
  const currency = normalizeCurrency(currencyCode);

  const cached = await getCachedRate(context, currency, targetDate);
  if (cached !== null) {
    return { rate: cached, source: "cache" };
  }

  const fetched = await fetchCnbRateFromApi(currency, targetDate);
  await storeRate(context, currency, targetDate, fetched);
  return { rate: fetched, source: "api" };
}
