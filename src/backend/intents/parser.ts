import {
  IntentType,
  ParsedIntentOutcome,
  SupportedIntent,
  FetchCnbRateIntent,
  FxConvertCnbIntent,
  FinanceDedupeIntent,
  SortColumnIntent,
  VatRemoveIntent,
  HighlightNegativeIntent,
  SumColumnIntent,
  MonthlyRunRateIntent,
  PeriodSummaryIntent,
  RollingWindowIntent,
  VarianceVsBudgetIntent,
  PeriodComparisonIntent,
  SeedHolidaysIntent,
  NetworkdaysDueIntent
} from "./types";
import { normalizeCzechText } from "../utils/text";
import { formatISODate, parseCzechDateExpression } from "../utils/date";

const VAT_RATES: Record<string, number> = {
  "21": 0.21,
  "15": 0.15,
  "12": 0.12,
  "10": 0.1
};

const VAT_ALIASES: Record<string, string> = {
  "21": "21 %",
  "15": "15 %",
  "12": "12 %",
  "10": "10 %"
};

const CZK_KEYWORDS = ["czk", "korun", "korunu", "koruny", "korunách", "korunach", "kč"];
const DEDUPE_KEYWORDS = ["duplic", "duplik", "dedupe", "duplicit", "duplikat"];
const SORT_KEYWORDS = ["serad", "seřaď", "sort", "seřadit", "usporadej", "uspořádej"];
const SORT_ASC_KEYWORDS = ["vzestup", "ascending", "nahoru", "vzestupne", "vzestupně"];
const SORT_DESC_KEYWORDS = ["sestup", "descending", "dolu", "dolů", "sestupne", "sestupně"];
const VAT_REMOVE_KEYWORDS = ["bez dph", "odeber dph", "odstran dph", "reverse charge", "vycisti dph", "bez dane"];
const RUNRATE_KEYWORDS = ["run-rate", "runrate", "run rate"];
const PERIOD_COMPARISON_KEYWORDS = ["mezimesic", "meziměs", "mom", "qoq", "yoy", "meziroční", "meziměsíční", "čtvrtletní", "quarter", "year over year"];
const PERIOD_SUMMARY_KEYWORDS = ["ytd", "mtd", "qtd", "year to date", "month to date", "quarter to date", "souhrn z", "aktualni rok"];
const ROLLING_WINDOW_KEYWORDS = ["rolling", "klouzav", "rolling window", "rolling 12", "rolling 6"];
const VARIANCE_KEYWORDS = ["odchylka", "variance", "vs budget", "rozpočet", "rozpocet", "skutečnost", "skutecnost"];
const HIGHLIGHT_NEGATIVE_KEYWORDS = ["zvyrazni zaporna", "zvýrazni záporná", "highlight negative", "zvyrazni minus", "obarvi zaporne"];
const SUM_KEYWORDS = ["součet", "soucet", "sumuj", "sum", "souhrn", "total"];

const COLUMN_PATTERN =
  /\bsloup(?:ec|ce|ci)\s+([a-záčďéěíňóřšťúůýžA-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ]{1,3}|\w?\d+)\b|\bcolumn\s+([a-zA-Z])\b/iu;

const CURRENCY_PATTERN = /\b([A-Z]{3})\b/;
const SUPPORTED_CURRENCIES = new Set([
  "AUD",
  "BGN",
  "BRL",
  "CAD",
  "CHF",
  "CNY",
  "DKK",
  "EUR",
  "GBP",
  "HKD",
  "HUF",
  "IDR",
  "ILS",
  "INR",
  "ISK",
  "JPY",
  "KRW",
  "MXN",
  "MYR",
  "NOK",
  "NZD",
  "PHP",
  "PLN",
  "RON",
  "SEK",
  "SGD",
  "THB",
  "TRY",
  "USD",
  "XDR",
  "ZAR"
]);

const NUMBER_PATTERN = /(-?\d+)/;

function extractColumnLetter(source: string): string | undefined {
  const match = source.match(COLUMN_PATTERN);
  if (!match) {
    return undefined;
  }

  const candidate = (match[1] ?? match[2] ?? match[3] ?? "").trim();
  if (!candidate) {
    return undefined;
  }

  const letter = candidate[0];
  if (!/[a-z]/i.test(letter)) {
    return undefined;
  }

  return letter.toUpperCase();
}

function extractColumnRoles(source: string): Array<{ letter: string; label: string }> {
  const results: Array<{ letter: string; label: string }> = [];
  const pattern = /([A-Z])\s*\(([^)]+)\)/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(source.toUpperCase())) !== null) {
    results.push({ letter: match[1]!, label: match[2]! });
  }
  return results;
}

function detectVatIntent(originalText: string, normalized: string): SupportedIntent | undefined {
  if (!/\b(dph|vat)\b/.test(normalized)) {
    return undefined;
  }

  if (VAT_REMOVE_KEYWORDS.some((keyword) => normalized.includes(keyword))) {
    return undefined;
  }

  const rateMatch = normalized.match(/(\d{1,2})(?:\s*%|\s*procent|)/);
  if (!rateMatch) {
    return undefined;
  }

  const rateKey = rateMatch[1];
  const rate = VAT_RATES[rateKey];
  if (typeof rate !== "number") {
    return undefined;
  }

  const columnLetter = extractColumnLetter(originalText);

  return {
    type: IntentType.VatAdd,
    rate,
    rateLabel: VAT_ALIASES[rateKey] ?? `${rateKey} %`,
    columnLetter,
    originalText,
    confidence: columnLetter ? 0.95 : 0.85
  };
}

function detectMonthlyRunRateIntent(originalText: string, normalized: string): MonthlyRunRateIntent | undefined {
  const mentionsRunRate = RUNRATE_KEYWORDS.some((keyword) => normalized.includes(keyword));
  if (!mentionsRunRate) {
    return undefined;
  }

  const monthsMatch = normalized.match(/(\d+)\s*(?:měsíc|mesic|m)/);
  const months = monthsMatch ? Math.max(1, parseInt(monthsMatch[1]!, 10)) : 3;

  const roles = extractColumnRoles(originalText);
  const dateColumn = roles.find((entry) => /datum|date/.test(entry.label.toLowerCase()))?.letter;
  const amountColumn = roles.find((entry) => /část|cast|cena|hodnot|amount|tržb|trzb/.test(entry.label.toLowerCase()))?.letter;

  return {
    type: IntentType.MonthlyRunRate,
    amountColumn,
    dateColumn,
    months,
    originalText,
    confidence: amountColumn && dateColumn ? 0.9 : 0.6
  };
}

function detectPeriodSummaryIntent(originalText: string, normalized: string): PeriodSummaryIntent | undefined {
  const mentionsSummary = PERIOD_SUMMARY_KEYWORDS.some((keyword) => normalized.includes(keyword));
  if (!mentionsSummary) {
    return undefined;
  }

  const roles = extractColumnRoles(originalText);
  const dateColumn = roles.find((entry) => /datum|date/.test(entry.label.toLowerCase()))?.letter;
  const amountColumn = roles.find((entry) => /část|cast|cena|hodnot|amount|tržb|trzb/.test(entry.label.toLowerCase()))?.letter;

  return {
    type: IntentType.PeriodSummary,
    amountColumn,
    dateColumn,
    originalText,
    confidence: amountColumn && dateColumn ? 0.9 : 0.65
  };
}

function detectPeriodComparisonIntent(originalText: string, normalized: string): PeriodComparisonIntent | undefined {
  const mentionsComparison = PERIOD_COMPARISON_KEYWORDS.some((keyword) => normalized.includes(keyword));
  if (!mentionsComparison) {
    return undefined;
  }

  const roles = extractColumnRoles(originalText);
  const dateColumn = roles.find((entry) => /datum|date/.test(entry.label.toLowerCase()))?.letter;
  const amountColumn = roles.find((entry) => /část|cast|cena|hodnot|amount|tržb|trzb/.test(entry.label.toLowerCase()))?.letter;

  return {
    type: IntentType.PeriodComparison,
    amountColumn,
    dateColumn,
    originalText,
    confidence: amountColumn && dateColumn ? 0.9 : 0.65
  };
}

function detectFormatIntent(originalText: string, normalized: string): SupportedIntent | undefined {
  const hasFormatKeyword = normalized.includes("format") || normalized.includes("formát");
  const hasCzkKeyword = CZK_KEYWORDS.some((keyword) => normalized.includes(keyword));

  if (!(hasFormatKeyword && hasCzkKeyword)) {
    return undefined;
  }

  const columnLetter = extractColumnLetter(originalText);

  return {
    type: IntentType.FormatCzk,
    columnLetter,
    originalText,
    confidence: columnLetter ? 0.9 : 0.8
  };
}

function detectDedupeIntent(originalText: string, normalized: string): FinanceDedupeIntent | undefined {
  const mentionsDedupe = DEDUPE_KEYWORDS.some((keyword) => normalized.includes(keyword));
  if (!mentionsDedupe) {
    return undefined;
  }

  const columnLetter = extractColumnLetter(originalText);

  return {
    type: IntentType.FinanceDedupe,
    columnLetter,
    originalText,
    confidence: columnLetter ? 0.85 : 0.75
  };
}

function detectSortIntent(originalText: string, normalized: string): SortColumnIntent | undefined {
  const mentionsSort = SORT_KEYWORDS.some((keyword) => normalized.includes(keyword));
  if (!mentionsSort) {
    return undefined;
  }

  const columnLetter = extractColumnLetter(originalText);
  const mentionsAsc = SORT_ASC_KEYWORDS.some((keyword) => normalized.includes(keyword));
  const mentionsDesc = SORT_DESC_KEYWORDS.some((keyword) => normalized.includes(keyword));
  const direction: "asc" | "desc" = mentionsDesc ? "desc" : mentionsAsc ? "asc" : "asc";

  return {
    type: IntentType.SortColumn,
    columnLetter,
    direction,
    originalText,
    confidence: columnLetter ? 0.85 : 0.7
  };
}

function detectVatRemoveIntent(originalText: string, normalized: string): VatRemoveIntent | undefined {
  const mentionsRemoval = VAT_REMOVE_KEYWORDS.some((keyword) => normalized.includes(keyword));
  if (!mentionsRemoval) {
    return undefined;
  }

  const rateMatch = normalized.match(/(\d{1,2})(?:\s*%|\s*procent|\s*dph)/);
  const rateKey = rateMatch?.[1] ?? "21";
  const rateValue = VAT_RATES[rateKey] ?? VAT_RATES["21"];
  const rateLabel = VAT_ALIASES[rateKey] ?? `${rateKey} %`;
  const columnLetter = extractColumnLetter(originalText);

  return {
    type: IntentType.VatRemove,
    rate: rateValue,
    rateLabel,
    columnLetter,
    originalText,
    confidence: columnLetter ? 0.85 : 0.7
  };
}

function detectHighlightNegativeIntent(originalText: string, normalized: string): HighlightNegativeIntent | undefined {
  const mentionsHighlight = HIGHLIGHT_NEGATIVE_KEYWORDS.some((keyword) => normalized.includes(keyword));
  if (!mentionsHighlight) {
    return undefined;
  }

  const columnLetter = extractColumnLetter(originalText);

  return {
    type: IntentType.HighlightNegative,
    columnLetter,
    originalText,
    confidence: columnLetter ? 0.85 : 0.75
  };
}

function detectSumIntent(originalText: string, normalized: string): SumColumnIntent | undefined {
  const mentionsSum = SUM_KEYWORDS.some((keyword) => normalized.includes(keyword));
  if (!mentionsSum) {
    return undefined;
  }

  const columnLetter = extractColumnLetter(originalText);

  return {
    type: IntentType.SumColumn,
    columnLetter,
    originalText,
    confidence: columnLetter ? 0.85 : 0.75
  };
}

function detectFetchCnbRateIntent(originalText: string, normalized: string): FetchCnbRateIntent | undefined {
  if (!normalized.includes("kurz") || !normalized.includes("cnb")) {
    return undefined;
  }

  const currencyMatch = originalText.toUpperCase().match(CURRENCY_PATTERN);
  const currency = currencyMatch?.[1];
  if (!currency || !SUPPORTED_CURRENCIES.has(currency)) {
    return undefined;
  }

  const date = parseCzechDateExpression(originalText);
  const targetDate = formatISODate(date ?? new Date());

  return {
    type: IntentType.FetchCnbRate,
    currency,
    targetDate,
    source: "auto",
    originalText,
    confidence: 0.85
  };
}

function detectFxConvertIntent(originalText: string, normalized: string): FxConvertCnbIntent | undefined {
  const conversionKeywords = ["preved", "převeď", "přepoč", "prepoc"];
  const hasConvertKeyword = conversionKeywords.some((keyword) => normalized.includes(keyword));
  const mentionsCzk = CZK_KEYWORDS.some((keyword) => normalized.includes(keyword));
  const mentionsCnb = normalized.includes("cnb");

  if (!(hasConvertKeyword && mentionsCzk && mentionsCnb)) {
    return undefined;
  }

  const currencyMatch = originalText.toUpperCase().match(CURRENCY_PATTERN);
  const currency = currencyMatch?.[1];
  if (!currency || currency === "CZK" || !SUPPORTED_CURRENCIES.has(currency)) {
    return undefined;
  }

  const date = parseCzechDateExpression(originalText);
  const targetDate = formatISODate(date ?? new Date());
  const columnLetter = extractColumnLetter(originalText);

  return {
    type: IntentType.FxConvertCnb,
    currency,
    targetDate,
    columnLetter,
    originalText,
    confidence: columnLetter ? 0.9 : 0.8
  };
}

function determineWindowSize(normalized: string, fallback: number): number {
  const match = normalized.match(/(\d+)\s*(?:mesi|měsí|month|rolling|period|window)/);
  if (match) {
    const value = parseInt(match[1]!, 10);
    if (Number.isFinite(value) && value > 0 && value <= 120) {
      return value;
    }
  }
  return fallback;
}

function detectRollingWindowIntent(originalText: string, normalized: string): RollingWindowIntent | undefined {
  const mentionsRolling = ROLLING_WINDOW_KEYWORDS.some((keyword) => normalized.includes(keyword));
  if (!mentionsRolling) {
    return undefined;
  }

  const roles = extractColumnRoles(originalText);
  const dateColumn = roles.find((entry) => /datum|date/.test(entry.label.toLowerCase()))?.letter;
  const amountColumn = roles.find((entry) => /část|cast|cena|hodnot|amount|tržb|trzb/.test(entry.label.toLowerCase()))?.letter;
  const aggregation = normalized.includes("avg") || normalized.includes("prům") ? "avg" : "sum";
  const windowSize = determineWindowSize(normalized, 12);

  return {
    type: IntentType.RollingWindow,
    amountColumn,
    dateColumn,
    windowSize,
    aggregation,
    originalText,
    confidence: amountColumn && dateColumn ? 0.85 : 0.65
  };
}

function detectSeedHolidaysIntent(originalText: string, normalized: string): SeedHolidaysIntent | undefined {
  if (!normalized.includes("svatk") && !normalized.includes("svátk")) {
    return undefined;
  }

  const yearMatch = originalText.match(/\b(20\d{2})\b/);
  const year = yearMatch ? Number(yearMatch[1]) : new Date().getFullYear();

  if (year < 2000 || year > 2100) {
    return undefined;
  }

  return {
    type: IntentType.SeedHolidays,
    year,
    originalText,
    confidence: 0.9
  };
}

function detectNetworkdaysIntent(originalText: string, normalized: string): NetworkdaysDueIntent | undefined {
  if (!normalized.includes("pracovn") && !normalized.includes("business") && !normalized.includes("sla")) {
    return undefined;
  }

  const numberMatch = normalized.match(NUMBER_PATTERN);
  if (!numberMatch) {
    return undefined;
  }

  const businessDays = parseInt(numberMatch[1], 10);
  if (!Number.isFinite(businessDays) || businessDays === 0 || Math.abs(businessDays) > 365) {
    return undefined;
  }

  const startDate = parseCzechDateExpression(originalText);

  return {
    type: IntentType.NetworkdaysDue,
    businessDays,
    startDate: startDate ? formatISODate(startDate) : undefined,
    originalText,
    confidence: 0.8
  };
}

export function parseCzechRequest(text: string): ParsedIntentOutcome | null {
  const trimmed = text.trim();
  if (!trimmed) {
    return null;
  }

  const normalized = normalizeCzechText(trimmed);

  const detectors: Array<(original: string, normalized: string) => SupportedIntent | undefined> = [
    detectVatIntent,
    detectFormatIntent,
    detectDedupeIntent,
    detectVatRemoveIntent,
    detectSortIntent,
    detectHighlightNegativeIntent,
    detectPeriodSummaryIntent,
    detectRollingWindowIntent,
    detectSumIntent,
    detectMonthlyRunRateIntent,
    detectVarianceVsBudgetIntent,
    detectPeriodComparisonIntent,
    detectFxConvertIntent,
    detectFetchCnbRateIntent,
    detectSeedHolidaysIntent,
    detectNetworkdaysIntent
  ];

  for (const detector of detectors) {
    const intent = detector(trimmed, normalized);
    if (intent) {
      return { intent, issues: [] };
    }
  }

  return null;
}
function detectVarianceVsBudgetIntent(originalText: string, normalized: string): VarianceVsBudgetIntent | undefined {
  const mentionsVariance = VARIANCE_KEYWORDS.some((keyword) => normalized.includes(keyword));
  if (!mentionsVariance) {
    return undefined;
  }

  const roles = extractColumnRoles(originalText);
  const actualColumn = roles.find((entry) => /skute|actual/.test(entry.label.toLowerCase()))?.letter;
  const budgetColumn = roles.find((entry) => /plán|plan|budget/.test(entry.label.toLowerCase()))?.letter;
  const dateColumn = roles.find((entry) => /datum|date/.test(entry.label.toLowerCase()))?.letter;

  return {
    type: IntentType.VarianceVsBudget,
    actualColumn,
    budgetColumn,
    dateColumn,
    originalText,
    confidence: actualColumn && budgetColumn ? 0.85 : 0.6
  };
}
