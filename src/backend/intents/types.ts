export enum IntentType {
  VatAdd = "vat.add",
  FormatCzk = "format.currency",
  FetchCnbRate = "cnb.fetch_rate",
  FxConvertCnb = "cnb.fx_convert",
  FinanceDedupe = "finance.dedupe",
  SortColumn = "sheet.sort_column",
  VatRemove = "vat.remove",
  HighlightNegative = "sheet.highlight_negative",
  SumColumn = "sheet.sum_column",
  MonthlyRunRate = "analysis.monthly_runrate",
  PeriodSummary = "analysis.period_summary",
  RollingWindow = "analysis.rolling_window",
  VarianceVsBudget = "analysis.variance_vs_budget",
  PeriodComparison = "analysis.period_comparison",
  SeedHolidays = "holidays.seed",
  NetworkdaysDue = "schedule.networkdays_due"
}

export interface IntentBase {
  type: IntentType;
  originalText: string;
  confidence: number;
}

export interface VatAddIntent extends IntentBase {
  type: IntentType.VatAdd;
  rate: number;
  rateLabel: string;
  columnLetter?: string;
}

export interface FormatCzkIntent extends IntentBase {
  type: IntentType.FormatCzk;
  columnLetter?: string;
}

export interface FetchCnbRateIntent extends IntentBase {
  type: IntentType.FetchCnbRate;
  currency: string;
  targetDate: string; // ISO date
  source: "cache" | "api" | "auto";
}

export interface FxConvertCnbIntent extends IntentBase {
  type: IntentType.FxConvertCnb;
  currency: string;
  targetDate: string;
  columnLetter?: string;
}

export interface FinanceDedupeIntent extends IntentBase {
  type: IntentType.FinanceDedupe;
  columnLetter?: string;
}

export interface SortColumnIntent extends IntentBase {
  type: IntentType.SortColumn;
  columnLetter?: string;
  direction: "asc" | "desc";
}

export interface VatRemoveIntent extends IntentBase {
  type: IntentType.VatRemove;
  rate: number;
  rateLabel: string;
  columnLetter?: string;
}

export interface HighlightNegativeIntent extends IntentBase {
  type: IntentType.HighlightNegative;
  columnLetter?: string;
}

export interface SumColumnIntent extends IntentBase {
  type: IntentType.SumColumn;
  columnLetter?: string;
}

export interface MonthlyRunRateIntent extends IntentBase {
  type: IntentType.MonthlyRunRate;
  amountColumn?: string;
  dateColumn?: string;
  months: number;
}

export interface PeriodComparisonIntent extends IntentBase {
  type: IntentType.PeriodComparison;
  amountColumn?: string;
  dateColumn?: string;
}

export interface PeriodSummaryIntent extends IntentBase {
  type: IntentType.PeriodSummary;
  amountColumn?: string;
  dateColumn?: string;
}

export interface RollingWindowIntent extends IntentBase {
  type: IntentType.RollingWindow;
  amountColumn?: string;
  dateColumn?: string;
  windowSize: number;
  aggregation: "sum" | "avg";
}

export interface VarianceVsBudgetIntent extends IntentBase {
  type: IntentType.VarianceVsBudget;
  actualColumn?: string;
  budgetColumn?: string;
  dateColumn?: string;
}

export interface SeedHolidaysIntent extends IntentBase {
  type: IntentType.SeedHolidays;
  year: number;
}

export interface NetworkdaysDueIntent extends IntentBase {
  type: IntentType.NetworkdaysDue;
  businessDays: number;
  startDate?: string;
}

export type SupportedIntent =
  | VatAddIntent
  | FormatCzkIntent
  | FetchCnbRateIntent
  | FxConvertCnbIntent
  | FinanceDedupeIntent
  | SortColumnIntent
  | VatRemoveIntent
  | HighlightNegativeIntent
  | SumColumnIntent
  | MonthlyRunRateIntent
  | VarianceVsBudgetIntent
  | PeriodSummaryIntent
  | RollingWindowIntent
  | PeriodComparisonIntent
  | SeedHolidaysIntent
  | NetworkdaysDueIntent;

export interface ParsedIntentOutcome {
  intent: SupportedIntent;
  issues: string[];
}

export interface SampleTable {
  headers: string[];
  rows: string[][];
}

export interface VatAddApplyPayload {
  sheetName: string;
  rowIndex: number;
  rowCount: number;
  columnIndex: number;
  rate: number;
  rateLabel: string;
  hasHeader: boolean;
}

export interface FormatCzkApplyPayload {
  sheetName: string;
  rowIndex: number;
  rowCount: number;
  columnIndex: number;
}

export interface FetchCnbRateApplyPayload {
  currency: string;
  targetDate: string;
}

export interface FxConvertCnbApplyPayload {
  sheetName: string;
  rowIndex: number;
  rowCount: number;
  columnIndex: number;
  currency: string;
  targetDate: string;
  hasHeader: boolean;
}

export interface SeedHolidaysApplyPayload {
  year: number;
}

export interface NetworkdaysDueApplyPayload {
  sheetName: string;
  startCell: string;
  businessDays: number;
  startDateISO: string;
}

export interface FinanceDedupeApplyPayload {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  hasHeader: boolean;
}

export interface SortColumnApplyPayload {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  hasHeader: boolean;
  ascending: boolean;
}

export interface VatRemoveApplyPayload {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  hasHeader: boolean;
  rate: number;
  rateLabel: string;
}

export interface HighlightNegativeApplyPayload {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  hasHeader: boolean;
}

export interface SumColumnApplyPayload {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  hasHeader: boolean;
}

export interface MonthlyRunRateApplyPayload {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  hasHeader: boolean;
  amountColumnLetter: string;
  dateColumnLetter: string;
  months: number;
}

export interface PeriodComparisonApplyPayload {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  hasHeader: boolean;
  amountColumnLetter: string;
  dateColumnLetter: string;
}

export interface PeriodSummaryApplyPayload {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  hasHeader: boolean;
  amountColumnLetter: string;
  dateColumnLetter: string;
}

export interface RollingWindowApplyPayload {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  hasHeader: boolean;
  amountColumnLetter: string;
  dateColumnLetter: string;
  windowSize: number;
  aggregation: "sum" | "avg";
}

export interface VarianceVsBudgetApplyPayload {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  hasHeader: boolean;
  actualColumnLetter: string;
  budgetColumnLetter: string;
  dateColumnLetter: string;
}

export type ApplyPayload =
  | VatAddApplyPayload
  | FormatCzkApplyPayload
  | FetchCnbRateApplyPayload
  | FxConvertCnbApplyPayload
  | FinanceDedupeApplyPayload
  | SortColumnApplyPayload
  | VatRemoveApplyPayload
  | HighlightNegativeApplyPayload
  | SumColumnApplyPayload
  | MonthlyRunRateApplyPayload
  | PeriodSummaryApplyPayload
  | RollingWindowApplyPayload
  | VarianceVsBudgetApplyPayload
  | PeriodComparisonApplyPayload
  | SeedHolidaysApplyPayload
  | NetworkdaysDueApplyPayload;

export interface IntentPreview<TPayload extends ApplyPayload = ApplyPayload> {
  intent: SupportedIntent;
  planText: string;
  sample?: SampleTable;
  issues: string[];
  applyPayload: TPayload;
}

export interface FailedPreview {
  error: string;
  issues?: string[];
}

export type PreviewResult<TPayload extends ApplyPayload = ApplyPayload> = IntentPreview<TPayload> | FailedPreview;

export function isFailedPreview(value: PreviewResult): value is FailedPreview {
  return (value as FailedPreview).error !== undefined;
}

export interface ApplyResult {
  message: string;
  warnings?: string[];
}
