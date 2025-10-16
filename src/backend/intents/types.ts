export enum IntentType {
  VatAdd = "vat.add",
  FormatCzk = "format.currency",
  FetchCnbRate = "cnb.fetch_rate",
  FxConvertCnb = "cnb.fx_convert",
  FinanceDedupe = "finance.dedupe",
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

export type ApplyPayload =
  | VatAddApplyPayload
  | FormatCzkApplyPayload
  | FetchCnbRateApplyPayload
  | FxConvertCnbApplyPayload
  | FinanceDedupeApplyPayload
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
