import { describe, expect, it } from "vitest";
import { parseCzechRequest } from "../src/backend/intents/parser";
import { IntentType } from "../src/backend/intents/types";

describe("parseCzechRequest finance intents", () => {
  it("recognizes add_vat phrasing as VAT intent", () => {
    const outcome = parseCzechRequest("add_vat 21 % do sloupce C");
    expect(outcome).not.toBeNull();
    if (!outcome) {
      return;
    }
    expect(outcome.intent.type).toBe(IntentType.VatAdd);
    expect("columnLetter" in outcome.intent ? outcome.intent.columnLetter : undefined).toBe("C");
  });

  it("recognizes Czech dedupe request", () => {
    const outcome = parseCzechRequest("Odeber duplicity ve sloupci B");
    expect(outcome).not.toBeNull();
    if (!outcome) {
      return;
    }
    expect(outcome.intent.type).toBe(IntentType.FinanceDedupe);
    expect("columnLetter" in outcome.intent ? outcome.intent.columnLetter : undefined).toBe("B");
  });

  it("handles English dedupe keyword", () => {
    const outcome = parseCzechRequest("dedupe column D");
    expect(outcome).not.toBeNull();
    if (!outcome) {
      return;
    }
    expect(outcome.intent.type).toBe(IntentType.FinanceDedupe);
    expect("columnLetter" in outcome.intent ? outcome.intent.columnLetter : undefined).toBe("D");
  });

  it("detects sort ascending intent", () => {
    const outcome = parseCzechRequest("Seřaď vzestupně sloupec B");
    expect(outcome).not.toBeNull();
    if (!outcome) {
      return;
    }
    expect(outcome.intent.type).toBe(IntentType.SortColumn);
    if (outcome.intent.type === IntentType.SortColumn) {
      expect(outcome.intent.columnLetter).toBe("B");
      expect(outcome.intent.direction).toBe("asc");
    }
  });

  it("detects vat remove intent", () => {
    const outcome = parseCzechRequest("Odeber DPH 21 % ze sloupce C");
    expect(outcome).not.toBeNull();
    if (!outcome) {
      return;
    }
    expect(outcome.intent.type).toBe(IntentType.VatRemove);
    if (outcome.intent.type === IntentType.VatRemove) {
      expect(outcome.intent.columnLetter).toBe("C");
      expect(outcome.intent.rate).toBeCloseTo(0.21);
    }
  });

  it("detects highlight negative intent", () => {
    const outcome = parseCzechRequest("Zvýrazni záporná čísla ve sloupci D");
    expect(outcome).not.toBeNull();
    if (!outcome) {
      return;
    }
    expect(outcome.intent.type).toBe(IntentType.HighlightNegative);
    if (outcome.intent.type === IntentType.HighlightNegative) {
      expect(outcome.intent.columnLetter).toBe("D");
    }
  });

  it("detects sum column intent", () => {
    const outcome = parseCzechRequest("Vypočítej součet ve sloupci E");
    expect(outcome).not.toBeNull();
    if (!outcome) {
      return;
    }
    expect(outcome.intent.type).toBe(IntentType.SumColumn);
    if (outcome.intent.type === IntentType.SumColumn) {
      expect(outcome.intent.columnLetter).toBe("E");
    }
  });

  it("detects monthly run-rate intent", () => {
    const outcome = parseCzechRequest("Run-rate z posledních 3 měsíců podle B (částky) a A (datum)");
    expect(outcome).not.toBeNull();
    if (!outcome) {
      return;
    }
    expect(outcome.intent.type).toBe(IntentType.MonthlyRunRate);
    if (outcome.intent.type === IntentType.MonthlyRunRate) {
      expect(outcome.intent.amountColumn).toBe("B");
      expect(outcome.intent.dateColumn).toBe("A");
      expect(outcome.intent.months).toBe(3);
    }
  });
});
