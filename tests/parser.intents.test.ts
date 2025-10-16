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
});
