import { describe, expect, it } from "vitest";
import {
  isContractMovementType,
  motivationContractOverlapsWindow,
  parseContractDatesFromChangeText,
} from "./ejoosSyncPlan";

describe("contract movement parsing", () => {
  it("recognizes regular and motivation contract event types", () => {
    expect(isContractMovementType("КОНТРАКТ")).toBe(true);
    expect(isContractMovementType("МОТИВАЦ КОНТР")).toBe(true);
    expect(isContractMovementType("МОТИВАЦ")).toBe(true);
    expect(isContractMovementType("ПОСАДА")).toBe(false);
  });

  it("extracts motivation contract start and end dates", () => {
    expect(
      parseContractDatesFromChangeText(
        "на 10 (десять) місяців з 09.07.2026 до 09.05.2027",
      ),
    ).toEqual({
      contractFrom: "09.07.2026",
      contractTo: "09.05.2027",
    });
  });

  it("treats July contract start as in-scope for August snapshot", () => {
    const augustStart = Date.UTC(2026, 7, 1);
    const augustEnd = Date.UTC(2026, 7, 25);
    expect(
      motivationContractOverlapsWindow(
        {
          changeText: "на 10 (десять) місяців з 09.07.2026 до 09.05.2027",
          basisDate: "09.07.2026",
          orderDate: "09.07.2026",
        },
        augustStart,
        augustEnd,
      ),
    ).toBe(true);
  });

  it("normalizes slash-separated dates and a two-digit year", () => {
    expect(
      parseContractDatesFromChangeText(
        "на 2 роки із 9/7/26 до 9/7/2028.",
      ),
    ).toEqual({
      contractFrom: "09.07.2026",
      contractTo: "09.07.2028",
    });
  });
});
