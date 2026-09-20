import { describe, expect, it } from "vitest";
import {
  buildExcelLabProcessedData,
  describeExcelLabProcessedData,
} from "./parseUploadedExcel";

describe("buildExcelLabProcessedData", () => {
  it("merges staff, military cards and ejoos uploads", () => {
    const processed = buildExcelLabProcessedData({
      staff: {
        fileName: "Штатка.xlsx",
        state: {
          "ІВАНОВ Іван": {
            alias: "БАЛЯ",
            status: "активний",
            unit: "1 рота",
            countUpdate: 1,
          },
        },
      },
      militaryCards: {
        fileName: "квитки.xlsx",
        cards: {
          "ІВАНОВ Іван": {
            ВК: {
              status: true,
              alias: "БАЛЯ",
              cardNumber: "АГ 123",
              whereToGet: "",
            },
          },
        },
      },
      ejoos: {
        fileName: "ЄЖООС.xlsx",
        state: {
          "ІВАНОВ Іван": { cardNumber: "АГ 123" },
        },
      },
    });

    expect(processed.staff["ІВАНОВ Іван"]?.alias).toBe("БАЛЯ");
    expect(processed.militaryCards["ІВАНОВ Іван"]?.ВК?.cardNumber).toBe(
      "АГ 123",
    );
    expect(processed.ejoos["ІВАНОВ Іван"]?.cardNumber).toBe("АГ 123");
    expect(describeExcelLabProcessedData(processed)).toContain("Завантажено: 3/3");
  });
});
