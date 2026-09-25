import { describe, expect, it } from "vitest";
import {
  mapCollectedVKPersonToExportRow,
  mapCollectedVKSheetsToExportRows,
} from "./collectVKExcelExport";

describe("collectVKExcelExport", () => {
  it("maps collected person into flat export row", () => {
    const row = mapCollectedVKPersonToExportRow("ІВАНОВ Іван", {
      staff: {
        alias: "БАЛЯ",
        status: "активний",
        unit: "1 рота",
        countUpdate: 0,
      },
      ejoos: { cardNumber: "АГ 123" },
      ksp: {
        cardType: "ВК",
        soldierStatus: "у строю",
        soldierOperation: "операція",
        note: "примітка",
      },
      militaryCards: {
        ВК: {
          status: true,
          cardNumber: "АГ 123",
          alias: "БАЛЯ",
          whereToGet: "",
        },
      },
    });

    expect(row.pib).toBe("ІВАНОВ Іван");
    expect(row.staffAlias).toBe("БАЛЯ");
    expect(row.ejoosCardNumber).toBe("АГ 123");
    expect(row.kspCardType).toBe("ВК");
    expect(row.vkStatus).toBe("Є");
    expect(row.vkCardNumber).toBe("АГ 123");
  });

  it("builds rows from sheets record", () => {
    const rows = mapCollectedVKSheetsToExportRows({
      "ІВАНОВ Іван": {
        staff: {
          alias: "",
          status: "активний",
          unit: "1 рота",
          countUpdate: 0,
        },
        ejoos: undefined,
        militaryCards: undefined,
        ksp: undefined,
      },
    });

    expect(rows).toHaveLength(1);
    expect(rows[0]?.staffStatus).toBe("активний");
  });
});
