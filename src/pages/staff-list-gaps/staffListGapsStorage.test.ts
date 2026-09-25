import { describe, expect, it } from "vitest";
import { buildStaffListGapsSession } from "./staffListGapsStorage";
import type { StaffListSnapshot } from "./staffListGapsParse";

describe("staffListGapsStorage", () => {
  it("stores edited rows in snapshot", () => {
    const snapshot: StaffListSnapshot = {
      fileName: "phones.xlsx",
      sheetName: "Sheet1",
      sheetIndex: 0,
      columnIndexes: [0],
      columns: [{ id: "pib_0", label: "ПІБ", index: 0 }],
      rows: [
        {
          __rowId: "r1",
          __rowNumber: 1,
          __excelRowNumber: 2,
          pib: "ГАЛУШКО Олексій Олександрович",
          birthDate: "",
          values: { pib_0: "ГАЛУШКО Олексій Олександрович", phone_1: "" },
        },
      ],
      pibColumnId: "pib_0",
      birthDateColumnId: "",
    };
    const edited = [
      {
        ...snapshot.rows[0]!,
        values: {
          ...snapshot.rows[0]!.values,
          phone_1: "0501234567",
        },
      },
    ];
    const session = buildStaffListGapsSession(
      snapshot,
      edited,
      ["phone_1"],
      new ArrayBuffer(8),
    );
    expect(session.snapshot.rows[0]?.values.phone_1).toBe("0501234567");
    expect(session.gapColumnIds).toEqual(["phone_1"]);
  });
});
