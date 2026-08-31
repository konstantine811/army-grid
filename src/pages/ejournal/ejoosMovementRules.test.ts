import { describe, expect, it } from "vitest";
import { findLatestPriorOwnUnitStaffMove } from "./ejoosMovementRules";

const staffMove = (excelRow: number, orderDate: string) => ({
  type: "ПОСАДА",
  rank: "солдат",
  destination: "1ПБ",
  changeText: "2103161 → 2103222",
  status: "В СТРОЮ",
  note: "",
  previousIndex: "2103161",
  nextIndex: "2103222",
  orderDate,
  basisDate: orderDate,
  excelRow,
});

describe("findLatestPriorOwnUnitStaffMove", () => {
  it("keeps the earlier same-day РУХ as prior (tie-break by excelRow)", () => {
    const first = staffMove(10, "10.08.2026");
    const second = staffMove(11, "10.08.2026");
    const found = findLatestPriorOwnUnitStaffMove([first, second], second, {
      samePerson: () => true,
      inWindow: () => true,
      eventTime: () => Date.UTC(2026, 7, 10),
    });
    expect(found?.excelRow).toBe(10);
  });
});
