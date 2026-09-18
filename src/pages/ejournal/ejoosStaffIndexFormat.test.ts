import { describe, expect, it } from "vitest";
import {
  EJOOS_EXCLUDED_RNOKPP_COLUMN,
  EJOOS_GENERAL_NUM_FMT_ID,
  EJOOS_OOS_RNOKPP_COLUMN,
  EJOOS_TEXT_NUM_FMT_ID,
  isEjoosRnokppColumn,
  isEjoosStaffIndexColumn,
} from "./ejoosStaffIndexFormat";

describe("ejoosStaffIndexFormat", () => {
  it("maps staff index columns per sheet", () => {
    expect(isEjoosStaffIndexColumn("6. Табель", 2)).toBe(true);
    expect(isEjoosStaffIndexColumn(/табель/i, 2)).toBe(true);
    expect(isEjoosStaffIndexColumn("6. Табель", 3)).toBe(false);
    expect(isEjoosStaffIndexColumn("6. Табель", 2, 2)).toBe(false);
    expect(isEjoosStaffIndexColumn("6. Табель", 2, 7)).toBe(true);
    expect(isEjoosStaffIndexColumn("1. ШПО", 1)).toBe(true);
    expect(isEjoosStaffIndexColumn("3. Виключені", 4)).toBe(true);
  });

  it("maps РНОКПП columns on ООС / Виключені", () => {
    expect(isEjoosRnokppColumn("2. ООС", EJOOS_OOS_RNOKPP_COLUMN)).toBe(true);
    expect(isEjoosRnokppColumn(/оос/i, 22, 7)).toBe(true);
    expect(isEjoosRnokppColumn("2. ООС", 3)).toBe(false);
    expect(isEjoosRnokppColumn("2. ООС", 22, 2)).toBe(false);
    expect(isEjoosRnokppColumn("3. Виключені", EJOOS_EXCLUDED_RNOKPP_COLUMN)).toBe(
      true,
    );
    expect(isEjoosRnokppColumn("3. Виключені", 22)).toBe(false);
  });

  it("uses built-in text and general numFmtId", () => {
    expect(EJOOS_TEXT_NUM_FMT_ID).toBe("49");
    expect(EJOOS_GENERAL_NUM_FMT_ID).toBe("0");
  });
});
