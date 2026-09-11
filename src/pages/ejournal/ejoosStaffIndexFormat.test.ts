import { describe, expect, it } from "vitest";
import {
  EJOOS_TEXT_NUM_FMT_ID,
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

  it("uses built-in text numFmtId", () => {
    expect(EJOOS_TEXT_NUM_FMT_ID).toBe("49");
  });
});
