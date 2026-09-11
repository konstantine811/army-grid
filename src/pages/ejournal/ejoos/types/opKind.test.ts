import { describe, expect, it } from "vitest";
import { EJOOS_OP_KINDS } from "./syncOp";

describe("EJOOS op kind contract", () => {
  it("lists all supported op kinds for plan/apply routing", () => {
    expect(EJOOS_OP_KINDS).toHaveLength(12);
    expect([...EJOOS_OP_KINDS].sort()).toEqual(
      [
        "absent_close",
        "absent_upsert",
        "arrival",
        "contract_update",
        "data_mismatch",
        "exclude_transfer",
        "move_to_disposition",
        "other_manual",
        "position_change",
        "rank_change",
        "shpo_occupant",
        "timesheet_day",
      ].sort(),
    );
  });
});
