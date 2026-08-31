import { describe, expect, it } from "vitest";
import {
  isInformationalOp,
  isWorkbookApplyOp,
  type PersonChange,
} from "./ejoosPersonDiff";
import type { EjoosSyncOp } from "./ejoosSyncPlan";

const op = (partial: Partial<EjoosSyncOp> & Pick<EjoosSyncOp, "kind">): EjoosSyncOp => ({
  id: "op-1",
  class: "needs_input",
  sheet: "тест",
  personId: "1",
  fullName: "ТЕСТ",
  positionIndex: "",
  rank: "",
  before: "",
  after: "",
  sourceRef: "",
  why: "",
  confidence: "manual",
  payload: {},
  checkedDefault: false,
  ...partial,
});

describe("isWorkbookApplyOp", () => {
  it("rejects standalone ПРИБУВ — there is no apply handler", () => {
    expect(isWorkbookApplyOp(op({ kind: "arrival" }))).toBe(false);
  });

  it("rejects generic other_manual", () => {
    expect(
      isWorkbookApplyOp(op({ kind: "other_manual", payload: { type: "ЗВІЛЬН" } })),
    ).toBe(false);
  });

  it("treats CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH as review-only", () => {
    const cancel = op({
      kind: "other_manual",
      payload: {
        type: "TRANSFER_CANCELLED",
        reviewReason: "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH",
      },
    });
    expect(isInformationalOp(cancel)).toBe(true);
    expect(isWorkbookApplyOp(cancel)).toBe(false);
  });

  it("does not apply TRANSFER_CANCELLED when nothing is left to write", () => {
    const cancel = op({
      kind: "other_manual",
      class: "ready",
      payload: { type: "TRANSFER_CANCELLED" },
    });
    expect(isInformationalOp(cancel)).toBe(true);
    expect(isWorkbookApplyOp(cancel)).toBe(false);
  });

  it("applies TRANSFER_CANCELLED when timesheet or excluded still needs a write", () => {
    const withExcluded = op({
      kind: "other_manual",
      class: "ready",
      payload: { type: "TRANSFER_CANCELLED", excludedExcelRow: "6" },
    });
    const withTimesheet = op({
      kind: "other_manual",
      class: "ready",
      payload: { type: "TRANSFER_CANCELLED", restoreTimesheet: "1" },
    });
    expect(isWorkbookApplyOp(withExcluded)).toBe(true);
    expect(isWorkbookApplyOp(withTimesheet)).toBe(true);
  });

  it("allows exclude_transfer", () => {
    expect(
      isWorkbookApplyOp(op({ kind: "exclude_transfer", class: "ready" })),
    ).toBe(true);
  });
});

describe("personHasWorkbookApplyOps typing smoke", () => {
  it("arrival-only person has no workbook apply", () => {
    const person = {
      ops: [op({ kind: "arrival" })],
    } as Pick<PersonChange, "ops">;
    expect(person.ops.some(isWorkbookApplyOp)).toBe(false);
  });
});
