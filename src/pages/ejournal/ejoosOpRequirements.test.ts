import { describe, expect, it } from "vitest";
import type { EjoosSyncOp } from "./ejoosSyncPlan";
import {
  personApplyBlockReason,
  personOpsBlockApply,
} from "./ejoosOpRequirements";
import { personChangesFromOps } from "./ejoosPersonDiff";

const op = (partial: Partial<EjoosSyncOp>): EjoosSyncOp => ({
  id: "test",
  kind: "exclude_transfer",
  class: "ready",
  sheet: "test",
  personId: "1",
  fullName: "TEST",
  positionIndex: "2103000",
  rank: "солдат",
  before: "",
  after: "",
  sourceRef: "",
  why: "",
  confidence: "high",
  payload: {},
  ...partial,
});

describe("personOpsBlockApply", () => {
  it("does not block exclude_transfer when archive return review is grouped", () => {
    const ops = [
      op({
        payload: {
          destination: "в/ч А1234",
          excludeDate: "28.01.2026",
          orderNumber: "100",
          orderDate: "28.01.2026",
          timesheetActiveFrom: "10.01.2026",
          timesheetAbsenceSpans: "11-19:лік|20-27:СЗЧ",
        },
      }),
      op({
        kind: "absent_upsert",
        class: "needs_input",
        payload: {
          mismatchKind: "ARCHIVE_RETURN_SH_STILL_ABSENT",
          returnDate: "09.01.2026",
          statusRaw: "СЗЧ",
        },
      }),
    ];
    expect(personOpsBlockApply(ops)).toBe(false);
    expect(personApplyBlockReason(ops)).toBeNull();
  });

  it("does not block a ready ПЕРЕВ because of sibling archive / Табель ops", () => {
    const ops = [
      op({
        payload: {
          destination: "НА_ЩИТІ",
          excludeDate: "14.08.2026",
          orderNumber: "235",
          orderDate: "14.08.2026",
        },
      }),
      op({
        kind: "absent_upsert",
        class: "needs_input",
        payload: {
          absenceType: "БЕЗВІСТИ",
          departDate: "20.07.2026",
          timesheetCode: "ЗБ",
        },
      }),
      op({
        kind: "timesheet_day",
        class: "needs_input",
        payload: {
          day: "14",
          timesheetCode: "",
        },
      }),
    ];
    expect(personOpsBlockApply(ops)).toBe(false);
    expect(personApplyBlockReason(ops)).toBeNull();
    expect(personChangesFromOps(ops, 14)[0]?.severity).toBe("ready");
  });

  it("allows exclude_transfer when only documentsDest is filled", () => {
    const ops = [
      op({
        payload: {
          destination: "",
          documentsDest: "в/ч А4784, м. Київ",
          excludeDate: "28.08.2026",
          orderNumber: "100",
          orderDate: "28.08.2026",
        },
      }),
    ];
    expect(personOpsBlockApply(ops)).toBe(false);
  });

  it("allows return from disposition with open СЗЧ and later РОЗПОРЯДЖ", () => {
    const ops = [
      op({
        kind: "position_change",
        payload: {
          returningFromDisposition: "1",
          openAbsenceExcelRow: "55",
          orderDate: "10.08.2026",
          nextIndex: "2103229",
        },
      }),
      op({
        kind: "move_to_disposition",
        payload: {
          orderDate: "28.08.2026",
          keepOpenSzchTimesheet: "1",
        },
      }),
    ];
    expect(personOpsBlockApply(ops)).toBe(false);
    expect(personApplyBlockReason(ops)).toBeNull();
  });

  it("still blocks incomplete exclude_transfer", () => {
    const ops = [
      op({
        payload: {
          destination: "",
          excludeDate: "28.01.2026",
          orderNumber: "100",
          orderDate: "28.01.2026",
        },
      }),
    ];
    expect(personOpsBlockApply(ops)).toBe(true);
    expect(personApplyBlockReason(ops)).toMatch(/куди вибув/i);
  });
});
