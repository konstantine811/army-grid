import { describe, expect, it } from "vitest";
import { applyConfirmedEjoosOps } from "./ejoosSyncApply";
import {
  MONTH_ROLLOVER_BLOCK_MESSAGE,
  planBlocksWorkbookApply,
  refreshPlanTimesheetHorizon,
  sourceTimesheetHorizonNote,
  type EjoosSyncOp,
  type EjoosSyncPlan,
} from "./ejoosSyncPlan";
import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";

const plan = (
  extra: Partial<EjoosSyncPlan> = {},
): EjoosSyncPlan => ({
  ejoosName: "ejoos.xlsx",
  pbName: "1ПБ_25082026.xlsx",
  timesheetDay: 25,
  timesheetDayLabel: "25.08.2026",
  ops: [],
  summary: { ready: 0, needsInput: 0, conflict: 0 },
  ...extra,
});

describe("timesheet horizon stays on the 1PB source date", () => {
  it("does not stretch 25.08 to 31.08 just because today is later", () => {
    expect(
      refreshPlanTimesheetHorizon(
        { timesheetDay: 25, timesheetDayLabel: "25.08.2026" },
        new Date(2026, 7, 31),
      ),
    ).toEqual({ timesheetDay: 25, timesheetDayLabel: "25.08.2026" });
  });

  it("states the source-as-of limit in the operator note", () => {
    expect(sourceTimesheetHorizonNote("25.08.2026")).toBe(
      "Джерело 1ПБ станом на 25.08.2026. Табель буде оновлено лише по 25.08.2026.",
    );
  });
});

describe("month rollover blocks apply", () => {
  it("treats a mismatched timesheet month as a workbook apply block", () => {
    const blocked = plan({ monthRolloverRequired: true });
    expect(planBlocksWorkbookApply(blocked)).toBe(true);
    expect(planBlocksWorkbookApply(plan())).toBe(false);
  });

  it("applyConfirmedEjoosOps refuses to write when rollover is required", async () => {
    const op: EjoosSyncOp = {
      id: "op-1",
      kind: "timesheet_day",
      class: "ready",
      sheet: "6. Табель",
      personId: "1",
      fullName: "ТЕСТ",
      positionIndex: "",
      rank: "",
      before: "",
      after: "",
      sourceRef: "",
      why: "",
      confidence: "high",
      payload: {},
      checkedDefault: true,
    };
    const file = new File([new Uint8Array([80, 75])], "ejoos.xlsx");
    const ejoos = {
      file,
      fileName: "ejoos.xlsx",
      sheetName: "",
      headerRows: [],
      rows: [],
      columnCount: 0,
      columnIndexes: [],
      dataStartRow: 1,
      sheets: [],
    } as ExcelWorkbookSnapshot;
    await expect(
      applyConfirmedEjoosOps({
        ejoos,
        plan: plan({ monthRolloverRequired: true, ops: [op] }),
        ops: [op],
      }),
    ).rejects.toThrow(MONTH_ROLLOVER_BLOCK_MESSAGE);
  });
});
