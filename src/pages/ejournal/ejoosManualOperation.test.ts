import { describe, expect, it } from "vitest";
import type {
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  buildManualEjoosOperation,
  collectManualEjoosPeople,
  hydrateManualEjoosOperation,
} from "./ejoosManualOperation";
import {
  personCanEnterApplyQueue,
  personChangesFromOps,
} from "./ejoosPersonDiff";

const sheet = (
  sheetName: string,
  rawRows: Array<Array<string | number | null>>,
): ExcelSheetSnapshot =>
  ({
    sheetIndex: 0,
    sheetName,
    rawRows,
    rows: [],
    headerRows: [],
    columnCount: 40,
    columnIndexes: [],
    dataStartRow: 1,
  }) as unknown as ExcelSheetSnapshot;

const workbook = (options?: { withoutTimesheetPerson?: boolean }) => {
  const empty = Array.from({ length: 6 }, () => []);
  const shpoPerson = Array(8).fill("");
  shpoPerson[0] = "2103117";
  shpoPerson[5] = "солдат";
  shpoPerson[6] = "КІЯНЕНКО Андрій Олександрович";
  shpoPerson[7] = "1961288";
  const occupied = Array(8).fill("");
  occupied[0] = "2103119";
  occupied[5] = "солдат";
  occupied[6] = "СИДОРЕНКО Єгор Олегович";
  occupied[7] = "incumbent-id";
  const free = Array(8).fill("");
  free[0] = "2103120";

  const oosPerson = Array(21).fill("");
  oosPerson[0] = "солдат";
  oosPerson[1] = "КІЯНЕНКО Андрій Олександрович";
  oosPerson[2] = "1961288";
  oosPerson[3] = "2103117";

  const timesheetPerson = Array(40).fill("");
  timesheetPerson[1] = "2103117";
  timesheetPerson[5] = "солдат";
  timesheetPerson[6] = options?.withoutTimesheetPerson
    ? ""
    : "КІЯНЕНКО Андрій Олександрович";
  timesheetPerson[7] = options?.withoutTimesheetPerson ? "" : "1961288";
  timesheetPerson[8] = "+";
  const occupiedTimesheet = Array(40).fill("");
  occupiedTimesheet[1] = "2103119";
  occupiedTimesheet[6] = "СИДОРЕНКО Єгор Олегович";
  occupiedTimesheet[7] = "incumbent-id";
  const freeTimesheet = Array(40).fill("");
  freeTimesheet[1] = "2103120";

  return {
    file: new Blob([]) as File,
    fileName: "ЄЖООС.xlsx",
    sheetName: "1. ШПО",
    rows: [],
    headerRows: [],
    columnCount: 40,
    columnIndexes: [],
    dataStartRow: 1,
    sheets: [
      sheet("1. ШПО", [...empty, shpoPerson, occupied, free]),
      sheet("2. ООС", [...empty.slice(0, 5), oosPerson]),
      sheet("6. Табель", [
        ...empty,
        timesheetPerson,
        occupiedTimesheet,
        freeTimesheet,
      ]),
    ],
  } as ExcelWorkbookSnapshot;
};

const baseValues = {
  personKey: "id:1961288",
  orderNumber: "123",
  orderDate: "2026-09-05",
};

describe("manual EJOOS operation builder", () => {
  it("collects current personnel and builds an atomic-ready ПЕРЕВ preview", () => {
    const ejoos = workbook();
    expect(collectManualEjoosPeople(ejoos)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          personId: "1961288",
          positionIndex: "2103117",
        }),
      ]),
    );

    const op = buildManualEjoosOperation({
      ejoos,
      timesheetDay: 5,
      values: {
        ...baseValues,
        type: "exclude_transfer",
        destination: "військова частина А0001",
      },
    });

    expect(op.class).toBe("ready");
    expect(op.payload).toMatchObject({
      manualOperation: "1",
      shpoExcelRow: "7",
      oosExcelRow: "6",
      timesheetExcelRow: "7",
      fromPositionIndex: "2103117",
      destination: "військова частина А0001",
    });
    expect(
      personCanEnterApplyQueue(personChangesFromOps([op], 5)[0]),
    ).toBe(true);
  });

  it("blocks ПЕРЕВ when no active Timesheet source can be resolved", () => {
    const op = buildManualEjoosOperation({
      ejoos: workbook({ withoutTimesheetPerson: true }),
      timesheetDay: 5,
      values: {
        ...baseValues,
        type: "exclude_transfer",
        destination: "військова частина А0001",
      },
    });

    expect(op.class).toBe("conflict");
    expect(op.payload.timesheetExcelRow).toBe("");
    expect(personCanEnterApplyQueue(personChangesFromOps([op], 5)[0])).toBe(
      false,
    );
  });

  it("builds ЗВІЛЬН through the same atomic exclusion route", () => {
    const op = buildManualEjoosOperation({
      ejoos: workbook(),
      timesheetDay: 5,
      values: {
        ...baseValues,
        type: "dismissal",
        destination: "звільнений у запас",
      },
    });

    expect(op).toMatchObject({
      kind: "exclude_transfer",
      class: "ready",
      payload: {
        type: "ЗВІЛЬН",
        exclusionReason: "ЗВІЛЬНЕННЯ",
        destination: "звільнений у запас",
        timesheetAction: "MOVE_TO_HISTORY",
      },
    });
  });

  it("reports the incumbent conflict for a manual position change", () => {
    const op = buildManualEjoosOperation({
      ejoos: workbook(),
      timesheetDay: 5,
      values: {
        ...baseValues,
        type: "position_change",
        nextPositionIndex: "2103119",
      },
    });

    expect(op.class).toBe("conflict");
    expect(op.payload.targetOccupantName).toBe(
      "СИДОРЕНКО Єгор Олегович",
    );
  });

  it("builds ready position and rank operations with workbook row targets", () => {
    const ejoos = workbook();
    const position = buildManualEjoosOperation({
      ejoos,
      timesheetDay: 5,
      values: {
        ...baseValues,
        type: "position_change",
        nextPositionIndex: "2103120",
      },
    });
    const rank = buildManualEjoosOperation({
      ejoos,
      timesheetDay: 5,
      values: {
        ...baseValues,
        type: "rank_change",
        nextRank: "старший солдат",
      },
    });

    expect(position.class).toBe("ready");
    expect(position.payload).toMatchObject({
      previousIndex: "2103117",
      nextIndex: "2103120",
      shpoExcelRow: "9",
      timesheetExcelRow: "9",
    });
    expect(rank.class).toBe("ready");
    expect(rank.payload).toMatchObject({
      previousRank: "солдат",
      nextRank: "старший солдат",
      shpoExcelRow: "7",
      oosExcelRow: "6",
      timesheetExcelRow: "7",
    });
  });

  it("rehydrates a persisted draft with its stable server ID and decision", () => {
    const op = hydrateManualEjoosOperation({
      ejoos: workbook(),
      timesheetDay: 5,
      draft: {
        id: "draft-server-id",
        unitLabel: "1ПБ",
        status: "draft",
        decision: "accepted",
        input: {
          ...baseValues,
          type: "rank_change",
          nextRank: "старший солдат",
        },
        baseVersionId: "version-10",
        createdByEmail: "operator@example.test",
        createdAt: "2026-09-05T10:00:00.000Z",
        updatedAt: "2026-09-05T10:05:00.000Z",
      },
    });

    expect(op).toMatchObject({
      id: "draft-server-id",
      class: "ready",
      payload: {
        manualDraftId: "draft-server-id",
        manualDecision: "accepted",
        manualBaseVersionId: "version-10",
        manualCreatedBy: "operator@example.test",
      },
    });
  });

  it("requires re-saving a draft anchored to an older workbook version", () => {
    const op = hydrateManualEjoosOperation({
      ejoos: workbook(),
      timesheetDay: 5,
      currentVersionId: "version-11",
      draft: {
        id: "stale-draft",
        unitLabel: "1ПБ",
        status: "draft",
        decision: "accepted",
        input: {
          ...baseValues,
          type: "rank_change",
          nextRank: "старший солдат",
        },
        baseVersionId: "version-10",
        createdAt: "2026-09-05T10:00:00.000Z",
        updatedAt: "2026-09-05T10:05:00.000Z",
      },
    });

    expect(op?.class).toBe("needs_input");
    expect(op?.checkedDefault).toBe(false);
    expect(op?.why).toMatch(/попередньої версії/i);
  });
});
