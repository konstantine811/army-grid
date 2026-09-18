import { describe, expect, it } from "vitest";
import type {
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  buildManualEjoosOperation,
  collectManualEjoosPeople,
  hydrateManualEjoosOperation,
  manualInputFromBackend,
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

  it.each([
    { type: "move_to_disposition" as const, kind: "move_to_disposition", event: "РОЗПОРЯДЖ", destination: ", який знаходиться у розпорядженні командира військової частини А4862", orderNumber: "№245", orderDate: "2026-08-24", basisIssuer: "командувача Сухопутних військ Збройних Сил України", basisNumber: "№743-РС", basisDate: "2026-08-24" },
    { type: "exclusion" as const, kind: "exclude_transfer", event: "ВИКЛЮЧ", destination: "НА_ЩИТІ", orderNumber: "№ 234", orderDate: "2026-08-13", basisIssuer: "командира військової частини А4862", basisNumber: "№ 304-РС", basisDate: "2026-08-11" },
  ])("builds and restores manual $event with both orders", (example) => {
    const values = { ...baseValues, ...example };
    const op = buildManualEjoosOperation({ ejoos: workbook(), timesheetDay: 25, values });
    expect(op.kind).toBe(example.kind);
    expect(op.class).toBe("ready");
    expect(op.payload.type).toBe(example.event);
    expect(op.payload.destination).toBe(example.destination);
    expect(op.payload.basisIssuer).toBe(example.basisIssuer);
    expect(op.payload.basisNumber).toBe(example.basisNumber.replace(/^№\s*/, ""));
    expect(op.after).toContain(example.basisIssuer);
    expect(personCanEnterApplyQueue(personChangesFromOps([op], 25)[0])).toBe(true);
    if (example.type === "move_to_disposition") expect(op.payload.remainsInOos).toBe("true");
    else expect(op.payload.exclusionReason).toBe("НА_ЩИТІ");
    const draft = { id: "manual-new-kind", unitLabel: "1ПБ", status: "draft" as const, decision: "pending" as const, input: values, baseVersionId: "version-1", createdAt: "", updatedAt: "" };
    expect(manualInputFromBackend(draft)).toMatchObject({ type: example.type, basisIssuer: example.basisIssuer, basisNumber: example.basisNumber, basisDate: example.basisDate });
    const restored = hydrateManualEjoosOperation({ draft, ejoos: workbook(), timesheetDay: 25 });
    expect(restored?.kind).toBe(example.kind);
    expect(restored?.payload).toMatchObject(op.payload);
  });

  it("does not mark disposition ready when the destination or active source is missing", () => {
    for (const [ejoos, destination] of [[workbook(), ""], [workbook({ withoutTimesheetPerson: true }), "у розпорядження командира"]] as const) {
      const op = buildManualEjoosOperation({ ejoos, timesheetDay: 5, values: { ...baseValues, type: "move_to_disposition", destination } });
      expect(op.class).toBe("needs_input");
      expect(op.checkedDefault).toBe(false);
    }
  });

  it.each([["СЗЧ", "СЗЧ"], ["БЕЗВІСТИ", "ЗБ"]])("keeps open %s when placing a person in disposition", (ground, code) => {
    const ejoos = workbook();
    const absence = Array(13).fill("");
    absence[1] = "КІЯНЕНКО Андрій Олександрович";
    absence[2] = "1961288";
    absence[3] = "2103117";
    absence[4] = ground;
    ejoos.sheets.push(sheet("5. Тимчасово відсутні", [...Array.from({ length: 5 }, () => []), absence]));
    const op = buildManualEjoosOperation({ ejoos, timesheetDay: 5, values: { ...baseValues, type: "move_to_disposition", destination: "у розпорядження командира" } });
    expect(op.payload).toMatchObject({ absenceExcelRow: "6", absenceCode: code, keepOpenSzchTimesheet: "1", needsAbsenceRecord: "", remainsInOos: "true" });
  });

  it("uses the shared ВП template for a wounded recovery leave", () => {
    const ejoos = workbook();
    const absence = Array(13).fill("");
    absence[1] = "КІЯНЕНКО Андрій Олександрович";
    absence[2] = "1961288";
    absence[4] = "ВІДПУСТКА ДЛЯ ЛІКУВАННЯ ПІСЛЯ   ПОРАНЕННЯ";
    ejoos.sheets.push(sheet("5. Тимчасово відсутні", [...Array.from({ length: 5 }, () => []), absence]));
    const op = buildManualEjoosOperation({ ejoos, timesheetDay: 5, values: { ...baseValues, type: "move_to_disposition", destination: "у розпорядження командира" } });
    expect(op.payload.absenceCode).toBe("ВП");
    expect(op.payload.keepOpenSzchTimesheet).toBe("");
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

  it("keeps a ready draft when the newer workbook still supports it", () => {
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

    expect(op?.class).toBe("ready");
    expect(op?.checkedDefault).toBe(true);
    expect(op?.why).not.toMatch(/попередньої версії/i);
    expect(
      personCanEnterApplyQueue(personChangesFromOps([op!], 5)[0]),
    ).toBe(true);
  });

  it("asks to re-save a draft only when the newer workbook no longer supports it", () => {
    const op = hydrateManualEjoosOperation({
      ejoos: workbook({ withoutTimesheetPerson: true }),
      timesheetDay: 5,
      currentVersionId: "version-11",
      draft: {
        id: "stale-blocked-draft",
        unitLabel: "1ПБ",
        status: "draft",
        decision: "accepted",
        input: {
          ...baseValues,
          type: "exclude_transfer",
          destination: "НА_ЩИТІ",
        },
        baseVersionId: "version-10",
        createdAt: "2026-09-05T10:00:00.000Z",
        updatedAt: "2026-09-05T10:05:00.000Z",
      },
    });

    expect(op?.class).toBe("conflict");
    expect(op?.checkedDefault).toBe(false);
    expect(op?.why).toMatch(/попередньої версії/i);
    expect(
      personCanEnterApplyQueue(personChangesFromOps([op!], 5)[0]),
    ).toBe(false);
  });
});
