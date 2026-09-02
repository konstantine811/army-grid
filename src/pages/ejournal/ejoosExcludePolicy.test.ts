import { describe, expect, it } from "vitest";
import {
  absenceOnlyBlocksExclusion,
  excludedRowsToClear,
  excludedTimesheetWrite,
  externalTransferProcessState,
  isFalseInternalHopExclusion,
  staleExcludedRowsToClear,
  laterReturnSupersedesOutbound,
  positionCloseWritesExcluded,
  skipExternalIfAlreadyProcessed,
} from "./ejoosExcludePolicy";

describe("externalTransferProcessState", () => {
  it("is ALREADY_PROCESSED only when all four invariants hold", () => {
    expect(
      externalTransferProcessState({
        onStaffShpo: false,
        onStaffOos: false,
        hasMatchingExcluded: true,
        timesheetClosed: true,
      }),
    ).toBe("ALREADY_PROCESSED");
  });

  it("is PARTIALLY_PROCESSED when Excluded is missing", () => {
    expect(
      externalTransferProcessState({
        onStaffShpo: false,
        onStaffOos: false,
        hasMatchingExcluded: false,
        timesheetClosed: true,
      }),
    ).toBe("PARTIALLY_PROCESSED");
  });

  it("does not skip a partial external transfer", () => {
    expect(
      skipExternalIfAlreadyProcessed({
        stillInEjoos: false,
        unrecordedTransit: false,
        processState: "PARTIALLY_PROCESSED",
      }),
    ).toBe(false);
    expect(
      skipExternalIfAlreadyProcessed({
        stillInEjoos: false,
        unrecordedTransit: false,
        processState: "ALREADY_PROCESSED",
      }),
    ).toBe(true);
  });
});

describe("laterReturnSupersedesOutbound", () => {
  it("does not exclude Atrakhov after 05.08 leave and 07.08 return", () => {
    expect(
      laterReturnSupersedesOutbound({
        stillInSh: true,
        returnedAfterOutbound: true,
      }),
    ).toBe(true);
    expect(
      laterReturnSupersedesOutbound({
        stillInSh: false,
        returnedAfterOutbound: true,
      }),
    ).toBe(true);
    expect(
      laterReturnSupersedesOutbound({
        stillInSh: false,
        returnedAfterOutbound: false,
      }),
    ).toBe(false);
  });
});

describe("excludedTimesheetWrite", () => {
  it("replaces the existing вибув row instead of inserting", () => {
    expect(
      excludedTimesheetWrite([
        { excelRow: 140, hasDepartureText: true },
      ]),
    ).toEqual({
      createHistory: false,
      replaceInPlace: true,
      sourceExcelRow: 140,
    });
  });

  it("closes the active episode and keeps the 05.08 вибув history", () => {
    expect(
      excludedTimesheetWrite([
        { excelRow: 110, hasDepartureText: false },
        { excelRow: 140, hasDepartureText: true },
      ]),
    ).toEqual({
      createHistory: false,
      replaceInPlace: false,
      sourceExcelRow: 110,
    });
  });

  it("copies the active staff row down instead of painting вибув on the slot", () => {
    expect(
      excludedTimesheetWrite([
        { excelRow: 88, hasDepartureText: false },
      ]),
    ).toEqual({
      createHistory: false,
      replaceInPlace: false,
      sourceExcelRow: 88,
    });
  });

  it("creates history only when the person is not on the timesheet", () => {
    expect(excludedTimesheetWrite([])).toEqual({
      createHistory: true,
      replaceInPlace: false,
      sourceExcelRow: 0,
    });
  });

  it("copies from the vacant staff slot instead of painting the person onto it", () => {
    expect(excludedTimesheetWrite([], 7)).toEqual({
      createHistory: false,
      replaceInPlace: false,
      sourceExcelRow: 7,
    });
  });
});

describe("positionCloseWritesExcluded", () => {
  it("does not write Виключені for an internal 1ПБ index hop", () => {
    expect(
      positionCloseWritesExcluded({
        closeOldPosition: "1",
        internalStaffHop: "1",
      }),
    ).toBe(false);
  });

  it("still writes Виключені when closing a position that left the unit", () => {
    expect(
      positionCloseWritesExcluded({
        closeOldPosition: "1",
      }),
    ).toBe(true);
  });

  it("infers an internal hop from 1ПБ indexes even without the flag", () => {
    expect(
      positionCloseWritesExcluded({
        closeOldPosition: "1",
        previousIndex: "2103725",
        nextIndex: "2103255",
        statusRaw: "_5 1ПБ",
        changeText:
          "номер обслуги кулеметного взводу 1 піхотної роти 1 піхотного батальйону",
      }),
    ).toBe(false);
  });
});

describe("isFalseInternalHopExclusion", () => {
  it("flags Шибряєв-style ПЕРЕВЕДЕННЯ 1 ПБ while still in unit", () => {
    expect(
      isFalseInternalHopExclusion({
        destination:
          "номер обслуги кулеметного взводу 1 піхотної роти 1 піхотного батальйону",
        note: "ПЕРЕВЕДЕННЯ _ 1 ПБ",
      }),
    ).toBe(true);
  });

  it("does not flag a real transfer to another в/ч", () => {
    expect(
      isFalseInternalHopExclusion({
        destination: "в/ч А4784",
        note: "ПЕРЕВЕДЕННЯ",
      }),
    ).toBe(false);
  });

  it("flags Сіряченко-style hop into a 1ПБ position title", () => {
    expect(
      isFalseInternalHopExclusion({
        destination: "вогнеметник 2 піхотного відділення 1 піхотного батальйону",
        note: "ПЕРЕВЕДЕННЯ _ 1 ПБ",
      }),
    ).toBe(true);
  });

  it("does not flag ПЕРЕВЕДЕННЯ without 1ПБ in note or destination", () => {
    expect(
      isFalseInternalHopExclusion({
        destination: "стрілець",
        note: "ПЕРЕВЕДЕННЯ",
      }),
    ).toBe(false);
  });
});

describe("staleExcludedRowsToClear", () => {
  it("Гіріченко: old April row is stale against the current August ПЕРЕВ", () => {
    expect(
      staleExcludedRowsToClear({
        rows: [
          { excelRow: 1193, orderNumber: "108", orderDate: "13.04.2026" },
          { excelRow: 1901, orderNumber: "232", orderDate: "11.08.2026" },
        ],
        currentOrderNumber: "232",
        currentOrderDate: "11.08.2026",
      }).map((row) => row.excelRow),
    ).toEqual([1193]);
  });

  it("treats every previous row as stale when the current order is not yet written", () => {
    expect(
      staleExcludedRowsToClear({
        rows: [{ excelRow: 1193, orderNumber: "108", orderDate: "13.04.2026" }],
        currentOrderNumber: "232",
        currentOrderDate: "11.08.2026",
      }).map((row) => row.excelRow),
    ).toEqual([1193]);
  });

  it("keeps the matching current row and drops earlier duplicates", () => {
    expect(
      staleExcludedRowsToClear({
        rows: [
          { excelRow: 6, orderNumber: "108", orderDate: "13.04.2026" },
          { excelRow: 7, orderNumber: "232", orderDate: "11.08.2026" },
        ],
        currentOrderNumber: "232",
        currentOrderDate: "11.08.2026",
      }).map((row) => row.excelRow),
    ).toEqual([6]);
  });
});

describe("excludedRowsToClear", () => {
  it("reads both a single row and a comma-separated list", () => {
    expect(
      excludedRowsToClear({
        clearExcludedExcelRow: "6",
        clearExcludedExcelRows: "6, 12",
      }),
    ).toEqual([6, 12]);
  });
});

describe("absenceOnlyBlocksExclusion", () => {
  it("blocks an older ПЕРЕВ when БЕЗВІСТИ is later", () => {
    expect(
      absenceOnlyBlocksExclusion({
        absenceAt: Date.UTC(2026, 7, 13),
        outboundAt: Date.UTC(2026, 7, 4),
      }),
    ).toBe(true);
  });

  it("does not block a later real outbound after MIA", () => {
    expect(
      absenceOnlyBlocksExclusion({
        absenceAt: Date.UTC(2026, 7, 13),
        outboundAt: Date.UTC(2026, 7, 20),
      }),
    ).toBe(false);
  });
});
