import { describe, expect, it } from "vitest";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
  type ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { applyExcludeTransfersWithZip } from "./ejoosExcludeTransferZip";
import { loadTimesheetGridFromFile } from "./ejoosTimesheetPersonRows";
import type { EjoosSyncOp, EjoosSyncPlan } from "./ejoosSyncPlan";
import { TimesheetTransferError } from "./ejoosTimesheetTransferAction";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

if (typeof window === "undefined") {
  (globalThis as { window?: typeof globalThis }).window = globalThis;
}

if (typeof FileReader === "undefined") {
  class NodeFileReader {
    result: ArrayBuffer | null = null;
    onload: ((event: { target: NodeFileReader }) => void) | null = null;
    onerror: ((event: { target: NodeFileReader }) => void) | null = null;
    onloadend: ((event: { target: NodeFileReader }) => void) | null = null;
    readyState = 0;
    readAsArrayBuffer(blob: Blob) {
      void blob.arrayBuffer().then((buffer) => {
        this.result = buffer;
        this.readyState = 2;
        const event = { target: this };
        this.onload?.(event);
        this.onloadend?.(event);
      });
    }
    readAsBinaryString() {}
    readAsDataURL() {}
    readAsText() {}
    abort() {}
    addEventListener() {}
    removeEventListener() {}
    dispatchEvent() {
      return false;
    }
  }
  globalThis.FileReader = NodeFileReader as unknown as typeof FileReader;
}

const loadPopulate = async () => {
  const module = await import(
    "xlsx-populate/browser/xlsx-populate-no-encryption"
  );
  return module.default;
};

const snapshotOf = async (
  blob: Blob,
  fileName: string,
): Promise<ExcelWorkbookSnapshot> =>
  readWorkbookSnapshot(
    new File([blob], fileName, { type: XLSX_MIME }),
    EJOOS_SYNC_READ_OPTIONS,
  );

const lekay = {
  name: "ЛЕКАЙ Володимир Петрович",
  id: "18801",
  index: "2103777",
  rank: "солдат",
};

const incompleteHint: Record<string, string> = {
  destination: "А4784",
  timesheetDestination: "в/ч А4784",
  documentsDest: "А4784",
  excludeDate: "19.08.2026",
  orderNumber: "240",
  orderDate: "19.08.2026",
  fromRank: lekay.rank,
  fromName: lekay.name,
  fromPersonId: lekay.id,
  fromPositionIndex: lekay.index,
  previousIndex: lekay.index,
  shpoExcelRow: "7",
  oosExcelRow: "6",
  timesheetCreateHistory: "",
  timesheetReplaceInPlace: "",
  timesheetExcelRow: "",
};

const excludeOp = (payload: Record<string, string>): EjoosSyncOp => ({
  id: "excl-lekay",
  kind: "exclude_transfer",
  class: "ready",
  sheet: "Виключені → Табель → ШПО/ООС",
  personId: lekay.id,
  fullName: lekay.name,
  positionIndex: lekay.index,
  rank: lekay.rank,
  before: "",
  after: "",
  sourceRef: "",
  why: "",
  confidence: "high",
  checkedDefault: true,
  payload,
});

const planOf = (op: EjoosSyncOp): EjoosSyncPlan => ({
  ejoosName: "ЄЖООС.xlsx",
  pbName: "1ПБ.xlsx",
  timesheetDay: 25,
  timesheetDayLabel: "25.08.2026",
  ops: [op],
  summary: { ready: 1, needsInput: 0, conflict: 0 },
});

const buildWorkbook = async (onTimesheet: boolean) => {
  const XlsxPopulate = await loadPopulate();
  const workbook = await XlsxPopulate.fromBlankAsync();
  const shpo = workbook.sheet(0);
  shpo.name("1. ШПО");
  const oos = workbook.addSheet("2. ООС");
  const excluded = workbook.addSheet("3. Виключені");
  excluded.cell(1, 1).value("3. Виключені");
  excluded.cell(5, 2).value("ПІБ");
  workbook.addSheet("4. Тимчасово прибулі");
  workbook.addSheet("5. Тимчасово відсутні");
  const timesheet = workbook.addSheet("6. Табель");
  shpo.cell(7, 1).value(lekay.index);
  shpo.cell(7, 6).value(lekay.rank);
  shpo.cell(7, 7).value(lekay.name);
  shpo.cell(7, 8).value(lekay.id);
  oos.cell(6, 1).value(lekay.rank);
  oos.cell(6, 2).value(lekay.name);
  oos.cell(6, 3).value(lekay.id);
  oos.cell(6, 4).value(lekay.index);
  timesheet.cell(2, 9).value("Серпень 2026 р.");
  timesheet.cell(6, 1).value("1 ПІХОТНА РОТА");
  timesheet.cell(6, 2).value("1 ПІХОТНА РОТА");
  timesheet.cell(7, 2).value("2103101");
  timesheet.cell(7, 7).value("ІНШИЙ Першої");
  timesheet.cell(7, 8).value("10001");
  if (onTimesheet) {
    timesheet.cell(20, 1).value("1 ПІХОТНА РОТА");
    timesheet.cell(20, 2).value(lekay.index);
    timesheet.cell(20, 3).value("100");
    timesheet.cell(20, 4).value("7");
    timesheet.cell(20, 6).value(lekay.rank);
    timesheet.cell(20, 7).value(lekay.name);
    timesheet.cell(20, 8).value(lekay.id);
    timesheet.cell(20, 9).value("+");
  } else {
    timesheet.cell(20, 2).value("2103999");
    timesheet.cell(20, 7).value("ЧУЖИЙ Рядок");
    timesheet.cell(20, 8).value("19991");
  }
  return snapshotOf(
    (await workbook.outputAsync("blob")) as Blob,
    "ЄЖООС_лекay.xlsx",
  );
};

describe("atomic exclude_transfer Apply", () => {
  it("ЛЕКАЙ: incomplete Tab hint still writes Виключені + history Tab + SHPO/OOS", async () => {
    const ejoos = await buildWorkbook(true);
    const op = excludeOp(incompleteHint);
    const blob = await applyExcludeTransfersWithZip({
      ejoos,
      plan: planOf(op),
      ops: [op],
    });
    const excludedGrid = await loadTimesheetGridFromFile(blob, "3. Виключені");
    const oosGrid = await loadTimesheetGridFromFile(blob, "2. ООС");
    const shpoGrid = await loadTimesheetGridFromFile(blob, "1. ШПО");
    const excludedText = excludedGrid
      .flat()
      .map((cell) => String(cell ?? ""))
      .join(" ");
    expect(excludedText).toMatch(/ЛЕКАЙ/i);
    expect(String(shpoGrid[6]?.[6] ?? "")).toBe("");
    expect(String(shpoGrid[6]?.[7] ?? "")).toBe("");
    expect(String(oosGrid[5]?.[1] ?? "")).toBe("");
    const grid = await loadTimesheetGridFromFile(blob, "6. Табель");
    const cell = (row: number, column: number) =>
      String(grid[row - 1]?.[column - 1] ?? "").trim();
    expect(cell(20, 7)).toBe("");
    expect(cell(20, 8)).toBe("");
    const history = grid.find((row) =>
      String(row?.[6] ?? "").includes("ЛЕКАЙ"),
    );
    expect(history?.[6]).toMatch(/ЛЕКАЙ/i);
    expect(
      Array.from({ length: 31 }, (_, day) => String(history?.[8 + day] ?? "")).some(
        (mark) => /вибув/i.test(mark),
      ),
    ).toBe(true);
  }, 30_000);

  it("ЛЕКАЙ: incomplete hint without a Tab source blocks the whole Apply", async () => {
    const ejoos = await buildWorkbook(false);
    const op = excludeOp(incompleteHint);
    await expect(
      applyExcludeTransfersWithZip({
        ejoos,
        plan: planOf(op),
        ops: [op],
      }),
    ).rejects.toThrow(TimesheetTransferError);
    const excluded = ejoos.sheets.find((sheet) => /виключен/i.test(sheet.sheetName));
    const excludedText = (excluded?.rawRows ?? [])
      .flat()
      .map((cell) => String(cell ?? ""))
      .join(" ");
    expect(excludedText).not.toMatch(/ЛЕКАЙ/i);
  }, 30_000);
});
