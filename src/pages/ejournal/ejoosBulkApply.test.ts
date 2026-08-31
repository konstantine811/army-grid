import { describe, expect, it } from "vitest";
import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../../excelRoundTrip";
import { applyConfirmedEjoosOps } from "./ejoosSyncApply";
import { parseExcluded } from "./ejoosLiveViews";
import { parseEjoosOos, parseEjoosShpo, type EjoosSyncOp, type EjoosSyncPlan } from "./ejoosSyncPlan";
import {
  acceptAllReady,
  isWorkbookApplyOp,
  personCanEnterApplyQueue,
  personHasWorkbookApplyOps,
  type PersonChange,
} from "./ejoosPersonDiff";

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

const stubEjoos = (): ExcelWorkbookSnapshot => ({
  file: new File([new Uint8Array()], "ejoos.xlsx", { type: XLSX_MIME }),
  fileName: "ejoos.xlsx",
  sheetName: "1. ШПО",
  headerRows: [],
  rows: [],
  columnCount: 0,
  columnIndexes: [],
  dataStartRow: 1,
  sheets: [],
});

const op = (
  partial: Partial<EjoosSyncOp> & Pick<EjoosSyncOp, "kind">,
): EjoosSyncOp => ({
  id: "op-1",
  class: "ready",
  sheet: "тест",
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
  ...partial,
});

const dummyPlan = (ops: EjoosSyncOp[] = []): EjoosSyncPlan => ({
  ejoosName: "ejoos.xlsx",
  pbName: "1pb.xlsx",
  timesheetDay: 25,
  timesheetDayLabel: "25.08.2026",
  ops,
  summary: { ready: ops.length, needsInput: 0, conflict: 0 },
});

const person = (
  partial: Partial<PersonChange> & Pick<PersonChange, "ops" | "severity">,
): PersonChange =>
  ({
    id: "p1",
    personId: "1",
    fullName: "ТЕСТ",
    rank: "",
    positionIndex: "",
    category: "status",
    summaryBefore: "",
    summaryAfter: "",
    ejoosWillDo: [],
    sheetActions: [],
    sheetImpacts: [],
    sourceInfluences: [],
    timesheetPreview: null,
    decision: "pending",
    ...partial,
  }) as PersonChange;

describe("personCanEnterApplyQueue / acceptAllReady", () => {
  it("does not queue standalone ПРИБУВ even when ready", () => {
    const arrival = person({
      severity: "ready",
      ops: [op({ kind: "arrival", class: "ready" })],
    });
    expect(isWorkbookApplyOp(arrival.ops[0])).toBe(false);
    expect(personHasWorkbookApplyOps(arrival.ops)).toBe(false);
    expect(personCanEnterApplyQueue(arrival)).toBe(false);
    const session = {
      plan: dummyPlan(arrival.ops),
      people: [arrival],
      counters: {
        oosLike: 0,
        onDuty: 0,
        changes: 1,
        newcomers: 0,
        errors: 0,
        autoReady: 1,
        needsReview: 0,
      },
      pbFileName: "1pb.xlsx",
      analyzedAt: "",
    };
    expect(acceptAllReady(session).people[0].decision).toBe("pending");
  });

  it("queues a ready exclude_transfer", () => {
    const exclude = person({
      severity: "ready",
      ops: [
        op({
          kind: "exclude_transfer",
          payload: {
            destination: "А4784",
            excludeDate: "05.08.2026",
            orderNumber: "1",
            orderDate: "05.08.2026",
          },
        }),
      ],
    });
    expect(personCanEnterApplyQueue(exclude)).toBe(true);
  });
});

describe("applyConfirmedEjoosOps fail-closed", () => {
  it("throws on timesheet_day without excelRow", async () => {
    const ejoos = stubEjoos();
    await expect(
      applyConfirmedEjoosOps({
        ejoos,
        plan: dummyPlan(),
        ops: [
          op({
            kind: "timesheet_day",
            payload: { timesheetCode: "СЗЧ", day: "10" },
          }),
        ],
      }),
    ).rejects.toThrow("не мають усіх даних для безпечного застосування");
  });

  it("throws on standalone arrival instead of writing", async () => {
    const ejoos = stubEjoos();
    await expect(
      applyConfirmedEjoosOps({
        ejoos,
        plan: dummyPlan(),
        ops: [op({ kind: "arrival" })],
      }),
    ).rejects.toThrow("не має apply");
  });
});

const fillCanonicalSheets = async () => {
  const module = await import(
    "xlsx-populate/browser/xlsx-populate-no-encryption"
  );
  const workbook = await module.default.fromBlankAsync();
  const shpo = workbook.sheet(0);
  shpo.name("1. ШПО");
  const oos = workbook.addSheet("2. ООС");
  const excluded = workbook.addSheet("3. Виключені");
  const arrivals = workbook.addSheet("4. Тимчасово прибулі");
  const absents = workbook.addSheet("5. Тимчасово відсутні");
  const timesheet = workbook.addSheet("6. Табель");

  shpo.cell(1, 1).value("1. ШПО");
  oos.cell(1, 1).value("2. ООС");
  excluded.cell(1, 1).value("3. Виключені");
  arrivals.cell(1, 1).value("4. Тимчасово прибулі");
  absents.cell(1, 1).value("5. Тимчасово відсутні");
  timesheet.cell(1, 1).value("6. Табель");
  timesheet.cell(2, 9).value("Серпень 2026 р.");

  shpo.cell(4, 1).value("індекс");
  shpo.cell(4, 6).value("звання");
  shpo.cell(4, 7).value("ПІБ");
  shpo.cell(4, 8).value("ID");
  shpo.cell(7, 1).value("2103764");
  shpo.cell(7, 6).value("солдат");
  shpo.cell(7, 7).value("ХУБАЄВ Іван");
  shpo.cell(7, 8).value("111");

  oos.cell(4, 1).value("звання");
  oos.cell(4, 2).value("ПІБ");
  oos.cell(4, 3).value("ID");
  oos.cell(6, 1).value("солдат");
  oos.cell(6, 2).value("ХУБАЄВ Іван");
  oos.cell(6, 3).value("111");
  oos.cell(6, 4).value("2103764");
  oos.cell(7, 1).value("солдат");
  oos.cell(7, 2).value("АТРАХОВ Петро");
  oos.cell(7, 3).value("222");
  oos.cell(7, 4).value("2103764");

  excluded.cell(4, 2).value("ПІБ");
  excluded.cell(4, 31).value("куди вибув");

  timesheet.cell(4, 2).value("індекс");
  timesheet.cell(4, 7).value("ПІБ");
  timesheet.cell(7, 2).value("2103764");
  timesheet.cell(7, 6).value("солдат");
  timesheet.cell(7, 7).value("ХУБАЄВ Іван");
  timesheet.cell(7, 8).value("111");
  timesheet.cell(7, 9).value("+");

  const blob = (await workbook.outputAsync("blob")) as Blob;
  return readWorkbookSnapshot(
    new File([blob], "ejoos.xlsx", { type: XLSX_MIME }),
    EJOOS_SYNC_READ_OPTIONS,
  );
};

describe("bulk apply re-reads workbook between batches", () => {
  it("ХУБАЄВ outbound + АТРАХОВ occupant on 2103764 keeps Атрахова in SHPO", async () => {
    const ejoos = await fillCanonicalSheets();
    const hubayev = op({
      id: "excl-hubayev",
      kind: "exclude_transfer",
      personId: "111",
      fullName: "ХУБАЄВ Іван",
      positionIndex: "2103764",
      rank: "солдат",
      payload: {
        fromRank: "солдат",
        fromName: "ХУБАЄВ Іван",
        fromPersonId: "111",
        fromPositionIndex: "2103764",
        destination: "А4784",
        documentsDest: "А4784",
        timesheetDestination: "А4784",
        excludeDate: "05.08.2026",
        orderDate: "05.08.2026",
        orderNumber: "100",
        type: "ПЕРЕВ",
        exclusionReason: "ПЕРЕВЕДЕННЯ",
        shpoExcelRow: "7",
        oosExcelRow: "6",
        timesheetExcelRow: "7",
      },
    });
    const atrakhov = op({
      id: "occ-atrakhov",
      kind: "shpo_occupant",
      personId: "222",
      fullName: "АТРАХОВ Петро",
      positionIndex: "2103764",
      rank: "солдат",
      payload: {
        shpoExcelRow: "7",
        timesheetExcelRow: "7",
        nextName: "АТРАХОВ Петро",
        nextPersonId: "222",
        nextRank: "солдат",
      },
    });

    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan: dummyPlan([atrakhov, hubayev]),
      ops: [atrakhov, hubayev],
    });
    const after = await readWorkbookSnapshot(
      new File([blob], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const shpo = parseEjoosShpo(
      after.sheets.find((sheet) => /шпо/i.test(sheet.sheetName)),
    );
    const oos = parseEjoosOos(
      after.sheets.find((sheet) => /оос/i.test(sheet.sheetName)),
    );
    const excluded = parseExcluded(
      after.sheets.find((sheet) => /виключ/i.test(sheet.sheetName)),
    );
    const slot = shpo.find((row) => row.positionIndex === "2103764");
    expect(slot?.fullName).toMatch(/АТРАХОВ/i);
    expect(slot?.personId).toBe("222");
    expect(oos.some((row) => /АТРАХОВ/i.test(row.fullName))).toBe(true);
    expect(oos.some((row) => /ХУБАЄВ/i.test(row.fullName))).toBe(false);
    expect(excluded.some((row) => /ХУБАЄВ/i.test(row.fullName))).toBe(true);
  }, 30_000);
});
