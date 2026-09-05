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

  shpo.cell(9, 1).value("2103232");
  timesheet.cell(9, 2).value("2103232");
  timesheet.cell(9, 3).value("Водій-радіотелефоніст");
  timesheet.cell(9, 6).value("солдат");

  const blob = (await workbook.outputAsync("blob")) as Blob;
  return readWorkbookSnapshot(
    new File([blob], "ejoos.xlsx", { type: XLSX_MIME }),
    EJOOS_SYNC_READ_OPTIONS,
  );
};

describe("bulk apply re-reads workbook between batches", () => {
  it("applies multiple disposition families in one bulk queue", async () => {
    const ejoos = await fillCanonicalSheets();
    const dispositionOne = op({
      id: "disp-1",
      kind: "move_to_disposition",
      personId: "111",
      fullName: "ХУБАЄВ Іван",
      positionIndex: "2103764",
      rank: "солдат",
      payload: {
        destination: "у розпорядження командира",
        skipShpoDisposition: "1",
        orderDate: "05.08.2026",
        orderNumber: "100",
      },
    });
    const closeOne = op({
      id: "absent-close-1",
      kind: "absent_close",
      personId: "111",
      fullName: "ХУБАЄВ Іван",
      payload: {},
    });
    const dispositionTwo = op({
      id: "disp-2",
      kind: "move_to_disposition",
      personId: "222",
      fullName: "АТРАХОВ Петро",
      positionIndex: "2103764",
      rank: "солдат",
      payload: {
        destination: "у розпорядження командира",
        skipShpoDisposition: "1",
        orderDate: "06.08.2026",
        orderNumber: "101",
      },
    });
    const closeTwo = op({
      id: "absent-close-2",
      kind: "absent_close",
      personId: "222",
      fullName: "АТРАХОВ Петро",
      payload: {},
    });

    await expect(
      applyConfirmedEjoosOps({
        ejoos,
        plan: dummyPlan([dispositionOne, closeOne, dispositionTwo, closeTwo]),
        ops: [dispositionOne, closeOne, dispositionTwo, closeTwo],
      }),
    ).resolves.toMatchObject({ directXml: true });
  }, 30_000);

  it("ХУБАЄВ outbound + АТРАХОВ occupant on 2103764 keeps Атрахова in SHPO", async () => {
    const ejoos = await fillCanonicalSheets();
    const hubayevChangeText =
      "радіотелефоніст 2 штурмового взводу 2 штурмової роти";
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
        documentsDest: hubayevChangeText,
        timesheetDestination: "А4784",
        changeText: hubayevChangeText,
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
    const excludedSheet = after.sheets.find((sheet) =>
      /виключ/i.test(sheet.sheetName),
    );
    const excludedHubayevRow = excludedSheet?.rawRows.find((row) =>
      String(row[1] || "").includes("ХУБАЄВ"),
    );
    const slot = shpo.find((row) => row.positionIndex === "2103764");
    expect(slot?.fullName).toMatch(/АТРАХОВ/i);
    expect(slot?.personId).toBe("222");
    expect(oos.some((row) => /АТРАХОВ/i.test(row.fullName))).toBe(true);
    expect(oos.some((row) => /ХУБАЄВ/i.test(row.fullName))).toBe(false);
    expect(excluded.some((row) => /ХУБАЄВ/i.test(row.fullName))).toBe(true);
    expect(excludedHubayevRow?.[30]).toBe(hubayevChangeText);
  }, 30_000);
});

describe("position_change from СЗЧ does not clone the timesheet row", () => {
  it("paints the staff row in place and does not add a named copy below", async () => {
    const ejoos = await fillCanonicalSheets();
    const place = op({
      id: "pos-arushanyan",
      kind: "position_change",
      personId: "21692",
      fullName: "АРУШАНЯН Норайр Рубенович",
      positionIndex: "2103232",
      rank: "солдат",
      payload: {
        nextIndex: "2103232",
        nextName: "АРУШАНЯН Норайр Рубенович",
        nextPersonId: "21692",
        nextRank: "солдат",
        orderDate: "10.08.2026",
        orderNumber: "231",
        shpoExcelRow: "9",
        timesheetExcelRow: "9",
        timesheetActiveFrom: "10.08.2026",
        timesheetPreserveHistory: "1",
        returningFromDisposition: "1",
        timesheetSkipHistory: "1",
      },
    });
    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan: dummyPlan([place]),
      ops: [place],
    });
    const after = await readWorkbookSnapshot(
      new File([blob], "ЄЖООС_станом_на_25-08-2026.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const ts = after.sheets.find((sheet) => /табель/i.test(sheet.sheetName));
    const named = (ts?.rawRows ?? [])
      .map((row, index) => ({
        row: index + 1,
        name: String(row[6] ?? ""),
        id: String(row[7] ?? ""),
      }))
      .filter(
        (item) => /арушанян/i.test(item.name) || item.id === "21692",
      );
    expect(named.map((item) => item.row)).toEqual([9]);
    expect(String(ts?.rawRows[8]?.[8] ?? "").trim()).toBe("-");
    expect(String(ts?.rawRows[8]?.[17] ?? "").trim()).toBe("+");
  }, 30_000);
});

describe("СЗЧ → РОЗПОРЯДЖ then final sh occupant", () => {
  it("vacates DIDENKO then MAXIMENKO keeps 2103520; timesheet is one СЗЧ row", async () => {
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
    shpo.cell(7, 1).value("2103520");
    shpo.cell(7, 6).value("солдат");
    shpo.cell(7, 7).value("ДІДЕНКО Ілля Андрійович");
    shpo.cell(7, 8).value("11040");
    shpo.cell(20, 2).value(", який знаходиться у розпорядженні командира");

    oos.cell(4, 1).value("звання");
    oos.cell(4, 2).value("ПІБ");
    oos.cell(4, 3).value("ID");
    oos.cell(6, 1).value("солдат");
    oos.cell(6, 2).value("ДІДЕНКО Ілля Андрійович");
    oos.cell(6, 3).value("11040");
    oos.cell(6, 4).value("2103520");

    absents.cell(6, 2).value("ДІДЕНКО Ілля Андрійович");
    absents.cell(6, 5).value("СЗЧ");
    absents.cell(6, 7).value("13.07.2026");

    timesheet.cell(4, 2).value("індекс");
    timesheet.cell(4, 7).value("ПІБ");
    timesheet.cell(7, 2).value("2103520");
    timesheet.cell(7, 6).value("солдат");
    timesheet.cell(7, 7).value("ДІДЕНКО Ілля Андрійович");
    timesheet.cell(7, 8).value("22898");
    timesheet.cell(7, 9).value("СЗЧ");
    timesheet.cell(7, 18).value("СЗЧ");

    const blobIn = (await workbook.outputAsync("blob")) as Blob;
    const ejoos = await readWorkbookSnapshot(
      new File([blobIn], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );

    const didenko = op({
      id: "disp-didenko",
      kind: "move_to_disposition",
      personId: "11040",
      fullName: "ДІДЕНКО Ілля Андрійович",
      positionIndex: "2103520",
      rank: "солдат",
      payload: {
        previousIndex: "2103520",
        destination: "у розпорядження командира",
        orderDate: "10.08.2026",
        orderNumber: "231",
        shpoExcelRow: "7",
        timesheetExcelRow: "7",
        absenceExcelRow: "6",
        absenceCode: "СЗЧ",
        absenceType: "СЗЧ",
        remainsInOos: "true",
        timesheetFound: "true",
        keepOpenSzchTimesheet: "1",
        vacateTimesheetStaffSlot: "1",
      },
    });
    const maximenko = op({
      id: "pos-maximenko",
      kind: "position_change",
      personId: "9905",
      fullName: "МАКСИМЕНКО Олексій Євгенійович",
      positionIndex: "2103520",
      rank: "солдат",
      payload: {
        nextIndex: "2103520",
        nextName: "МАКСИМЕНКО Олексій Євгенійович",
        nextPersonId: "9905",
        nextRank: "солдат",
        orderDate: "11.08.2026",
        shpoExcelRow: "7",
        timesheetExcelRow: "7",
        timesheetActiveFrom: "11.08.2026",
      },
    });

    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan: dummyPlan([didenko, maximenko]),
      ops: [didenko, maximenko],
    });
    const after = await readWorkbookSnapshot(
      new File([blob], "ЄЖООС_станом_на_25-08-2026.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const shpoRows = parseEjoosShpo(
      after.sheets.find((sheet) => /шпо/i.test(sheet.sheetName)),
    );
    const oosRows = parseEjoosOos(
      after.sheets.find((sheet) => /оос/i.test(sheet.sheetName)),
    );
    const staff = shpoRows.find((row) => row.positionIndex === "2103520");
    expect(staff?.fullName).toMatch(/МАКСИМЕНКО/i);
    expect(staff?.personId).toBe("9905");
    expect(oosRows.some((row) => row.personId === "11040")).toBe(true);

    const ts = after.sheets.find((sheet) => /табель/i.test(sheet.sheetName));
    const didenkoRows = (ts?.rawRows ?? [])
      .map((row, index) => ({
        row: index + 1,
        name: String(row[6] ?? ""),
        id: String(row[7] ?? ""),
        days: Array.from({ length: 25 }, (_, day) =>
          String(row[8 + day] ?? "").trim(),
        ),
      }))
      .filter((item) => /діденко/i.test(item.name) || item.id === "11040");
    expect(didenkoRows).toHaveLength(1);
    expect(didenkoRows[0].id).toBe("11040");
    expect(didenkoRows[0].days.slice(0, 9).every((mark) => mark === "СЗЧ")).toBe(
      true,
    );
    expect(didenkoRows[0].days[9]).toBe(
      "вибув у розпорядження командира наказ №231 від 10.08.2026",
    );
    expect(didenkoRows[0].days.slice(10).every((mark) => mark === "-")).toBe(
      true,
    );
    const maximenkoRow = (ts?.rawRows ?? []).find(
      (row) => String(row[7] ?? "") === "9905",
    );
    const maximenkoDays = Array.from({ length: 25 }, (_, day) =>
      String(maximenkoRow?.[8 + day] ?? "").trim(),
    );
    expect(maximenkoDays.slice(0, 10).every((mark) => mark === "-")).toBe(true);
    expect(maximenkoDays.slice(10).every((mark) => mark === "+")).toBe(true);

    const abs = after.sheets.find((sheet) => /відсут/i.test(sheet.sheetName));
    expect(String(abs?.rawRows[5]?.[2] ?? "").trim()).toBe("11040");
  }, 30_000);

  it("appends DIDENKO into the ВИБУВ У РОЗПОРЯДЖЕННЯ subsection with name in column E", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const shpo = workbook.sheet(0);
    shpo.name("1. ШПО");
    workbook.addSheet("2. ООС").cell(1, 1).value("2. ООС");
    workbook.addSheet("3. Виключені").cell(1, 1).value("3. Виключені");
    workbook.addSheet("4. Тимчасово прибулі").cell(1, 1).value("4. Тимчасово прибулі");
    const absents = workbook.addSheet("5. Тимчасово відсутні");
    absents.cell(1, 1).value("5. Тимчасово відсутні");
    const timesheet = workbook.addSheet("6. Табель");
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    shpo.cell(7, 1).value("2103520");
    shpo.cell(7, 6).value("солдат");
    shpo.cell(7, 7).value("ДІДЕНКО Ілля Андрійович");
    shpo.cell(7, 8).value("11040");
    shpo.cell(20, 2).value(", який знаходиться у розпорядженні командира");
    absents.cell(6, 2).value("ДІДЕНКО Ілля Андрійович");
    absents.cell(6, 5).value("СЗЧ");
    absents.cell(6, 7).value("13.07.2026");
    timesheet.cell(7, 2).value("2103520");
    timesheet.cell(7, 6).value("солдат");
    timesheet.cell(7, 7).value("ДІДЕНКО Ілля Андрійович");
    timesheet.cell(7, 8).value("11040");
    for (let day = 1; day <= 9; day += 1) {
      timesheet.cell(7, 8 + day).value("СЗЧ");
    }
    timesheet
      .cell(15, 2)
      .value(
        "ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А 4862",
      );
    timesheet.cell(16, 1).value("РОЗПОРЯДЖЕННЯ");
    timesheet.cell(16, 3).value("СЗЧ");
    timesheet.cell(16, 4).value("солдат");
    timesheet.cell(16, 5).value("ГОРГУЦА Андрій Юрійович");

    const blobIn = (await workbook.outputAsync("blob")) as Blob;
    const ejoos = await readWorkbookSnapshot(
      new File([blobIn], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const didenko = op({
      id: "disp-didenko-section",
      kind: "move_to_disposition",
      personId: "11040",
      fullName: "ДІДЕНКО Ілля Андрійович",
      positionIndex: "2103520",
      rank: "солдат",
      payload: {
        previousIndex: "2103520",
        destination: "у розпорядження командира",
        orderDate: "10.08.2026",
        orderNumber: "231",
        shpoExcelRow: "7",
        timesheetExcelRow: "7",
        absenceCode: "СЗЧ",
        absenceType: "СЗЧ",
        remainsInOos: "true",
        timesheetFound: "true",
        keepOpenSzchTimesheet: "1",
        vacateTimesheetStaffSlot: "1",
      },
    });
    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan: dummyPlan([didenko]),
      ops: [didenko],
    });
    const after = await readWorkbookSnapshot(
      new File([blob], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const ts = after.sheets.find((sheet) => /табель/i.test(sheet.sheetName));
    const didenkoRow = (ts?.rawRows ?? []).find((row) =>
      row.some((cell) => /діденко/i.test(String(cell ?? ""))),
    );
    expect(didenkoRow).toBeDefined();
    expect(String(didenkoRow?.[0] ?? "")).toBe("РОЗПОРЯДЖЕННЯ");
    expect(String(didenkoRow?.[2] ?? "")).toBe("СЗЧ");
    expect(String(didenkoRow?.[4] ?? "")).toMatch(/діденко/i);
    expect(didenkoRow).not.toBe(ts?.rawRows?.[15]);
    expect(String(ts?.rawRows?.[15]?.[4] ?? "")).toMatch(/горгуц/i);
    expect(String(ts?.rawRows?.[6]?.[6] ?? "")).toBe("");
    expect(String(didenkoRow?.[8 + 10 - 1] ?? "")).toMatch(/вибув у розпорядження/i);
  }, 30_000);

  it("finds disposition block from full sheet XML when snapshot usedRange is shorter", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const shpo = workbook.sheet(0);
    shpo.name("1. ШПО");
    workbook.addSheet("2. ООС").cell(1, 1).value("2. ООС");
    workbook.addSheet("3. Виключені").cell(1, 1).value("3. Виключені");
    workbook.addSheet("4. Тимчасово прибулі").cell(1, 1).value("4. Тимчасово прибулі");
    const absents = workbook.addSheet("5. Тимчасово відсутні");
    absents.cell(1, 1).value("5. Тимчасово відсутні");
    const timesheet = workbook.addSheet("6. Табель");
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    shpo.cell(7, 1).value("2103454");
    shpo.cell(7, 6).value("солдат");
    shpo.cell(7, 7).value("ПІВКІН Станіслав Ігорович");
    shpo.cell(7, 8).value("14020");
    timesheet.cell(7, 2).value("2103454");
    timesheet.cell(7, 6).value("солдат");
    timesheet.cell(7, 7).value("ПІВКІН Станіслав Ігорович");
    timesheet.cell(7, 8).value("14020");
    for (let day = 1; day <= 20; day += 1) {
      timesheet.cell(7, 8 + day).value("+");
    }
    timesheet
      .cell(40, 2)
      .value(
        "ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А 4862",
      );
    timesheet.cell(41, 1).value("РОЗПОРЯДЖЕННЯ");
    timesheet.cell(41, 3).value("СЗЧ");
    timesheet.cell(41, 4).value("солдат");
    timesheet.cell(41, 5).value("ІСНУЮЧИЙ ЗАПИС");

    const blobIn = (await workbook.outputAsync("blob")) as Blob;
    const ejoos = await readWorkbookSnapshot(
      new File([blobIn], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const ts = ejoos.sheets.find((sheet) => /табель/i.test(sheet.sheetName));
    expect(ts).toBeTruthy();
    ts!.rawRows = ts!.rawRows.slice(0, 12);

    const pivkin = op({
      id: "disp-pivkin",
      kind: "move_to_disposition",
      personId: "14020",
      fullName: "ПІВКІН Станіслав Ігорович",
      positionIndex: "2103454",
      rank: "солдат",
      payload: {
        previousIndex: "2103454",
        destination: "у розпорядження командира",
        orderDate: "10.08.2026",
        orderNumber: "231",
        shpoExcelRow: "7",
        timesheetExcelRow: "7",
        absenceCode: "СЗЧ",
        absenceType: "СЗЧ",
        remainsInOos: "true",
        timesheetFound: "true",
        keepOpenSzchTimesheet: "1",
      },
    });
    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan: dummyPlan([pivkin]),
      ops: [pivkin],
    });
    const after = await readWorkbookSnapshot(
      new File([blob], "ejoos-out.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const outTs = after.sheets.find((sheet) => /табель/i.test(sheet.sheetName));
    const pivkinRow = (outTs?.rawRows ?? []).find((row) =>
      row.some((cell) => /півкін/i.test(String(cell ?? ""))),
    );
    expect(pivkinRow).toBeDefined();
    expect(String(pivkinRow?.[0] ?? "")).toBe("РОЗПОРЯДЖЕННЯ");
    expect(String(outTs?.rawRows?.[6]?.[6] ?? "")).toBe("");
  }, 30_000);

  it("appends DIDENKO at the end of disposition block (columns B/D/E/F like real Табель)", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const shpo = workbook.sheet(0);
    shpo.name("1. ШПО");
    workbook.addSheet("2. ООС").cell(1, 1).value("2. ООС");
    workbook.addSheet("3. Виключені").cell(1, 1).value("3. Виключені");
    workbook.addSheet("4. Тимчасово прибулі").cell(1, 1).value("4. Тимчасово прибулі");
    const absents = workbook.addSheet("5. Тимчасово відсутні");
    absents.cell(1, 1).value("5. Тимчасово відсутні");
    const timesheet = workbook.addSheet("6. Табель");
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    shpo.cell(7, 1).value("2103520");
    shpo.cell(7, 6).value("солдат");
    shpo.cell(7, 7).value("ДІДЕНКО Ілля Андрійович");
    shpo.cell(7, 8).value("11040");
    shpo.cell(20, 2).value(", який знаходиться у розпорядженні командира");
    absents.cell(6, 2).value("ДІДЕНКО Ілля Андрійович");
    absents.cell(6, 5).value("СЗЧ");
    absents.cell(6, 7).value("13.07.2026");
    timesheet.cell(7, 2).value("2103520");
    timesheet.cell(7, 6).value("солдат");
    timesheet.cell(7, 7).value("ДІДЕНКО Ілля Андрійович");
    timesheet.cell(7, 8).value("11040");
    for (let day = 1; day <= 9; day += 1) {
      timesheet.cell(7, 8 + day).value("СЗЧ");
    }
    timesheet
      .cell(15, 2)
      .value(
        "ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А 4862",
      );
    timesheet.cell(16, 1).value("РОЗПОРЯДЖЕННЯ");
    timesheet.cell(16, 3).value("Кулеметник");
    timesheet.cell(16, 4).value("СЗЧ");
    timesheet.cell(16, 6).value("солдат");
    timesheet.cell(16, 7).value("ГОРГУЦА Андрій Юрійович");
    timesheet.cell(17, 1).value("РОЗПОРЯДЖЕННЯ");
    timesheet.cell(17, 3).value("Розвідник");
    timesheet.cell(17, 4).value("ЗБ");
    timesheet.cell(17, 6).value("солдат");
    timesheet.cell(17, 7).value("ВОЛОС Руслан Вікторович");
    timesheet.cell(18, 1).value("РОЗПОРЯДЖЕННЯ");
    timesheet.cell(18, 6).value("солдат");
    timesheet.cell(18, 7).value("ШЕВЧЕНКО Юрій Сергійович");

    const blobIn = (await workbook.outputAsync("blob")) as Blob;
    const ejoos = await readWorkbookSnapshot(
      new File([blobIn], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const didenko = op({
      id: "disp-didenko-real-layout",
      kind: "move_to_disposition",
      personId: "11040",
      fullName: "ДІДЕНКО Ілля Андрійович",
      positionIndex: "2103520",
      rank: "солдат",
      payload: {
        previousIndex: "2103520",
        destination: "у розпорядження командира",
        orderDate: "10.08.2026",
        orderNumber: "231",
        shpoExcelRow: "7",
        timesheetExcelRow: "7",
        absenceCode: "СЗЧ",
        absenceType: "СЗЧ",
        remainsInOos: "true",
        timesheetFound: "true",
        keepOpenSzchTimesheet: "1",
        vacateTimesheetStaffSlot: "1",
      },
    });
    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan: dummyPlan([didenko]),
      ops: [didenko],
    });
    const after = await readWorkbookSnapshot(
      new File([blob], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const ts = after.sheets.find((sheet) => /табель/i.test(sheet.sheetName));
    expect(String(ts?.rawRows?.[18]?.[0] ?? "")).toBe("РОЗПОРЯДЖЕННЯ");
    expect(String(ts?.rawRows?.[18]?.[3] ?? "")).toBe("СЗЧ");
    expect(String(ts?.rawRows?.[18]?.[5] ?? "")).toBe("солдат");
    expect(String(ts?.rawRows?.[18]?.[6] ?? "")).toMatch(/діденко/i);
    expect(String(ts?.rawRows?.[17]?.[6] ?? "")).toMatch(/шевченко/i);
    expect(String(ts?.rawRows?.[15]?.[6] ?? "")).toMatch(/горгуц/i);
  }, 30_000);

  it("does not append into mid-sheet РОЗПОРЯДЖЕННЯ rows with staff index in column B", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const shpo = workbook.sheet(0);
    shpo.name("1. ШПО");
    workbook.addSheet("2. ООС").cell(1, 1).value("2. ООС");
    workbook.addSheet("3. Виключені").cell(1, 1).value("3. Виключені");
    workbook.addSheet("4. Тимчасово прибулі").cell(1, 1).value("4. Тимчасово прибулі");
    const absents = workbook.addSheet("5. Тимчасово відсутні");
    absents.cell(1, 1).value("5. Тимчасово відсутні");
    const timesheet = workbook.addSheet("6. Табель");
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    shpo.cell(7, 1).value("2103520");
    shpo.cell(7, 6).value("солдат");
    shpo.cell(7, 7).value("ДІДЕНКО Ілля Андрійович");
    shpo.cell(7, 8).value("11040");
    shpo.cell(20, 2).value(", який знаходиться у розпорядженні командира");
    absents.cell(6, 2).value("ДІДЕНКО Ілля Андрійович");
    absents.cell(6, 5).value("СЗЧ");
    absents.cell(6, 7).value("13.07.2026");
    timesheet.cell(7, 2).value("2103520");
    timesheet.cell(7, 6).value("солдат");
    timesheet.cell(7, 7).value("ДІДЕНКО Ілля Андрійович");
    timesheet.cell(7, 8).value("11040");
    for (let day = 1; day <= 9; day += 1) {
      timesheet.cell(7, 8 + day).value("СЗЧ");
    }
    timesheet.cell(10, 1).value("РОЗПОРЯДЖЕННЯ");
    timesheet.cell(10, 2).value("2103181");
    timesheet.cell(10, 4).value("солдат");
    timesheet.cell(10, 7).value("ТАФІ Іван Олегович");
    timesheet
      .cell(20, 2)
      .value(
        "ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А 4862",
      );
    timesheet.cell(21, 1).value("РОЗПОРЯДЖЕННЯ");
    timesheet.cell(21, 3).value("Водій");
    timesheet.cell(21, 4).value("СЗЧ");
    timesheet.cell(21, 6).value("солдат");
    timesheet.cell(21, 7).value("ГОРГУЦА Андрій Юрійович");
    timesheet.cell(22, 1).value("РОЗПОРЯДЖЕННЯ");
    timesheet.cell(22, 3).value("Розвідник");
    timesheet.cell(22, 4).value("ЗБ");
    timesheet.cell(22, 6).value("солдат");
    timesheet.cell(22, 7).value("ВОЛОС Руслан Вікторович");

    const blobIn = (await workbook.outputAsync("blob")) as Blob;
    const ejoos = await readWorkbookSnapshot(
      new File([blobIn], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const didenko = op({
      id: "disp-didenko-not-mid",
      kind: "move_to_disposition",
      personId: "11040",
      fullName: "ДІДЕНКО Ілля Андрійович",
      positionIndex: "2103520",
      rank: "солдат",
      payload: {
        previousIndex: "2103520",
        destination: "у розпорядження командира",
        orderDate: "10.08.2026",
        orderNumber: "231",
        shpoExcelRow: "7",
        timesheetExcelRow: "7",
        absenceCode: "СЗЧ",
        absenceType: "СЗЧ",
        remainsInOos: "true",
        timesheetFound: "true",
        keepOpenSzchTimesheet: "1",
        vacateTimesheetStaffSlot: "1",
      },
    });
    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan: dummyPlan([didenko]),
      ops: [didenko],
    });
    const after = await readWorkbookSnapshot(
      new File([blob], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const ts = after.sheets.find((sheet) => /табель/i.test(sheet.sheetName));
    expect(String(ts?.rawRows?.[9]?.[6] ?? "")).toMatch(/таф/i);
    expect(String(ts?.rawRows?.[9]?.[6] ?? "")).not.toMatch(/діденко/i);
    expect(String(ts?.rawRows?.[22]?.[0] ?? "")).toBe("РОЗПОРЯДЖЕННЯ");
    expect(String(ts?.rawRows?.[22]?.[6] ?? "")).toMatch(/діденко/i);
  }, 30_000);

  it("appends into Серпень subsection before the МАЙ marker row", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const shpo = workbook.sheet(0);
    shpo.name("1. ШПО");
    workbook.addSheet("2. ООС").cell(1, 1).value("2. ООС");
    workbook.addSheet("3. Виключені").cell(1, 1).value("3. Виключені");
    workbook.addSheet("4. Тимчасово прибулі").cell(1, 1).value("4. Тимчасово прибулі");
    const absents = workbook.addSheet("5. Тимчасово відсутні");
    absents.cell(1, 1).value("5. Тимчасово відсутні");
    const timesheet = workbook.addSheet("6. Табель");
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    shpo.cell(7, 1).value("2103520");
    shpo.cell(7, 6).value("солдат");
    shpo.cell(7, 7).value("ДІДЕНКО Ілля Андрійович");
    shpo.cell(7, 8).value("11040");
    shpo.cell(20, 2).value(", який знаходиться у розпорядженні командира");
    absents.cell(6, 2).value("ДІДЕНКО Ілля Андрійович");
    absents.cell(6, 5).value("СЗЧ");
    absents.cell(6, 7).value("13.07.2026");
    timesheet.cell(7, 2).value("2103520");
    timesheet.cell(7, 6).value("солдат");
    timesheet.cell(7, 7).value("ДІДЕНКО Ілля Андрійович");
    timesheet.cell(7, 8).value("11040");
    for (let day = 1; day <= 9; day += 1) {
      timesheet.cell(7, 8 + day).value("СЗЧ");
    }
    timesheet
      .cell(15, 2)
      .value(
        "ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А 4862",
      );
    timesheet.cell(16, 1).value("РОЗПОРЯДЖЕННЯ");
    timesheet.cell(16, 3).value("Кулеметник");
    timesheet.cell(16, 4).value("СЗЧ");
    timesheet.cell(16, 6).value("солдат");
    timesheet.cell(16, 7).value("ГУСАЧЕНКО Ігор Борисович");
    timesheet.cell(17, 3).value("Водій");
    timesheet.cell(17, 4).value("ЗБ");
    timesheet.cell(17, 6).value("солдат");
    timesheet.cell(17, 7).value("ОНОПА Володимир Юрійович");
    timesheet.cell(18, 7).value("МАЙ");
    timesheet.cell(19, 1).value("РОЗПОРЯДЖЕННЯ");
    timesheet.cell(19, 7).value("ЗАКАЛЮЖНИЙ Іван Олегович");

    const blobIn = (await workbook.outputAsync("blob")) as Blob;
    const ejoos = await readWorkbookSnapshot(
      new File([blobIn], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const didenko = op({
      id: "disp-didenko-august-slice",
      kind: "move_to_disposition",
      personId: "11040",
      fullName: "ДІДЕНКО Ілля Андрійович",
      positionIndex: "2103520",
      rank: "солдат",
      payload: {
        previousIndex: "2103520",
        destination: "у розпорядження командира",
        orderDate: "10.08.2026",
        orderNumber: "231",
        shpoExcelRow: "7",
        timesheetExcelRow: "7",
        absenceCode: "СЗЧ",
        absenceType: "СЗЧ",
        remainsInOos: "true",
        timesheetFound: "true",
        keepOpenSzchTimesheet: "1",
        vacateTimesheetStaffSlot: "1",
      },
    });
    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan: dummyPlan([didenko]),
      ops: [didenko],
    });
    const after = await readWorkbookSnapshot(
      new File([blob], "ejoos.xlsx", { type: XLSX_MIME }),
      EJOOS_SYNC_READ_OPTIONS,
    );
    const ts = after.sheets.find((sheet) => /табель/i.test(sheet.sheetName));
    expect(String(ts?.rawRows?.[17]?.[6] ?? "")).toMatch(/діденко/i);
    expect(String(ts?.rawRows?.[18]?.[6] ?? "")).toMatch(/закалюж/i);
  }, 30_000);
});
