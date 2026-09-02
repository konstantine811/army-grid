import { describe, expect, it } from "vitest";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
  type ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { applyConfirmedEjoosOps } from "./ejoosSyncApply";
import { parseExcluded } from "./ejoosLiveViews";
import {
  buildEjoosSyncPlan,
  createMovementKey,
  parseEjoosAbsents,
  parseEjoosShpo,
  type EjoosSyncOp,
} from "./ejoosSyncPlan";
import { DEFAULT_STATUS_RULES } from "./ejoosRules";
import { isWorkbookApplyOp } from "./ejoosPersonDiff";
import { personOpsBlockApply } from "./ejoosOpRequirements";

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

const sheetByName = (workbook: unknown, name: string) =>
  (workbook as { sheet: (sheetName: string) => unknown }).sheet(name) as {
    cell: (
      row: number,
      column: number,
    ) => { value: (value?: unknown) => unknown };
  };

const snapshotOf = async (
  blob: Blob,
  fileName: string,
): Promise<ExcelWorkbookSnapshot> =>
  readWorkbookSnapshot(
    new File([blob], fileName, { type: XLSX_MIME }),
    EJOOS_SYNC_READ_OPTIONS,
  );

const days25 = (mark: (day: number) => string | number) =>
  Array.from({ length: 25 }, (_, index) => mark(index + 1));

const marksOf = (days: Array<string | number>) => days.map(String);

type ArchivePeriod = {
  type: string;
  from: string;
  outDate: string;
  outOrder?: string;
  back?: string;
  backDate?: string;
  backOrder?: string;
  place?: string;
};

type AbsentRow = {
  ground: string;
  place?: string;
  departDate: string;
  returnDate?: string;
  index?: string;
};

type HopCase = {
  name: string;
  id: string;
  fromIndex: string;
  index: string;
  rank: string;
  title: string;
  nameRe: RegExp;
  shStatus: string;
  hop?: { previousIndex: string; nextIndex: string; order: string; date: string };
  hopProcessed?: boolean;
  excluded?: { destination: string; note: string } | null;
  archive: ArchivePeriod[];
  absents?: AbsentRow[];
  timesheetDays: Array<string | number>;
  timesheetPersonId?: string;
};

const writeArchive = (
  archive: {
    cell: (row: number, column: number) => { value: (value?: unknown) => unknown };
  },
  person: HopCase,
) => {
  const headers = [
    "№ з/п",
    "ID",
    "Прізвище",
    "Звання",
    "Вид вибуття",
    "З якої дати",
    "Куди вибув",
    "Номер наказу вибуття",
    "Дата наказу вибуття",
    "Планова дата",
    "Дата прибуття",
    "Дата наказу",
    "Номер наказу",
    "Займана посада",
  ];
  headers.forEach((header, index) => archive.cell(1, index + 1).value(header));
  person.archive.forEach((period, index) => {
    const row = 2 + index;
    archive.cell(row, 1).value(index + 1);
    archive.cell(row, 2).value(person.id);
    archive.cell(row, 3).value(person.name);
    archive.cell(row, 4).value(person.rank);
    archive.cell(row, 5).value(period.type);
    archive.cell(row, 6).value(period.from);
    archive.cell(row, 7).value(period.place || "");
    archive.cell(row, 8).value(period.outOrder || "");
    archive.cell(row, 9).value(period.outDate);
    archive.cell(row, 11).value(period.back || "");
    archive.cell(row, 12).value(period.backDate || "");
    archive.cell(row, 13).value(period.backOrder || "");
    archive.cell(row, 14).value(person.title);
  });
};

const buildCaseWorkbooks = async (person: HopCase) => {
  const XlsxPopulate = await loadPopulate();
  const pbWorkbook = await XlsxPopulate.fromBlankAsync();
  const sh = pbWorkbook.sheet(0);
  sh.name("sh");
  const archive = pbWorkbook.addSheet("archive");
  const ruh = pbWorkbook.addSheet("Рух");
  sh.cell(1, 1).value("ID");
  sh.cell(1, 2).value("ПІБ");
  sh.cell(1, 3).value("Звання");
  sh.cell(1, 4).value("Індекс посади");
  sh.cell(1, 5).value("Посада");
  sh.cell(1, 6).value("Статус");
  sh.cell(2, 1).value(person.id);
  sh.cell(2, 2).value(person.name);
  sh.cell(2, 3).value(person.rank);
  sh.cell(2, 4).value(person.index);
  sh.cell(2, 5).value(person.title);
  sh.cell(2, 6).value(person.shStatus);
  writeArchive(archive, person);

  const ruhHeaders = [
    "№",
    "Тип",
    "Статус",
    "",
    "ID",
    "Звання",
    "ПІБ",
    "Індекс попередній",
    "Індекс який",
    "Яка зміна",
    "Куди",
    "",
    "",
    "",
    "Підстава №",
    "Підстава дата",
    "",
    "Примітка",
    "Наказ",
    "Дата",
  ];
  ruhHeaders.forEach((header, index) => ruh.cell(1, index + 1).value(header));
  if (person.hop) {
    ruh.cell(2, 1).value("1");
    ruh.cell(2, 2).value("ПОСАДА");
    ruh.cell(2, 3).value("В СТРОЮ");
    ruh.cell(2, 5).value(person.id);
    ruh.cell(2, 6).value(person.rank);
    ruh.cell(2, 7).value(person.name);
    ruh.cell(2, 8).value(person.hop.previousIndex);
    ruh.cell(2, 9).value(person.hop.nextIndex);
    ruh.cell(2, 10).value(person.title);
    ruh.cell(2, 11).value("_5 1ПБ");
    ruh.cell(2, 19).value(person.hop.order);
    ruh.cell(2, 20).value(person.hop.date);
  }
  const pb = await snapshotOf(
    (await pbWorkbook.outputAsync("blob")) as Blob,
    "1ПБ_25082026.xlsx",
  );

  const ejoosWorkbook = await XlsxPopulate.fromBlankAsync();
  const shpo = ejoosWorkbook.sheet(0);
  shpo.name("1. ШПО");
  const oos = ejoosWorkbook.addSheet("2. ООС");
  const excluded = ejoosWorkbook.addSheet("3. Виключені");
  excluded.cell(1, 1).value("3. Виключені");
  ejoosWorkbook.addSheet("4. Тимчасово прибулі");
  const absents = ejoosWorkbook.addSheet("5. Тимчасово відсутні");
  absents.cell(1, 1).value("5. Тимчасово відсутні");
  const timesheet = ejoosWorkbook.addSheet("6. Табель");
  shpo.cell(1, 1).value("1. ШПО");
  shpo.cell(7, 1).value(person.index);
  shpo.cell(7, 6).value(person.rank);
  shpo.cell(7, 7).value(person.name);
  shpo.cell(7, 8).value(person.id);
  oos.cell(1, 1).value("2. ООС");
  oos.cell(6, 1).value(person.rank);
  oos.cell(6, 2).value(person.name);
  oos.cell(6, 3).value(person.id);
  oos.cell(6, 4).value(person.index);
  if (person.excluded) {
    excluded.cell(6, 1).value(person.rank);
    excluded.cell(6, 2).value(person.name);
    excluded.cell(6, 3).value(person.id);
    excluded.cell(6, 4).value(person.fromIndex);
    excluded.cell(6, 28).value(person.hop?.date || "11.08.2026");
    excluded.cell(6, 29).value(person.hop?.date || "11.08.2026");
    excluded.cell(6, 30).value(person.hop?.order || "232");
    excluded.cell(6, 31).value(person.excluded.destination);
    excluded.cell(6, 32).value(person.excluded.note);
  }
  (person.absents || []).forEach((row, index) => {
    const excelRow = 6 + index;
    absents.cell(excelRow, 1).value(person.rank);
    absents.cell(excelRow, 2).value(person.name);
    absents.cell(excelRow, 3).value(person.id);
    absents.cell(excelRow, 4).value(row.index || person.index);
    absents.cell(excelRow, 5).value(row.ground);
    absents.cell(excelRow, 6).value(row.place || "");
    absents.cell(excelRow, 7).value(row.departDate);
    if (row.returnDate) absents.cell(excelRow, 13).value(row.returnDate);
  });
  timesheet.cell(1, 1).value("6. Табель");
  timesheet.cell(2, 9).value("Серпень 2026 р.");
  timesheet.cell(7, 2).value(person.index);
  timesheet.cell(7, 6).value(person.rank);
  timesheet.cell(7, 7).value(person.name);
  timesheet.cell(7, 8).value(
    person.timesheetPersonId === undefined ? person.id : person.timesheetPersonId,
  );
  person.timesheetDays.forEach((mark, index) => {
    timesheet.cell(7, 9 + index).value(mark);
  });
  const ejoos = await snapshotOf(
    (await ejoosWorkbook.outputAsync("blob")) as Blob,
    "ЄЖООС_станом_на_25-08-2026.xlsx",
  );
  const processedMovementKeys =
    person.hop && person.hopProcessed
      ? [
          createMovementKey({
            personId: person.id,
            fullName: person.name,
            type: "ПОСАДА",
            orderDate: person.hop.date,
            orderNumber: person.hop.order,
            previousIndex: person.hop.previousIndex,
            nextIndex: person.hop.nextIndex,
          }),
        ]
      : [];
  return { ejoos, pb, processedMovementKeys };
};

const personOpsOf = (ops: EjoosSyncOp[], person: HopCase) =>
  ops.filter((op) => op.personId === person.id || person.nameRe.test(op.fullName));

const readTimesheetDays = async (blob: Blob, person: HopCase) => {
  const XlsxPopulate = await loadPopulate();
  const written = await XlsxPopulate.fromDataAsync(blob);
  const timesheet = sheetByName(written, "6. Табель");
  const name = String(timesheet.cell(7, 7).value() ?? "").trim();
  const id = String(timesheet.cell(7, 8).value() ?? "").trim();
  expect(id === person.id || person.nameRe.test(name)).toBe(true);
  return Array.from({ length: 25 }, (_, index) =>
    String(timesheet.cell(7, 9 + index).value() ?? "").trim(),
  );
};

const runRoundTrip = async (person: HopCase) => {
  const { ejoos, pb, processedMovementKeys } = await buildCaseWorkbooks(person);
  const plan = buildEjoosSyncPlan(ejoos, pb, {
    statusRules: DEFAULT_STATUS_RULES,
    processedMovementKeys,
  });
  const personOps = personOpsOf(plan.ops, person);
  const applyOps = personOps.filter(isWorkbookApplyOp);
  const kinds = personOps.map(
    (op) => `${op.kind}:${op.payload.type || ""}:${op.class}`,
  );
  expect(personOpsBlockApply(applyOps), kinds.join(" | ")).toBe(false);

  let afterDays = marksOf(person.timesheetDays);
  let after = ejoos;
  if (applyOps.length) {
    const { blob } = await applyConfirmedEjoosOps({ ejoos, plan, ops: applyOps });
    after = await snapshotOf(blob, "ejoos-after.xlsx");
    afterDays = await readTimesheetDays(blob, person);
  }
  const rebuilt = buildEjoosSyncPlan(after, pb, {
    statusRules: DEFAULT_STATUS_RULES,
    processedMovementKeys,
  });
  const remaining = personOpsOf(rebuilt.ops, person).filter(isWorkbookApplyOp);
  const excludedAfter = parseExcluded(
    after.sheets.find((sheet) => /виключ/i.test(sheet.sheetName)),
  ).filter((row) => row.personId === person.id);
  const absentsAfter = parseEjoosAbsents(
    after.sheets.find((sheet) => /тимчасов.*відсут/i.test(sheet.sheetName)),
  ).filter((row) => row.personId === person.id);
  return {
    personOps,
    kinds,
    beforeDays: marksOf(person.timesheetDays),
    afterDays,
    excludedAfter,
    absentsAfter,
    remaining,
  };
};

const siriachenko: HopCase = {
  name: "СІРЯЧЕНКО Микола Миколайович",
  id: "5693",
  fromIndex: "2110767",
  index: "2103316",
  rank: "солдат",
  title: "вогнеметник 2 піхотного відділення 1 піхотного батальйону",
  nameRe: /сіряченко/i,
  shStatus: "_5 1ПБ",
  hop: {
    previousIndex: "2110767",
    nextIndex: "2103316",
    order: "232",
    date: "11.08.2026",
  },
  hopProcessed: true,
  excluded: {
    destination: "вогнеметник 2 піхотного відділення 1 піхотного батальйону",
    note: "ПЕРЕВЕДЕННЯ _ 1 ПБ",
  },
  archive: [
    {
      type: "МЕДРОТА",
      from: "23.10.2025",
      outDate: "23.10.2025",
      outOrder: "314",
    },
    {
      type: "ПОРАНЕННЯ",
      from: "23.10.2025",
      outDate: "23.10.2025",
      outOrder: "314",
      back: "23.10.2025",
      backDate: "23.10.2025",
      backOrder: "314",
    },
    {
      type: "ВІДПУСТКА",
      from: "30.07.2026",
      outDate: "30.07.2026",
      outOrder: "210",
      back: "09.08.2026",
      backDate: "09.08.2026",
      backOrder: "220",
    },
  ],
  absents: [
    { ground: "МЕДРОТА", place: "медрота", departDate: "23.10.2025", index: "2110767" },
    {
      ground: "ВІДПУСТКА",
      place: "відпустка",
      departDate: "30.07.2026",
      returnDate: "09.08.2026",
    },
  ],
  timesheetDays: days25((day) => (day <= 8 ? "від" : day <= 24 ? "лік" : "+")),
};

const shibriaiev: HopCase = {
  name: "ШИБРЯЄВ Станіслав Ігорович",
  id: "11735",
  fromIndex: "2103725",
  index: "2103255",
  rank: "солдат",
  title: "номер обслуги кулеметного взводу 1 піхотної роти 1 піхотного батальйону",
  nameRe: /шибряєв/i,
  shStatus: "_5 1ПБ",
  hop: {
    previousIndex: "2103725",
    nextIndex: "2103255",
    order: "232",
    date: "11.08.2026",
  },
  hopProcessed: true,
  excluded: {
    destination:
      "номер обслуги кулеметного взводу 1 піхотної роти 1 піхотного батальйону",
    note: "ПЕРЕВЕДЕННЯ _ 1 ПБ",
  },
  archive: [],
  timesheetDays: days25((day) => (day <= 24 ? "+" : 19)),
};

describe("before / after: хибне Виключені і коди Табеля", () => {
  it("Сіряченко: до — Виключені + лік після відпустки; після — немає виключення, 09–25 +", async () => {
    const expectedAfter = marksOf(days25((day) => (day <= 8 ? "від" : "+")));
    expect(marksOf(siriachenko.timesheetDays).slice(0, 8).every((mark) => mark === "від")).toBe(
      true,
    );
    expect(marksOf(siriachenko.timesheetDays).slice(8, 24).every((mark) => mark === "лік")).toBe(
      true,
    );
    expect(expectedAfter.slice(8)).toEqual(Array(17).fill("+"));

    const result = await runRoundTrip(siriachenko);
    expect(result.beforeDays).toEqual(marksOf(siriachenko.timesheetDays));
    expect(result.personOps.some((op) => op.kind === "exclude_transfer")).toBe(false);
    expect(
      result.personOps.some((op) => Number(op.payload.clearExcludedExcelRow) === 6),
    ).toBe(true);
    expect(
      result.personOps.some(
        (op) =>
          op.kind === "absent_close" &&
          /медрот/i.test(op.before) &&
          op.payload.returnDate === "30.07.2026",
      ),
    ).toBe(true);
    const paint = result.personOps.find(
      (op) => op.payload.type === "PAINT_ARCHIVE" || /1–8:від|1-8:від/.test(op.after),
    );
    expect(paint?.payload.timesheetAbsenceSpans || "").toMatch(/1-8:від/);
    expect(paint?.payload.timesheetAbsenceSpans || "").not.toMatch(/лік/);

    expect(result.afterDays, result.kinds.join(" | ")).toEqual(expectedAfter);
    expect(result.excludedAfter).toHaveLength(0);
    const medrota = result.absentsAfter.find((row) => /медрот/i.test(row.ground));
    expect(medrota?.actualReturn).toBe("30.07.2026");
    const vacation = result.absentsAfter.find((row) => /відпуст/i.test(row.ground));
    expect(vacation?.actualReturn).toBe("09.08.2026");
    expect(
      result.remaining,
      result.remaining.map((op) => `${op.kind}:${op.after}`).join(" | "),
    ).toHaveLength(0);
  }, 30_000);

  it("Шибряєв: до — Виключені + 19 на 25-му; після — немає виключення, усі дні +", async () => {
    expect(marksOf(shibriaiev.timesheetDays)[24]).toBe("19");
    const expectedAfter = marksOf(days25(() => "+"));
    const result = await runRoundTrip(shibriaiev);
    expect(result.personOps.some((op) => op.kind === "exclude_transfer")).toBe(false);
    expect(
      result.personOps.some((op) => Number(op.payload.clearExcludedExcelRow) === 6),
    ).toBe(true);
    expect(result.afterDays, result.kinds.join(" | ")).toEqual(expectedAfter);
    expect(result.excludedAfter).toHaveLength(0);
    expect(result.remaining).toHaveLength(0);
  }, 30_000);

  it("справжнє ПЕРЕВ до іншої в/ч не прибираємо з Виключених", async () => {
    const result = await runRoundTrip({
      ...shibriaiev,
      id: "90001",
      name: "ЗОВНІШНІЙ Перевірочний",
      nameRe: /зовнішній/i,
      excluded: { destination: "в/ч А4784", note: "ПЕРЕВЕДЕННЯ" },
      timesheetDays: days25(() => "+"),
    });
    expect(
      result.personOps.some((op) => Number(op.payload.clearExcludedExcelRow) > 0),
    ).toBe(false);
    expect(result.excludedAfter.some((row) => row.personId === "90001")).toBe(true);
  }, 30_000);

  it("відкритий СЗЧ у sh лишається СЗЧ на 01–25, не стає +", async () => {
    const result = await runRoundTrip({
      name: "СЗЧЕНКО Тестовий",
      id: "90002",
      fromIndex: "2103001",
      index: "2103001",
      rank: "солдат",
      title: "стрілець",
      nameRe: /сзченко/i,
      shStatus: "СЗЧ",
      archive: [
        {
          type: "СЗЧ",
          from: "25.07.2026",
          outDate: "25.07.2026",
          outOrder: "200",
        },
      ],
      absents: [{ ground: "СЗЧ", departDate: "25.07.2026" }],
      timesheetDays: days25((day) => (day <= 24 ? "СЗЧ" : "+")),
    });
    expect(result.beforeDays[24]).toBe("+");
    expect(result.afterDays).toEqual(marksOf(days25(() => "СЗЧ")));
    expect(result.absentsAfter.some((row) => /сзч/i.test(row.ground) && !row.actualReturn)).toBe(
      true,
    );
    expect(result.remaining).toHaveLength(0);
  }, 30_000);

  it("лікування досі відкрите в sh — старий МЕДРОТА фарбує серпень лік, не +", async () => {
    const result = await runRoundTrip({
      name: "ЛІКУВАЛЬНИЙ Тестовий",
      id: "90003",
      fromIndex: "2103002",
      index: "2103002",
      rank: "солдат",
      title: "стрілець",
      nameRe: /лікувальн/i,
      shStatus: "МЕДРОТА",
      archive: [
        {
          type: "МЕДРОТА",
          from: "23.10.2025",
          outDate: "23.10.2025",
          outOrder: "314",
        },
      ],
      absents: [{ ground: "МЕДРОТА", place: "медрота", departDate: "23.10.2025" }],
      timesheetDays: days25(() => "лік"),
    });
    expect(result.afterDays).toEqual(marksOf(days25(() => "лік")));
    expect(
      result.personOps.some((op) => op.kind === "absent_close"),
    ).toBe(false);
    expect(result.absentsAfter.some((row) => /медрот/i.test(row.ground) && !row.actualReturn)).toBe(
      true,
    );
    expect(result.remaining).toHaveLength(0);
  }, 30_000);

  it("після відпустки справжнє серпневе лікування лишається лік, не затирається +", async () => {
    const before = days25((day) => (day <= 8 ? "від" : "+"));
    const expectedAfter = days25((day) => {
      if (day <= 8) return "від";
      if (day >= 15 && day <= 19) return "лік";
      return "+";
    });
    const result = await runRoundTrip({
      name: "ЛІКПІСЛЯВІД Тестовий",
      id: "90004",
      fromIndex: "2103003",
      index: "2103003",
      rank: "солдат",
      title: "стрілець",
      nameRe: /лікпіслявід/i,
      shStatus: "_5 1ПБ",
      archive: [
        {
          type: "ВІДПУСТКА",
          from: "30.07.2026",
          outDate: "30.07.2026",
          outOrder: "210",
          back: "09.08.2026",
          backDate: "09.08.2026",
          backOrder: "220",
        },
        {
          type: "ЛІКУВАННЯ",
          from: "15.08.2026",
          outDate: "15.08.2026",
          outOrder: "240",
          back: "20.08.2026",
          backDate: "20.08.2026",
          backOrder: "242",
        },
      ],
      absents: [
        {
          ground: "ВІДПУСТКА",
          departDate: "30.07.2026",
          returnDate: "09.08.2026",
        },
        {
          ground: "ЛІКУВАННЯ",
          departDate: "15.08.2026",
          returnDate: "20.08.2026",
        },
      ],
      timesheetDays: before,
    });
    expect(result.beforeDays.slice(14, 19).every((mark) => mark === "+")).toBe(true);
    expect(result.afterDays).toEqual(marksOf(expectedAfter));
    const paint = result.personOps.find((op) => op.payload.timesheetAbsenceSpans);
    expect(paint?.payload.timesheetAbsenceSpans || "").toMatch(/лік/);
    expect(result.remaining).toHaveLength(0);
  }, 30_000);

  it("КАН: + лише до 20 добиваємо до 25, навіть якщо в Табелі немає ID", async () => {
    const before = days25((day) => (day <= 20 ? "+" : ""));
    const expectedAfter = marksOf(days25(() => "+"));
    const result = await runRoundTrip({
      name: "КАН Володимир Володимирович",
      id: "5721",
      fromIndex: "2103800",
      index: "2103800",
      rank: "молодший сержант",
      title: "командир автомобільного відділення",
      nameRe: /кан\s+володимир/i,
      shStatus: "_5 1ПБ",
      timesheetPersonId: "",
      archive: [
        {
          type: "ВІДРЯДЖЕННЯ",
          from: "21.02.2026",
          outDate: "20.02.2026",
          outOrder: "52",
          back: "09.03.2026",
          backDate: "14.03.2026",
          backOrder: "78",
          place: "Східниця",
        },
      ],
      timesheetDays: before,
    });
    expect(result.beforeDays.slice(0, 20).every((mark) => mark === "+")).toBe(true);
    expect(result.beforeDays.slice(20)).toEqual(["", "", "", "", ""]);
    expect(
      result.personOps.some(
        (op) =>
          op.kind === "timesheet_day" &&
          (op.payload.timesheetCode === "+" || op.after === "+" || /\+/.test(op.after)),
      ),
    ).toBe(true);
    expect(result.personOps.some((op) => /порожн/i.test(op.why))).toBe(true);
    expect(result.afterDays, result.kinds.join(" | ")).toEqual(expectedAfter);
    expect(result.remaining).toHaveLength(0);
  }, 30_000);
});

describe("Гіріченко: повторне ПЕРЕВ не лишає два рядки у Виключених", () => {
  const person = {
    name: "ГІРІЧЕНКО Ігор Вікторович",
    id: "11301",
    index: "2103802",
    rank: "солдат",
    title:
      "оператора безпілотних літальних апаратів відділення ударних безпілотних авіаційних комплексів",
  };

  const writeRuhPerev = (
    ruh: {
      cell: (row: number, column: number) => { value: (value?: unknown) => unknown };
    },
  ) => {
    const headers = [
      "№",
      "Тип",
      "Статус",
      "",
      "ID",
      "Звання",
      "ПІБ",
      "Індекс попередній",
      "Індекс який",
      "Яка зміна",
      "Куди",
      "",
      "",
      "",
      "Підстава №",
      "Підстава дата",
      "",
      "Примітка",
      "Наказ",
      "Дата",
    ];
    headers.forEach((header, index) => ruh.cell(1, index + 1).value(header));
    ruh.cell(2, 1).value("88361");
    ruh.cell(2, 2).value("ПЕРЕВ");
    ruh.cell(2, 3).value("В СТРОЮ");
    ruh.cell(2, 5).value(person.id);
    ruh.cell(2, 6).value(person.rank);
    ruh.cell(2, 7).value(person.name);
    ruh.cell(2, 8).value(person.index);
    ruh.cell(2, 10).value(person.title);
    ruh.cell(2, 11).value("А7379");
    ruh.cell(2, 15).value("665-РС");
    ruh.cell(2, 16).value("03.08.2026");
    ruh.cell(2, 18).value("А7379");
    ruh.cell(2, 19).value("232");
    ruh.cell(2, 20).value("11.08.2026");
  };

  const writeOldExcluded = (
    excluded: {
      cell: (row: number, column: number) => { value: (value?: unknown) => unknown };
    },
    row: number,
    order: { number: string; date: string; dest: string; note: string },
  ) => {
    excluded.cell(row, 1).value(person.rank);
    excluded.cell(row, 2).value(person.name);
    excluded.cell(row, 3).value(person.id);
    excluded.cell(row, 4).value(person.index);
    excluded.cell(row, 28).value(order.date);
    excluded.cell(row, 29).value(order.date);
    excluded.cell(row, 30).value(order.number);
    excluded.cell(row, 31).value(order.dest);
    excluded.cell(row, 32).value(order.note);
  };

  it("прибирає квітневий рядок і лишає лише чинне серпневе ПЕРЕВ", async () => {
    const XlsxPopulate = await loadPopulate();
    const pbWorkbook = await XlsxPopulate.fromBlankAsync();
    const sh = pbWorkbook.sheet(0);
    sh.name("sh");
    pbWorkbook.addSheet("archive");
    const ruh = pbWorkbook.addSheet("Рух");
    sh.cell(1, 1).value("ID");
    sh.cell(1, 2).value("ПІБ");
    sh.cell(1, 3).value("Звання");
    sh.cell(1, 4).value("Індекс посади");
    sh.cell(1, 5).value("Посада");
    sh.cell(1, 6).value("Статус");
    writeRuhPerev(ruh);
    const pb = await snapshotOf(
      (await pbWorkbook.outputAsync("blob")) as Blob,
      "1ПБ_25082026.xlsx",
    );

    const ejoosWorkbook = await XlsxPopulate.fromBlankAsync();
    const shpo = ejoosWorkbook.sheet(0);
    shpo.name("1. ШПО");
    const oos = ejoosWorkbook.addSheet("2. ООС");
    const excluded = ejoosWorkbook.addSheet("3. Виключені");
    excluded.cell(1, 1).value("3. Виключені");
    ejoosWorkbook.addSheet("4. Тимчасово прибулі");
    ejoosWorkbook.addSheet("5. Тимчасово відсутні");
    const timesheet = ejoosWorkbook.addSheet("6. Табель");
    shpo.cell(1, 1).value("1. ШПО");
    shpo.cell(7, 1).value(person.index);
    shpo.cell(7, 6).value(person.rank);
    shpo.cell(7, 7).value(person.name);
    shpo.cell(7, 8).value(person.id);
    oos.cell(1, 1).value("2. ООС");
    oos.cell(6, 1).value(person.rank);
    oos.cell(6, 2).value(person.name);
    oos.cell(6, 3).value(person.id);
    oos.cell(6, 4).value(person.index);
    writeOldExcluded(excluded, 6, {
      number: "108",
      date: "13.04.2026",
      dest: "СТАРШОГО ВОДІЯ ГРАНАТОМЕТНОГО ВЗВОДУ",
      note: "ПЕРЕВЕДЕННЯ у 155",
    });
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(7, 6).value(person.rank);
    timesheet.cell(7, 7).value(person.name);
    timesheet.cell(7, 8).value(person.id);
    for (let day = 1; day <= 20; day += 1) {
      timesheet.cell(7, 8 + day).value("+");
    }
    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_25-08-2026.xlsx",
    );
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
    });
    const personOps = plan.ops.filter(
      (op) => op.personId === person.id || /гіріченко/i.test(op.fullName),
    );
    const exclude = personOps.find((op) => op.kind === "exclude_transfer");
    expect(exclude, personOps.map((op) => op.kind).join(" | ")).toBeTruthy();
    expect(Number(exclude?.payload.clearExcludedExcelRow)).toBe(6);
    expect(exclude?.payload.excludedExcelRow || "").toBe("");

    const applyOps = personOps.filter(isWorkbookApplyOp);
    expect(personOpsBlockApply(applyOps)).toBe(false);
    const { blob } = await applyConfirmedEjoosOps({ ejoos, plan, ops: applyOps });
    const after = await snapshotOf(blob, "ejoos-after.xlsx");
    const excludedAfter = parseExcluded(
      after.sheets.find((sheet) => /виключ/i.test(sheet.sheetName)),
    ).filter((row) => row.personId === person.id || /гіріченко/i.test(row.fullName));
    expect(excludedAfter).toHaveLength(1);
    expect(excludedAfter[0]?.orderNumber).toBe("232");
    expect(excludedAfter[0]?.orderDate).toMatch(/11\.08\.2026/);
    const shpoAfter = parseEjoosShpo(
      after.sheets.find((sheet) => /шпо/i.test(sheet.sheetName)),
    ).filter((row) => row.personId === person.id);
    expect(shpoAfter).toHaveLength(0);

    const rebuilt = buildEjoosSyncPlan(after, pb, {
      statusRules: DEFAULT_STATUS_RULES,
    });
    const remaining = rebuilt.ops.filter(
      (op) =>
        isWorkbookApplyOp(op) &&
        (op.personId === person.id || /гіріченко/i.test(op.fullName)),
    );
    expect(
      remaining,
      remaining.map((op) => `${op.kind}:${op.after}`).join(" | "),
    ).toHaveLength(0);
  }, 30_000);

  it("якщо вже два рядки і ПЕРЕВ проведене — прибирає лише старий дубль", async () => {
    const XlsxPopulate = await loadPopulate();
    const pbWorkbook = await XlsxPopulate.fromBlankAsync();
    const sh = pbWorkbook.sheet(0);
    sh.name("sh");
    pbWorkbook.addSheet("archive");
    const ruh = pbWorkbook.addSheet("Рух");
    sh.cell(1, 1).value("ID");
    sh.cell(1, 2).value("ПІБ");
    writeRuhPerev(ruh);
    const pb = await snapshotOf(
      (await pbWorkbook.outputAsync("blob")) as Blob,
      "1ПБ_25082026.xlsx",
    );

    const ejoosWorkbook = await XlsxPopulate.fromBlankAsync();
    const shpo = ejoosWorkbook.sheet(0);
    shpo.name("1. ШПО");
    ejoosWorkbook.addSheet("2. ООС").cell(1, 1).value("2. ООС");
    const excluded = ejoosWorkbook.addSheet("3. Виключені");
    excluded.cell(1, 1).value("3. Виключені");
    ejoosWorkbook.addSheet("4. Тимчасово прибулі");
    ejoosWorkbook.addSheet("5. Тимчасово відсутні");
    const timesheet = ejoosWorkbook.addSheet("6. Табель");
    shpo.cell(1, 1).value("1. ШПО");
    writeOldExcluded(excluded, 6, {
      number: "108",
      date: "13.04.2026",
      dest: "СТАРШОГО ВОДІЯ ГРАНАТОМЕТНОГО ВЗВОДУ",
      note: "ПЕРЕВЕДЕННЯ у 155",
    });
    writeOldExcluded(excluded, 7, {
      number: "232",
      date: "11.08.2026",
      dest: "А7379",
      note: "ПЕРЕВЕДЕННЯ",
    });
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(7, 6).value(person.rank);
    timesheet.cell(7, 7).value(person.name);
    timesheet.cell(7, 8).value(person.id);
    timesheet.cell(7, 8 + 11).value("вибув до А7379 від 11.08.2026");
    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_25-08-2026.xlsx",
    );
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
    });
    const personOps = plan.ops.filter(
      (op) => op.personId === person.id || /гіріченко/i.test(op.fullName),
    );
    expect(personOps.some((op) => op.kind === "exclude_transfer")).toBe(false);
    expect(
      personOps.some(
        (op) =>
          op.payload.type === "CLEAR_STALE_EXCLUSION_DUPLICATE" &&
          Number(op.payload.clearExcludedExcelRow) === 6,
      ),
    ).toBe(true);

    const applyOps = personOps.filter(isWorkbookApplyOp);
    const { blob } = await applyConfirmedEjoosOps({ ejoos, plan, ops: applyOps });
    const after = await snapshotOf(blob, "ejoos-after.xlsx");
    const excludedAfter = parseExcluded(
      after.sheets.find((sheet) => /виключ/i.test(sheet.sheetName)),
    ).filter((row) => row.personId === person.id || /гіріченко/i.test(row.fullName));
    expect(excludedAfter).toHaveLength(1);
    expect(excludedAfter[0]?.orderNumber).toBe("232");
  }, 30_000);
});
