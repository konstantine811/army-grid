import { describe, expect, it } from "vitest";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
  type ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { applyConfirmedEjoosOps } from "./ejoosSyncApply";
import { applyExcludeTransfersWithZip } from "./ejoosExcludeTransferZip";
import { loadTimesheetGridFromFile } from "./ejoosTimesheetPersonRows";
import { parseExcluded } from "./ejoosLiveViews";
import {
  buildEjoosSyncPlan,
  parseEjoosOos,
  parseEjoosShpo,
  planBlocksWorkbookApply,
  type EjoosSyncOp,
  type EjoosSyncPlan,
} from "./ejoosSyncPlan";
import { DEFAULT_STATUS_RULES } from "./ejoosRules";
import { isInformationalOp, isWorkbookApplyOp } from "./ejoosPersonDiff";
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

type ArchivePeriod = {
  type: string;
  from: string;
  place?: string;
  outOrder?: string;
  outDate: string;
  back: string;
  backDate: string;
  backOrder: string;
};

type CancelRestoreCase = {
  name: string;
  id: string;
  index: string;
  title: string;
  nameRe: RegExp;
  julyFromIndex?: string;
  existingOos?: boolean;
  vacantStaffTimesheet?: boolean;
  missingFromTimesheet?: boolean;
  expectedTimesheetDays?: string[];
  archive: ArchivePeriod[];
};

const buildPb = async (person: CancelRestoreCase) => {
  const XlsxPopulate = await loadPopulate();
  const workbook = await XlsxPopulate.fromBlankAsync();
  const sh = workbook.sheet(0);
  sh.name("sh");
  const archive = workbook.addSheet("archive");
  const ruh = workbook.addSheet("Рух");

  sh.cell(1, 1).value("ID");
  sh.cell(1, 2).value("ПІБ");
  sh.cell(1, 3).value("Звання");
  sh.cell(1, 4).value("Індекс посади");
  sh.cell(1, 5).value("Посада");
  sh.cell(1, 6).value("Статус");
  sh.cell(2, 1).value(person.id);
  sh.cell(2, 2).value(person.name);
  sh.cell(2, 3).value("солдат");
  sh.cell(2, 4).value(person.index);
  sh.cell(2, 5).value(person.title);
  sh.cell(2, 6).value("В СТРОЮ");

  archive.cell(1, 1).value("№ з/п");
  archive.cell(1, 2).value("ID");
  archive.cell(1, 3).value("Прізвище");
  archive.cell(1, 4).value("Звання");
  archive.cell(1, 5).value("Вид вибуття");
  archive.cell(1, 6).value("З якої дати");
  archive.cell(1, 7).value("Куди вибув");
  archive.cell(1, 8).value("Номер наказу вибуття");
  archive.cell(1, 9).value("Дата наказу вибуття");
  archive.cell(1, 10).value("Планова дата");
  archive.cell(1, 11).value("Дата прибуття");
  archive.cell(1, 12).value("Дата наказу");
  archive.cell(1, 13).value("Номер наказу");
  archive.cell(1, 14).value("Займана посада");
  person.archive.forEach((period, index) => {
    const row = 2 + index;
    archive.cell(row, 1).value(index + 1);
    archive.cell(row, 2).value(person.id);
    archive.cell(row, 3).value(person.name);
    archive.cell(row, 4).value("солдат");
    archive.cell(row, 5).value(period.type);
    archive.cell(row, 6).value(period.from);
    archive.cell(row, 7).value(period.place || "");
    archive.cell(row, 8).value(period.outOrder || "");
    archive.cell(row, 9).value(period.outDate);
    archive.cell(row, 11).value(period.back);
    archive.cell(row, 12).value(period.backDate);
    archive.cell(row, 13).value(period.backOrder);
    archive.cell(row, 14).value(person.title);
  });

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
  ruhHeaders.forEach((header, index) => {
    ruh.cell(1, index + 1).value(header);
  });
  const ruhRows: Array<Record<number, string | number>> = [
    ...(person.julyFromIndex
      ? [
          {
            1: 1,
            2: "РОЗПОРЯДЖ",
            5: person.id,
            6: "солдат",
            7: person.name,
            8: person.julyFromIndex,
            19: "191",
            20: "01.07.2026",
          },
        ]
      : []),
    {
      1: person.julyFromIndex ? 2 : 1,
      2: "ПОСАДА",
      5: person.id,
      6: "солдат",
      7: person.name,
      8: "РОЗПОРЯДЖЕННЯ",
      9: person.index,
      10: `РОЗПОРЯДЖЕННЯ → ${person.index} ${person.title}`,
      19: "223",
      20: "02.08.2026",
    },
    {
      1: person.julyFromIndex ? 3 : 2,
      2: "ПЕРЕВ",
      5: person.id,
      6: "солдат",
      7: person.name,
      8: person.index,
      11: "А0409",
      19: "225",
      20: "04.08.2026",
    },
    {
      1: person.julyFromIndex ? 4 : 3,
      2: "СКАСУВАННЯ",
      5: person.id,
      6: "солдат",
      7: person.name,
      9: person.index,
      11: "А0409",
      18: "скасування переведення до А0409",
      19: "228",
      20: "07.08.2026",
    },
  ];
  ruhRows.forEach((values, index) => {
    const row = 2 + index;
    for (const [column, value] of Object.entries(values)) {
      ruh.cell(row, Number(column)).value(value);
    }
  });

  const blob = (await workbook.outputAsync("blob")) as Blob;
  return snapshotOf(blob, "1ПБ_25082026.xlsx");
};

const buildStaleEjoos = async (person: CancelRestoreCase) => {
  const XlsxPopulate = await loadPopulate();
  const workbook = await XlsxPopulate.fromBlankAsync();
  const shpo = workbook.sheet(0);
  shpo.name("1. ШПО");
  const oos = workbook.addSheet("2. ООС");
  const excluded = workbook.addSheet("3. Виключені");
  workbook.addSheet("4. Тимчасово прибулі").cell(1, 1).value("4. Тимчасово прибулі");
  workbook.addSheet("5. Тимчасово відсутні").cell(1, 1).value("5. Тимчасово відсутні");
  const timesheet = workbook.addSheet("6. Табель");

  shpo.cell(1, 1).value("1. ШПО");
  shpo.cell(4, 1).value("індекс");
  shpo.cell(7, 1).value(person.index);

  oos.cell(1, 1).value("2. ООС");
  oos.cell(4, 2).value("ПІБ");
  oos.cell(6, 1).value("солдат");
  oos.cell(6, 2).value("НОВІКОВ Олександр Сергійович");
  oos.cell(6, 3).value("1");
  oos.cell(6, 4).value("2103378");
  if (person.existingOos) {
    oos.cell(7, 1).value("солдат");
    oos.cell(7, 2).value(person.name);
    oos.cell(7, 3).value(person.id);
    oos.cell(7, 4).value(person.index);
  }

  excluded.cell(1, 1).value("3. Виключені");
  excluded.cell(6, 1).value("солдат");
  excluded.cell(6, 2).value(person.name);
  excluded.cell(6, 3).value(Number(person.id));
  excluded.cell(6, 4).value(person.index);
  excluded.cell(6, 28).value("04.08.2026");
  excluded.cell(6, 29).value("04.08.2026");
  excluded.cell(6, 30).value("225");
  excluded.cell(6, 31).value("а0409");
  excluded.cell(6, 32).value("ПЕРЕВЕДЕННЯ");

  timesheet.cell(1, 1).value("6. Табель");
  timesheet.cell(2, 9).value("Серпень 2026 р.");
  timesheet.cell(4, 2).value("індекс");
  timesheet.cell(4, 7).value("ПІБ");
  timesheet.cell(7, 2).value(person.index);
  if (person.missingFromTimesheet) {
    // штатний рядок є, ПІБ немає — типовий стан після виключення
  } else if (person.vacantStaffTimesheet) {
    timesheet.cell(8, 6).value("солдат");
    timesheet.cell(8, 7).value(person.name);
    timesheet.cell(8, 8).value(person.id);
    timesheet.cell(8, 9).value("+");
    timesheet.cell(8, 10).value("+");
    timesheet.cell(8, 11).value("+");
    timesheet.cell(8, 12).value("вибув до А0409");
  } else {
    timesheet.cell(7, 6).value("солдат");
    timesheet.cell(7, 7).value(person.name);
    timesheet.cell(7, 8).value(person.id);
    timesheet.cell(7, 9).value("+");
    timesheet.cell(7, 10).value("+");
    timesheet.cell(7, 11).value("+");
    timesheet.cell(7, 12).value("вибув до А0409");
  }

  const blob = (await workbook.outputAsync("blob")) as Blob;
  return snapshotOf(blob, "ЄЖООС_станом_на_12-08-2026.xlsx");
};

const assertCancelRestoreRoundTrip = async (person: CancelRestoreCase) => {
  const [pb, ejoos] = await Promise.all([
    buildPb(person),
    buildStaleEjoos(person),
  ]);
  const plan = buildEjoosSyncPlan(ejoos, pb, {
    statusRules: DEFAULT_STATUS_RULES,
  });
  const ops = plan.ops.filter(
    (op) => op.personId === person.id || person.nameRe.test(op.fullName),
  );
  const kinds = ops.map(
    (op) => `${op.kind}:${op.payload.type || ""}:${op.class}`,
  );

  expect(plan.monthRolloverRequired).toBe(false);
  expect(plan.timesheetMonthHeaderUnknown).toBe(false);
  expect(plan.sourceDateUnknown).toBe(false);
  expect(ops.some((op) => op.kind === "exclude_transfer")).toBe(false);
  expect(
    ops.some(
      (op) =>
        op.kind === "other_manual" &&
        op.payload.type === "TRANSFER_CANCELLED" &&
        op.payload.reviewReason !== "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH",
    ),
  ).toBe(true);
  expect(ops.some((op) => op.kind === "position_change")).toBe(true);

  const applyOps = ops.filter(isWorkbookApplyOp);
  expect(personOpsBlockApply(applyOps), kinds.join(" | ")).toBe(false);

  const { blob } = await applyConfirmedEjoosOps({
    ejoos,
    plan,
    ops: applyOps,
  });
  const after = await snapshotOf(blob, "ejoos-after.xlsx");
  const shpo = parseEjoosShpo(
    after.sheets.find((sheet) => /шпо/i.test(sheet.sheetName)),
  );
  const oos = parseEjoosOos(
    after.sheets.find((sheet) => /оос/i.test(sheet.sheetName)),
  );
  const excluded = parseExcluded(
    after.sheets.find((sheet) => /виключ/i.test(sheet.sheetName)),
  );
  const XlsxPopulate = await loadPopulate();
  const written = await XlsxPopulate.fromDataAsync(blob);
  const timesheetSheet = sheetByName(written, "6. Табель");
  const timesheetHits: Array<{
    row: number;
    index: string;
    name: string;
    id: string;
    days: string[];
  }> = [];
  for (let row = 7; row <= 40; row += 1) {
    const name = String(timesheetSheet.cell(row, 7).value() ?? "").trim();
    const id = String(timesheetSheet.cell(row, 8).value() ?? "").trim();
    if (!name && !id) continue;
    if (id !== person.id && !person.nameRe.test(name)) continue;
    const days = Array.from({ length: 25 }, (_, index) =>
      String(timesheetSheet.cell(row, 9 + index).value() ?? "").trim(),
    );
    timesheetHits.push({
      row,
      index: String(timesheetSheet.cell(row, 2).value() ?? "").trim(),
      name,
      id,
      days,
    });
  }
  const activeTimesheet = timesheetHits.filter(
    (row) => !/вибув/i.test(row.days[3] || "") && row.days[24] !== "",
  );
  const slot = shpo.find((row) => row.positionIndex === person.index);
  const personOos = oos.filter((row) => row.personId === person.id);

  expect(slot?.fullName).toMatch(person.nameRe);
  expect(slot?.personId).toBe(person.id);
  expect(personOos, JSON.stringify(oos)).toHaveLength(1);
  expect(excluded.some((row) => row.personId === person.id)).toBe(false);
  expect(activeTimesheet, JSON.stringify(timesheetHits)).toHaveLength(1);
  expect(activeTimesheet[0]?.index).toBe(person.index);
  expect(activeTimesheet[0]?.days[24]).toBe("+");
  if (person.expectedTimesheetDays) {
    expect(activeTimesheet[0]?.days).toEqual(person.expectedTimesheetDays);
  }

  const rebuilt = buildEjoosSyncPlan(after, pb, {
    statusRules: DEFAULT_STATUS_RULES,
  });
  const remaining = rebuilt.ops.filter(
    (op) =>
      isWorkbookApplyOp(op) &&
      (op.personId === person.id || person.nameRe.test(op.fullName)),
  );
  expect(
    remaining,
    remaining
      .map((op) => `${op.kind}:${op.payload.type || ""}:${op.after}`)
      .join(" | "),
  ).toHaveLength(0);
};

const dobrovolskyi: CancelRestoreCase = {
  name: "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович",
  id: "12840",
  index: "2110786",
  title: "Водій зенітного ракетного відділення",
  nameRe: /добровольськ/i,
  julyFromIndex: "2103239",
  expectedTimesheetDays: [
    "-",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "лік",
    "лік",
    "лік",
    "лік",
    "лік",
    "лік",
    "лік",
    "лік",
    "СЗЧ",
    "СЗЧ",
    "+",
    "+",
    "+",
    "+",
    "+",
  ],
  archive: [
    {
      type: "ЛІКУВАННЯ",
      from: "11.08.2026",
      place: "заклад охорони здоров'я",
      outOrder: "233",
      outDate: "11.08.2026",
      back: "19.08.2026",
      backDate: "20.08.2026",
      backOrder: "241",
    },
    {
      type: "СЗЧ",
      from: "19.08.2026",
      place: "",
      outOrder: "",
      outDate: "19.08.2026",
      back: "21.08.2026",
      backDate: "21.08.2026",
      backOrder: "242",
    },
  ],
};

const yamkovyi: CancelRestoreCase = {
  name: "ЯМКОВИЙ Руслан Костянтинович",
  id: "14732",
  index: "2103435",
  title: "Стрілець",
  nameRe: /ямков/i,
  expectedTimesheetDays: [
    "-",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "+",
    "лік",
    "лік",
    "+",
    "+",
    "+",
    "+",
    "+",
  ],
  archive: [
    {
      type: "ЛІКУВАННЯ",
      from: "19.08.2026",
      place: "заклад охорони здоров'я",
      outOrder: "240",
      outDate: "19.08.2026",
      back: "21.08.2026",
      backDate: "21.08.2026",
      backOrder: "242",
    },
  ],
};

describe("August fixtures: cancelled transfer still in sh", () => {
  it("ДОБРОВОЛЬСЬКИЙ: ПОСАДА → ПЕРЕВ → СКАСУВАННЯ → лік → СЗЧ → повернення", async () => {
    await assertCancelRestoreRoundTrip(dobrovolskyi);
  }, 30_000);

  it("ЯМКОВИЙ: ПОСАДА → ПЕРЕВ → СКАСУВАННЯ → лікування → повернення", async () => {
    await assertCancelRestoreRoundTrip(yamkovyi);
  }, 30_000);

  it("ДОБРОВОЛЬСЬКИЙ: існуюча картка ООС не дублюється при restore", async () => {
    await assertCancelRestoreRoundTrip({ ...dobrovolskyi, existingOos: true });
  }, 30_000);

  it("ДОБРОВОЛЬСЬКИЙ: вакантний штатний Табель + історичний «вибув»", async () => {
    await assertCancelRestoreRoundTrip({
      ...dobrovolskyi,
      vacantStaffTimesheet: true,
    });
  }, 30_000);

  it("ДОБРОВОЛЬСЬКИЙ: у Табелі немає ПІБ — записати на штатний рядок", async () => {
    await assertCancelRestoreRoundTrip({
      ...dobrovolskyi,
      missingFromTimesheet: true,
    });
  }, 30_000);
});

describe("August fixtures: active staff row after processed position", () => {
  it("repairs an active row that still contains old departure text", async () => {
    const person = {
      name: "АТРАХОВ Олександр Анатолійович",
      id: "21155",
      index: "2103764",
      title: "оператор відділення радіоелектронної боротьби",
      rank: "солдат",
    };
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
    sh.cell(2, 6).value("В СТРОЮ");
    archive.cell(1, 1).value("№ з/п");
    archive.cell(1, 2).value("ID");
    archive.cell(1, 3).value("Прізвище");
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
    ruh.cell(2, 1).value("86733");
    ruh.cell(2, 2).value("ПОСАДА");
    ruh.cell(2, 3).value("ВІДСУТНІЙ в АРХІВІ");
    ruh.cell(2, 5).value(person.id);
    ruh.cell(2, 6).value(person.rank);
    ruh.cell(2, 7).value(person.name);
    ruh.cell(2, 8).value("4907559");
    ruh.cell(2, 9).value(person.index);
    ruh.cell(2, 10).value(`${person.title} 1 піхотного батальйону`);
    ruh.cell(2, 11).value("_5 1ПБ");
    ruh.cell(2, 15).value("296-РС");
    ruh.cell(2, 16).value("06.08.2026");
    ruh.cell(2, 19).value("228");
    ruh.cell(2, 20).value("07.08.2026");
    const pb = await snapshotOf(
      (await pbWorkbook.outputAsync("blob")) as Blob,
      "1ПБ_25082026.xlsx",
    );

    const ejoosWorkbook = await XlsxPopulate.fromBlankAsync();
    const shpo = ejoosWorkbook.sheet(0);
    shpo.name("1. ШПО");
    const oos = ejoosWorkbook.addSheet("2. ООС");
    ejoosWorkbook.addSheet("3. Виключені").cell(1, 1).value("3. Виключені");
    ejoosWorkbook.addSheet("4. Тимчасово прибулі").cell(1, 1).value("4. Тимчасово прибулі");
    ejoosWorkbook.addSheet("5. Тимчасово відсутні").cell(1, 1).value("5. Тимчасово відсутні");
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
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(7, 6).value(person.rank);
    timesheet.cell(7, 7).value(person.name);
    timesheet.cell(7, 8).value(person.id);
    for (let day = 1; day <= 4; day += 1) {
      timesheet.cell(7, 8 + day).value("+");
    }
    timesheet
      .cell(7, 13)
      .value("вибув у 2 піхотного батальйону від 05.08.2026");
    for (let day = 6; day <= 24; day += 1) {
      timesheet.cell(7, 8 + day).value("-");
    }
    timesheet.cell(7, 33).value("+");
    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_25-08-2026.xlsx",
    );

    const processedPositionKey =
      `id:${person.id}|посада|07.08.2026|228|4907559|${person.index}`;
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
      processedMovementKeys: [processedPositionKey],
    });
    const repair = plan.ops.find(
      (op) =>
        op.personId === person.id &&
        op.kind === "timesheet_day" &&
        op.payload.type === "PAINT_ARCHIVE",
    );
    expect(repair?.payload.timesheetActiveFrom).toBe("07.08.2026");

    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan,
      ops: repair ? [repair] : [],
    });
    const written = await XlsxPopulate.fromDataAsync(blob);
    const out = sheetByName(written, "6. Табель");
    const days = Array.from({ length: 25 }, (_, index) =>
      String(out.cell(7, 9 + index).value() ?? "").trim(),
    );
    expect(days).toEqual([
      "-",
      "-",
      "-",
      "-",
      "-",
      "-",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
    ]);
  }, 30_000);
});

describe("August fixtures: Atrakhov return after 05.08 leave", () => {
  it("does not exclude a current sh occupant and keeps both timesheet episodes", async () => {
    const XlsxPopulate = await loadPopulate();
    const person = {
      name: "АТРАХОВ Олександр Анатолійович",
      id: "21155",
      index: "2103764",
      rank: "солдат",
    };
    const pbWorkbook = await XlsxPopulate.fromBlankAsync();
    const sh = pbWorkbook.sheet(0);
    sh.name("sh");
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
    sh.cell(2, 5).value("оператор РЕБ");
    sh.cell(2, 6).value("В СТРОЮ");
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
    ruh.cell(2, 1).value("8001");
    ruh.cell(2, 2).value("ПЕРЕВ");
    ruh.cell(2, 5).value(person.id);
    ruh.cell(2, 6).value(person.rank);
    ruh.cell(2, 7).value(person.name);
    ruh.cell(2, 8).value(person.index);
    ruh.cell(2, 10).value("2 піхотного батальйону");
    ruh.cell(2, 11).value("А0409");
    ruh.cell(2, 19).value("200");
    ruh.cell(2, 20).value("05.08.2026");
    ruh.cell(3, 1).value("8002");
    ruh.cell(3, 2).value("ПОСАДА");
    ruh.cell(3, 5).value(person.id);
    ruh.cell(3, 6).value(person.rank);
    ruh.cell(3, 7).value(person.name);
    ruh.cell(3, 9).value(person.index);
    ruh.cell(3, 10).value("оператор РЕБ 1 піхотного батальйону");
    ruh.cell(3, 11).value("_5 1ПБ");
    ruh.cell(3, 19).value("228");
    ruh.cell(3, 20).value("07.08.2026");
    const pb = await snapshotOf(
      (await pbWorkbook.outputAsync("blob")) as Blob,
      "1ПБ_25082026.xlsx",
    );

    const ejoosWorkbook = await XlsxPopulate.fromBlankAsync();
    const shpo = ejoosWorkbook.sheet(0);
    shpo.name("1. ШПО");
    const oos = ejoosWorkbook.addSheet("2. ООС");
    oos.cell(1, 1).value("2. ООС");
    oos.cell(6, 1).value(person.rank);
    oos.cell(6, 2).value(person.name);
    oos.cell(6, 3).value(person.id);
    oos.cell(6, 4).value(person.index);
    const excluded = ejoosWorkbook.addSheet("3. Виключені");
    ejoosWorkbook.addSheet("4. Тимчасово прибулі");
    ejoosWorkbook.addSheet("5. Тимчасово відсутні");
    const timesheet = ejoosWorkbook.addSheet("6. Табель");
    shpo.cell(1, 1).value("1. ШПО");
    shpo.cell(4, 1).value("індекс");
    shpo.cell(7, 1).value(person.index);
    excluded.cell(1, 1).value("3. Виключені");
    excluded.cell(6, 2).value(person.name);
    excluded.cell(6, 3).value(person.id);
    excluded.cell(6, 31).value("2 піхотного батальйону");
    excluded.cell(6, 28).value("05.08.2026");
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(4, 2).value("індекс");
    timesheet.cell(4, 7).value("ПІБ");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(20, 2).value(person.index);
    timesheet.cell(20, 7).value(person.name);
    timesheet.cell(20, 8).value(person.id);
    for (let day = 1; day <= 4; day += 1) timesheet.cell(20, 8 + day).value("+");
    timesheet.cell(20, 13).value("вибув у 2 піхотного батальйону");
    for (let day = 6; day <= 25; day += 1) timesheet.cell(20, 8 + day).value("-");
    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_25-08-2026.xlsx",
    );
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
    });
    const personOps = plan.ops.filter((op) => op.personId === person.id);
    expect(
      personOps.filter((op) => op.kind === "exclude_transfer"),
    ).toHaveLength(0);
    expect(
      personOps.filter(
        (op) =>
          op.kind === "timesheet_day" &&
          Number(op.payload.excelRow || 0) === 20,
      ),
    ).toHaveLength(0);
    const place = personOps.find((op) => op.kind === "position_change");
    expect(place?.payload.timesheetActiveFrom).toMatch(/07\.08\.2026/);
    expect(place?.payload.timesheetExcelRow).toBe("7");
    expect(place?.payload.returningToStaffIndex).toBe("1");
    expect(
      Number(place?.payload.historyTimesheetExcelRow || 0) === 20 ||
        Number(place?.payload.previousTimesheetExcelRow || 0) === 20,
    ).toBe(true);

    const applyOps = personOps.filter(isWorkbookApplyOp);
    expect(personOpsBlockApply(applyOps)).toBe(false);
    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan,
      ops: applyOps,
    });
    const written = await XlsxPopulate.fromDataAsync(blob);
    const ts = sheetByName(written, "6. Табель");
    const daysAt = (row: number) =>
      Array.from({ length: 25 }, (_, index) =>
        String(ts.cell(row, 9 + index).value() ?? "").trim(),
      );
    expect(String(ts.cell(20, 7).value() ?? "")).toMatch(/атрахов/i);
    expect(daysAt(20).slice(0, 6)).toEqual([
      "+",
      "+",
      "+",
      "+",
      expect.stringMatching(/вибув/i),
      "-",
    ]);
    expect(String(ts.cell(7, 7).value() ?? "")).toMatch(/атрахов/i);
    expect(String(ts.cell(7, 2).value() ?? "")).toBe(person.index);
    expect(daysAt(7)).toEqual([
      "-",
      "-",
      "-",
      "-",
      "-",
      "-",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
    ]);
  }, 30_000);
});

describe("August fixtures: processed outbound movement", () => {
  it("does not offer the same external transfer again after it was applied", async () => {
    const XlsxPopulate = await loadPopulate();
    const person = {
      name: "ЯКОВЕНКО Ярослав Романович",
      id: "17560",
      index: "2103445",
    };

    const pbWorkbook = await XlsxPopulate.fromBlankAsync();
    const sh = pbWorkbook.sheet(0);
    sh.name("sh");
    const ruh = pbWorkbook.addSheet("Рух");
    sh.cell(1, 1).value("ID");
    sh.cell(1, 2).value("ПІБ");
    sh.cell(1, 3).value("Звання");
    sh.cell(1, 4).value("Індекс посади");
    sh.cell(1, 5).value("Посада");
    sh.cell(1, 6).value("Статус");
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
    ruh.cell(2, 1).value("9001");
    ruh.cell(2, 2).value("ПЕРЕВ");
    ruh.cell(2, 5).value(person.id);
    ruh.cell(2, 6).value("солдат");
    ruh.cell(2, 7).value(person.name);
    ruh.cell(2, 8).value(person.index);
    ruh.cell(2, 10).value("радіотелефоніст радіостанції");
    ruh.cell(2, 11).value("А4784");
    ruh.cell(2, 15).value("309-РС");
    ruh.cell(2, 16).value("14.08.2026");
    ruh.cell(2, 19).value("236");
    ruh.cell(2, 20).value("15.08.2026");
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
    excluded.cell(6, 2).value(person.name);
    excluded.cell(6, 3).value(person.id);
    excluded.cell(6, 4).value(person.index);
    excluded.cell(6, 30).value("236");
    excluded.cell(6, 29).value("15.08.2026");
    ejoosWorkbook.addSheet("4. Тимчасово прибулі").cell(1, 1).value("4. Тимчасово прибулі");
    ejoosWorkbook.addSheet("5. Тимчасово відсутні").cell(1, 1).value("5. Тимчасово відсутні");
    const timesheet = ejoosWorkbook.addSheet("6. Табель");
    shpo.cell(1, 1).value("1. ШПО");
    shpo.cell(7, 1).value("9999999");
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(7, 6).value("солдат");
    timesheet.cell(7, 7).value(person.name);
    timesheet.cell(7, 8).value(person.id);
    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_25-08-2026.xlsx",
    );

    const before = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
    });
    const transfer = before.ops.find(
      (op) => op.personId === person.id && op.kind === "exclude_transfer",
    );
    expect(transfer?.movementKey).toBeTruthy();

    const after = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
      processedMovementKeys: [transfer?.movementKey || ""],
    });
    expect(
      after.ops.some(
        (op) => op.personId === person.id && op.kind === "exclude_transfer",
      ),
    ).toBe(false);
  }, 30_000);
});

describe("August fixture: ОЛІЙНИК Нема в sh", () => {
  it("ПЕРЕВ → СКАСУВАННЯ без sh дає NEEDS_REVIEW і нуль workbook writes", async () => {
    const person: CancelRestoreCase = {
      name: "ОЛІЙНИК Андрій",
      id: "99901",
      index: "2103999",
      title: "Стрілець",
      nameRe: /олійник/i,
      archive: [],
    };
    const pb = await buildPb({ ...person, archive: [] });
    const XlsxPopulate = await loadPopulate();
    const workbook = await XlsxPopulate.fromDataAsync(pb.file);
    const shSheet = sheetByName(workbook, "sh");
    shSheet.cell(2, 1).value(null);
    shSheet.cell(2, 2).value(null);
    shSheet.cell(2, 3).value(null);
    shSheet.cell(2, 4).value(null);
    shSheet.cell(2, 5).value(null);
    shSheet.cell(2, 6).value("Нема в sh");
    const pbBlob = (await workbook.outputAsync("blob")) as Blob;
    const pbMissing = await snapshotOf(pbBlob, "1ПБ_25082026.xlsx");
    const ejoos = await buildStaleEjoos(person);
    const before = new Uint8Array(await ejoos.file.arrayBuffer());

    const plan = buildEjoosSyncPlan(ejoos, pbMissing, {
      statusRules: DEFAULT_STATUS_RULES,
    });
    const ops = plan.ops.filter(
      (op) => op.personId === person.id || person.nameRe.test(op.fullName),
    );
    const cancel = ops.find(
      (op) =>
        op.kind === "other_manual" && op.payload.type === "TRANSFER_CANCELLED",
    );
    expect(cancel?.payload.reviewReason).toBe(
      "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH",
    );
    expect(cancel && isInformationalOp(cancel)).toBe(true);
    expect(ops.filter(isWorkbookApplyOp)).toHaveLength(0);

    const applyOps = ops.filter(isWorkbookApplyOp);
    if (applyOps.length) {
      await applyConfirmedEjoosOps({ ejoos, plan, ops: applyOps });
    }
    const after = new Uint8Array(await ejoos.file.arrayBuffer());
    expect(after).toEqual(before);
  }, 30_000);
});

describe("August fixtures: Pochepetskyi outbound after rank", () => {
  it("moves the closed episode below staff 2103461 and keeps index / VOS / tariff", async () => {
    const XlsxPopulate = await loadPopulate();
    const person = {
      name: "ПОЧЕПЕЦЬКИЙ Олексій Олександрович",
      id: "16196",
      index: "2103461",
      rank: "солдат",
    };

    const pbWorkbook = await XlsxPopulate.fromBlankAsync();
    const sh = pbWorkbook.sheet(0);
    sh.name("sh");
    const archive = pbWorkbook.addSheet("archive");
    const ruh = pbWorkbook.addSheet("Рух");
    sh.cell(1, 1).value("ID");
    sh.cell(1, 2).value("ПІБ");
    archive.cell(1, 1).value("№ з/п");
    archive.cell(1, 2).value("ID");
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
    ruh.cell(2, 1).value("9101");
    ruh.cell(2, 2).value("ЗВАННЯ");
    ruh.cell(2, 5).value(person.id);
    ruh.cell(2, 6).value(person.rank);
    ruh.cell(2, 7).value(person.name);
    ruh.cell(2, 8).value(person.index);
    ruh.cell(2, 10).value("солдат → старший солдат");
    ruh.cell(2, 18).value("солдат → старший солдат");
    ruh.cell(2, 15).value("301-РС");
    ruh.cell(2, 16).value("08.08.2026");
    ruh.cell(2, 19).value("230");
    ruh.cell(2, 20).value("09.08.2026");
    ruh.cell(3, 1).value("9102");
    ruh.cell(3, 2).value("ПЕРЕВ");
    ruh.cell(3, 3).value("ВІДСУТНІЙ в АРХІВІ");
    ruh.cell(3, 5).value(person.id);
    ruh.cell(3, 6).value(person.rank);
    ruh.cell(3, 7).value(person.name);
    ruh.cell(3, 8).value(person.index);
    ruh.cell(3, 11).value("А7379");
    ruh.cell(3, 15).value("665-РС");
    ruh.cell(3, 16).value("03.08.2026");
    ruh.cell(3, 19).value("240");
    ruh.cell(3, 20).value("19.08.2026");
    const pb = await snapshotOf(
      (await pbWorkbook.outputAsync("blob")) as Blob,
      "1ПБ_25082026.xlsx",
    );

    const ejoosWorkbook = await XlsxPopulate.fromBlankAsync();
    const shpo = ejoosWorkbook.sheet(0);
    shpo.name("1. ШПО");
    const oos = ejoosWorkbook.addSheet("2. ООС");
    const excluded = ejoosWorkbook.addSheet("3. Виключені");
    ejoosWorkbook.addSheet("4. Тимчасово прибулі");
    ejoosWorkbook.addSheet("5. Тимчасово відсутні");
    const timesheet = ejoosWorkbook.addSheet("6. Табель");
    shpo.cell(1, 1).value("1. ШПО");
    shpo.cell(4, 1).value("індекс");
    shpo.cell(7, 1).value(person.index);
    shpo.cell(7, 6).value(person.rank);
    shpo.cell(7, 7).value(person.name);
    shpo.cell(7, 8).value(person.id);
    oos.cell(1, 1).value("2. ООС");
    oos.cell(6, 1).value(person.rank);
    oos.cell(6, 2).value(person.name);
    oos.cell(6, 3).value(person.id);
    oos.cell(6, 4).value(person.index);
    excluded.cell(1, 1).value("3. Виключені");
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(4, 2).value("індекс");
    timesheet.cell(4, 3).value("ВОС");
    timesheet.cell(4, 4).value("тарифний план");
    timesheet.cell(4, 7).value("ПІБ");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(7, 3).value("100");
    timesheet.cell(7, 4).value("7");
    timesheet.cell(8, 2).value("2103462");
    timesheet.cell(8, 7).value("СУСІД Роти");
    timesheet.cell(8, 8).value("99991");
    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_25-08-2026.xlsx",
    );
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
    });
    const personOps = plan.ops.filter((op) => op.personId === person.id);
    const exclude = personOps.find((op) => op.kind === "exclude_transfer");
    expect(exclude?.payload.fromRank).toMatch(/старший солдат/i);
    expect(exclude?.payload.timesheetExcelRow).toBe("7");
    expect(exclude?.payload.timesheetCreateHistory).not.toBe("1");
    expect(exclude?.payload.timesheetReplaceInPlace).not.toBe("1");
    expect(exclude?.payload.excludeDate).toMatch(/19\.08\.2026/);
    expect(
      personOps.filter((op) => op.kind === "absent_upsert"),
    ).toHaveLength(0);
    expect(
      personOps.some(
        (op) => op.payload.mismatchKind === "ARCHIVE_REFERENCE_MISSING",
      ),
    ).toBe(true);

    const applyOps = personOps.filter(isWorkbookApplyOp);
    expect(personOpsBlockApply(applyOps)).toBe(false);
    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan,
      ops: applyOps,
    });
    const written = await XlsxPopulate.fromDataAsync(blob);
    const ts = sheetByName(written, "6. Табель");
    const shpoSheet = sheetByName(written, "1. ШПО");
    expect(String(shpoSheet.cell(7, 7).value() ?? "")).toBe("");
    expect(String(ts.cell(7, 7).value() ?? "")).toBe("");
    expect(String(ts.cell(7, 6).value() ?? "")).toBe("");
    expect(String(ts.cell(7, 8).value() ?? "")).toBe("");
    expect(String(ts.cell(7, 2).value() ?? "")).toBe(person.index);
    expect(String(ts.cell(7, 3).value() ?? "")).toBe("100");
    expect(String(ts.cell(7, 4).value() ?? "")).toBe("7");
    expect(String(ts.cell(8, 7).value() ?? "")).toMatch(/сусід/i);
    expect(String(ts.cell(9, 7).value() ?? "")).toMatch(/почепецьк/i);
    expect(String(ts.cell(9, 6).value() ?? "")).toMatch(/старший солдат/i);
    expect(String(ts.cell(9, 2).value() ?? "")).toBe(person.index);
    expect(String(ts.cell(9, 3).value() ?? "")).toBe("100");
    expect(String(ts.cell(9, 4).value() ?? "")).toBe("7");
    const days = Array.from({ length: 25 }, (_, index) =>
      String(ts.cell(9, 9 + index).value() ?? "").trim(),
    );
    expect(days.slice(0, 18).every((mark) => mark === "+")).toBe(true);
    expect(days[18]).toMatch(/вибув у в\/ч А7379/i);
    expect(days.slice(19).every((mark) => mark === "-")).toBe(true);
    expect(days.filter((mark) => mark === "+")).toHaveLength(18);
    expect(
      Array.from({ length: 25 }, (_, index) =>
        String(ts.cell(7, 9 + index).value() ?? "").trim(),
      ).every((mark) => mark === ""),
    ).toBe(true);
  }, 30_000);
});

describe("1PB filename without a date", () => {
  it("does not treat 1ПБ_актуальний.xlsx as today and blocks apply", async () => {
    const [pbDated, ejoos] = await Promise.all([
      buildPb(dobrovolskyi),
      buildStaleEjoos(dobrovolskyi),
    ]);
    const pb = await snapshotOf(pbDated.file, "1ПБ_актуальний.xlsx");
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
    });
    expect(plan.sourceDateUnknown).toBe(true);
    expect(plan.timesheetDay).toBe(0);
    expect(planBlocksWorkbookApply(plan)).toBe(true);
  }, 30_000);
});

describe("month rollover: September 1PB vs August EJOOS", () => {
  it("does not rename the August header when 1PB is September", async () => {
    const [pbAugust, ejoos] = await Promise.all([
      buildPb(dobrovolskyi),
      buildStaleEjoos(dobrovolskyi),
    ]);
    const pb = await snapshotOf(pbAugust.file, "1ПБ_01092026.xlsx");
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
    });
    expect(plan.timesheetDay).toBe(1);
    expect(plan.timesheetDayLabel).toBe("01.09.2026");
    expect(plan.monthRolloverRequired).toBe(true);
    expect(planBlocksWorkbookApply(plan)).toBe(false);
    expect(plan.ejoosTimesheetMonthLabel).toMatch(/серпень/i);

    const applyOps = plan.ops.filter(isWorkbookApplyOp);
    if (applyOps.length) {
      const { blob } = await applyConfirmedEjoosOps({
        ejoos,
        plan,
        ops: applyOps,
      });
      const XlsxPopulate = await loadPopulate();
      const written = await XlsxPopulate.fromDataAsync(blob);
      const header = String(
        sheetByName(written, "6. Табель").cell(2, 9).value() ?? "",
      );
      expect(header).toMatch(/серпень/i);
      expect(header).not.toMatch(/вересень/i);
    }
  }, 30_000);

  it("does not block August apply when I2 has no month header", async () => {
    const [pb, ejoosBase] = await Promise.all([
      buildPb(dobrovolskyi),
      buildStaleEjoos(dobrovolskyi),
    ]);
    const XlsxPopulate = await loadPopulate();
    const workbook = await XlsxPopulate.fromDataAsync(ejoosBase.file);
    sheetByName(workbook, "6. Табель").cell(2, 9).value(null);
    const ejoos = await snapshotOf(
      (await workbook.outputAsync("blob")) as Blob,
      ejoosBase.fileName,
    );
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
    });
    expect(plan.timesheetMonthHeaderUnknown).toBe(true);
    expect(plan.monthRolloverRequired).toBe(false);
    expect(planBlocksWorkbookApply(plan)).toBe(false);
  }, 30_000);
});

describe("exclude timesheet copy stays in the company", () => {
  it("copies from mid-sheet staff even when the plan asked to createHistory near R7", async () => {
    const XlsxPopulate = await loadPopulate();
    const person = {
      name: "СОРОКОПУД Данило Андрійович",
      id: "17111",
      index: "2103510",
      rank: "солдат",
    };
    const workbook = await XlsxPopulate.fromBlankAsync();
    const shpo = workbook.sheet(0);
    shpo.name("1. ШПО");
    const oos = workbook.addSheet("2. ООС");
    workbook.addSheet("3. Виключені");
    workbook.addSheet("4. Тимчасово прибулі");
    workbook.addSheet("5. Тимчасово відсутні");
    const timesheet = workbook.addSheet("6. Табель");
    shpo.cell(7, 1).value(person.index);
    shpo.cell(7, 6).value(person.rank);
    shpo.cell(7, 7).value(person.name);
    shpo.cell(7, 8).value(person.id);
    oos.cell(6, 1).value(person.rank);
    oos.cell(6, 2).value(person.name);
    oos.cell(6, 3).value(person.id);
    oos.cell(6, 4).value(person.index);
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(7, 2).value("2103101");
    timesheet.cell(7, 7).value("ІНШИЙ Першої");
    timesheet.cell(7, 8).value("10001");
    timesheet.cell(8, 2).value("2103102");
    timesheet.cell(8, 7).value("ЩЕ ОДИН Першої");
    timesheet.cell(8, 8).value("10002");
    for (let row = 10; row <= 19; row += 1) {
      timesheet.cell(row, 2).value(String(2103400 + row));
      timesheet.cell(row, 7).value(`ЗАПОВНЕННЯ ${row}`);
      timesheet.cell(row, 8).value(String(18000 + row));
    }
    timesheet.cell(20, 2).value(person.index);
    timesheet.cell(20, 3).value("100");
    timesheet.cell(20, 4).value("7");
    timesheet.cell(20, 6).value(person.rank);
    timesheet.cell(20, 7).value(person.name);
    timesheet.cell(20, 8).value(person.id);
    timesheet.cell(21, 2).value("2103511");
    timesheet.cell(21, 7).value("СУСІД Роти");
    timesheet.cell(21, 8).value("19991");
    const ejoos = await snapshotOf(
      (await workbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_25-08-2026.xlsx",
    );
    const exclude: EjoosSyncOp = {
      id: "excl-sorokopud",
      kind: "exclude_transfer",
      class: "ready",
      sheet: "Виключені → Табель → ШПО/ООС",
      personId: person.id,
      fullName: person.name,
      positionIndex: person.index,
      rank: person.rank,
      before: "",
      after: "",
      sourceRef: "",
      why: "",
      confidence: "high",
      checkedDefault: true,
      payload: {
        destination: "А7379",
        timesheetDestination: "в/ч А7379",
        documentsDest: "А7379",
        excludeDate: "19.08.2026",
        orderNumber: "240",
        orderDate: "19.08.2026",
        fromRank: person.rank,
        fromName: person.name,
        fromPersonId: person.id,
        previousIndex: person.index,
        shpoExcelRow: "7",
        oosExcelRow: "6",
        timesheetCreateHistory: "1",
        timesheetReplaceInPlace: "",
        timesheetExcelRow: "",
      },
    };
    const plan: EjoosSyncPlan = {
      ejoosName: ejoos.fileName,
      pbName: "1ПБ_25082026.xlsx",
      timesheetDay: 25,
      timesheetDayLabel: "25.08.2026",
      ops: [exclude],
      summary: { ready: 1, needsInput: 0, conflict: 0 },
    };
    const blob = await applyExcludeTransfersWithZip({
      ejoos,
      plan,
      ops: [exclude],
    });
    const grid = await loadTimesheetGridFromFile(blob, "6. Табель");
    const cell = (row: number, column: number) =>
      String(grid[row - 1]?.[column - 1] ?? "").trim();
    expect(cell(7, 7)).toMatch(/інший/i);
    expect(cell(9, 7)).not.toMatch(/сорокопуд/i);
    expect(cell(20, 7)).toBe("");
    expect(cell(20, 8)).toBe("");
    expect(cell(21, 7)).toMatch(/сусід/i);
    expect(cell(22, 7)).toMatch(/сорокопуд/i);
    expect(cell(22, 6)).toMatch(/солдат/i);
  }, 30_000);
});

describe("internal position hop does not start a new timesheet episode", () => {
  it("СВІДОВСЬКИЙ: 2103225 → 2103157 → 2103225 keeps 01–25 +", async () => {
    const person = {
      name: "СВІДОВСЬКИЙ Олексій Вікторович",
      id: "7587",
      index: "2103225",
      midIndex: "2103157",
      rank: "молодший сержант",
    };
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
    sh.cell(2, 5).value("Командир відділення");
    sh.cell(2, 6).value("В СТРОЮ");
    archive.cell(1, 1).value("№ з/п");
    archive.cell(1, 2).value("ID");
    archive.cell(1, 3).value("Прізвище");
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
    const writeMove = (
      row: number,
      from: string,
      to: string,
      order: string,
      date: string,
    ) => {
      ruh.cell(row, 1).value(String(86000 + row));
      ruh.cell(row, 2).value("ПОСАДА");
      ruh.cell(row, 3).value("В СТРОЮ");
      ruh.cell(row, 5).value(person.id);
      ruh.cell(row, 6).value(person.rank);
      ruh.cell(row, 7).value(person.name);
      ruh.cell(row, 8).value(from);
      ruh.cell(row, 9).value(to);
      ruh.cell(row, 10).value(`${from} → ${to} 1 піхотного батальйону`);
      ruh.cell(row, 11).value("1ПБ");
      ruh.cell(row, 19).value(order);
      ruh.cell(row, 20).value(date);
    };
    writeMove(2, person.index, person.midIndex, "232", "11.08.2026");
    writeMove(3, person.midIndex, person.index, "234", "13.08.2026");
    const pb = await snapshotOf(
      (await pbWorkbook.outputAsync("blob")) as Blob,
      "1ПБ_25082026.xlsx",
    );

    const ejoosWorkbook = await XlsxPopulate.fromBlankAsync();
    const shpo = ejoosWorkbook.sheet(0);
    shpo.name("1. ШПО");
    const oos = ejoosWorkbook.addSheet("2. ООС");
    ejoosWorkbook.addSheet("3. Виключені").cell(1, 1).value("3. Виключені");
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
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(7, 6).value("солдат");
    timesheet.cell(7, 7).value(person.name);
    timesheet.cell(7, 8).value(person.id);
    for (let day = 1; day <= 12; day += 1) {
      timesheet.cell(7, 8 + day).value("-");
    }
    for (let day = 13; day <= 25; day += 1) {
      timesheet.cell(7, 8 + day).value("+");
    }
    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_25-08-2026.xlsx",
    );
    const processed = [
      `id:${person.id}|посада|11.08.2026|232|${person.index}|${person.midIndex}`,
      `id:${person.id}|посада|13.08.2026|234|${person.midIndex}|${person.index}`,
    ];
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
      processedMovementKeys: processed,
    });
    const paint = plan.ops.find(
      (op) =>
        op.personId === person.id &&
        op.kind === "timesheet_day" &&
        op.payload.type === "PAINT_ARCHIVE",
    );
    expect(paint?.payload.timesheetActiveFrom).toBe("01.08.2026");
    expect(paint?.after).not.toMatch(/до постановки/i);

    const { blob } = await applyConfirmedEjoosOps({
      ejoos,
      plan,
      ops: paint ? [paint] : [],
    });
    const written = await XlsxPopulate.fromDataAsync(blob);
    const out = sheetByName(written, "6. Табель");
    const days = Array.from({ length: 25 }, (_, index) =>
      String(out.cell(7, 9 + index).value() ?? "").trim(),
    );
    expect(days.every((mark) => mark === "+")).toBe(true);
  }, 30_000);
});

describe("internal position hop does not write Виключені", () => {
  it("ЛІСЮТІН already on 2103207: 2103535 → 2103207 stays out of Excluded", async () => {
    const person = {
      name: "ЛІСЮТІН Микола Олександрович",
      id: "19374",
      fromIndex: "2103535",
      index: "2103207",
      rank: "солдат",
    };
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
    sh.cell(2, 5).value("Старший стрілець");
    sh.cell(2, 6).value("В СТРОЮ");
    archive.cell(1, 1).value("№ з/п");
    archive.cell(1, 2).value("ID");
    archive.cell(1, 3).value("Прізвище");
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
    ruh.cell(2, 1).value("232");
    ruh.cell(2, 2).value("ПОСАДА");
    ruh.cell(2, 3).value("В СТРОЮ");
    ruh.cell(2, 5).value(person.id);
    ruh.cell(2, 6).value(person.rank);
    ruh.cell(2, 7).value(person.name);
    ruh.cell(2, 8).value(person.fromIndex);
    ruh.cell(2, 9).value(person.index);
    ruh.cell(2, 10).value(`${person.fromIndex} → ${person.index}`);
    ruh.cell(2, 11).value("1ПБ");
    ruh.cell(2, 19).value("232");
    ruh.cell(2, 20).value("11.08.2026");
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
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(7, 6).value(person.rank);
    timesheet.cell(7, 7).value(person.name);
    timesheet.cell(7, 8).value(person.id);
    for (let day = 1; day <= 25; day += 1) {
      timesheet.cell(7, 8 + day).value("+");
    }
    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_25-08-2026.xlsx",
    );
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
      processedMovementKeys: [],
    });
    const personOps = plan.ops.filter(
      (op) => op.personId === person.id || /лісютін/i.test(op.fullName),
    );
    expect(
      personOps.some(
        (op) =>
          op.kind === "exclude_transfer" ||
          (op.payload.closeOldPosition === "1" &&
            op.payload.internalStaffHop !== "1"),
      ),
    ).toBe(false);
    expect(
      personOps.filter((op) => op.kind === "position_change"),
    ).toHaveLength(0);
    expect(
      personOps.some((op) => op.payload.exclusionReason === "ПЕРЕВЕДЕННЯ 1 ПБ"),
    ).toBe(false);
  }, 30_000);
});

describe("БЕЗВІСТИ does not write Виключені", () => {
  const person = {
    name: "БОНДАРУК Андрій Анатолійович",
    id: "22801",
    index: "2103444",
    rank: "солдат",
  };

  const buildBondaruk = async (opts: { inSh: boolean }) => {
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
    if (opts.inSh) {
      sh.cell(2, 1).value(person.id);
      sh.cell(2, 2).value(person.name);
      sh.cell(2, 3).value(person.rank);
      sh.cell(2, 4).value(person.index);
      sh.cell(2, 5).value("Стрілець");
      sh.cell(2, 6).value("БЕЗВІСТИ");
    }
    archive.cell(1, 1).value("№ з/п");
    archive.cell(1, 2).value("ID");
    archive.cell(1, 3).value("Прізвище");
    archive.cell(1, 4).value("Звання");
    archive.cell(1, 5).value("Вид вибуття");
    archive.cell(1, 6).value("З якої дати");
    archive.cell(1, 7).value("Куди вибув");
    archive.cell(1, 8).value("Номер наказу вибуття");
    archive.cell(1, 9).value("Дата наказу вибуття");
    archive.cell(1, 11).value("Дата прибуття");
    archive.cell(2, 1).value(1);
    archive.cell(2, 2).value(person.id);
    archive.cell(2, 3).value(person.name);
    archive.cell(2, 4).value(person.rank);
    archive.cell(2, 5).value("БЕЗВІСТИ");
    archive.cell(2, 6).value("13.08.2026");
    archive.cell(2, 7).value("");
    archive.cell(2, 9).value("13.08.2026");
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
    ruh.cell(2, 1).value("210");
    ruh.cell(2, 2).value("ПЕРЕВ");
    ruh.cell(2, 3).value("В СТРОЮ");
    ruh.cell(2, 5).value(person.id);
    ruh.cell(2, 6).value(person.rank);
    ruh.cell(2, 7).value(person.name);
    ruh.cell(2, 8).value(person.index);
    ruh.cell(2, 11).value("А0409");
    ruh.cell(2, 19).value("210");
    ruh.cell(2, 20).value("04.08.2026");
    ruh.cell(3, 1).value("240");
    ruh.cell(3, 2).value("БЕЗВІСТИ");
    ruh.cell(3, 3).value("БЕЗВІСТИ");
    ruh.cell(3, 5).value(person.id);
    ruh.cell(3, 6).value(person.rank);
    ruh.cell(3, 7).value(person.name);
    ruh.cell(3, 8).value(person.index);
    ruh.cell(3, 11).value("БЕЗВІСТИ");
    ruh.cell(3, 19).value("240");
    ruh.cell(3, 20).value("13.08.2026");
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
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(7, 6).value(person.rank);
    timesheet.cell(7, 7).value(person.name);
    timesheet.cell(7, 8).value(person.id);
    for (let day = 1; day <= 25; day += 1) {
      timesheet.cell(7, 8 + day).value("+");
    }
    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_25-08-2026.xlsx",
    );
    return { ejoos, pb };
  };

  const expectNoExclude = (ops: EjoosSyncOp[]) => {
    const personOps = ops.filter(
      (op) => op.personId === person.id || /бондар/i.test(op.fullName),
    );
    expect(personOps.some((op) => op.kind === "exclude_transfer")).toBe(false);
    expect(
      personOps.some((op) => op.payload.closeOldPosition === "1"),
    ).toBe(false);
    expect(
      personOps.some(
        (op) =>
          op.kind === "absent_upsert" && /безвіст/i.test(op.payload.absenceType),
      ),
    ).toBe(true);
  };

  it("keeps Bondaruk on SHPO/OOS when sh still lists БЕЗВІСТИ from 13.08", async () => {
    const { ejoos, pb } = await buildBondaruk({ inSh: true });
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
      processedMovementKeys: [],
    });
    expectNoExclude(plan.ops);
  }, 30_000);

  it("does not fall back to an older ПЕРЕВ after БЕЗВІСТИ when he is missing from sh", async () => {
    const { ejoos, pb } = await buildBondaruk({ inSh: false });
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
      processedMovementKeys: [],
    });
    expectNoExclude(plan.ops);
  }, 30_000);
});
