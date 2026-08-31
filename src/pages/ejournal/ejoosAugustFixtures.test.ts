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
  parseEjoosOos,
  parseEjoosShpo,
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
  timesheet.cell(7, 6).value("солдат");
  timesheet.cell(7, 7).value(person.name);
  timesheet.cell(7, 8).value(person.id);
  timesheet.cell(7, 9).value("+");
  timesheet.cell(7, 10).value("+");
  timesheet.cell(7, 11).value("+");
  timesheet.cell(7, 12).value("вибув до А0409");

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
  const timesheetSheet = written.sheet("6. Табель");
  const timesheetHits: Array<{
    row: number;
    index: string;
    name: string;
    id: string;
    day4: string;
    day25: string;
  }> = [];
  for (let row = 7; row <= 40; row += 1) {
    const name = String(timesheetSheet.cell(row, 7).value() ?? "").trim();
    const id = String(timesheetSheet.cell(row, 8).value() ?? "").trim();
    if (!name && !id) continue;
    if (id !== person.id && !person.nameRe.test(name)) continue;
    timesheetHits.push({
      row,
      index: String(timesheetSheet.cell(row, 2).value() ?? "").trim(),
      name,
      id,
      day4: String(timesheetSheet.cell(row, 12).value() ?? "").trim(),
      day25: String(timesheetSheet.cell(row, 33).value() ?? "").trim(),
    });
  }
  const activeTimesheet = timesheetHits.filter(
    (row) => !/вибув/i.test(row.day4) && row.day25 !== "",
  );
  const slot = shpo.find((row) => row.positionIndex === person.index);

  expect(slot?.fullName).toMatch(person.nameRe);
  expect(slot?.personId).toBe(person.id);
  expect(oos.some((row) => row.personId === person.id)).toBe(true);
  expect(excluded.some((row) => row.personId === person.id)).toBe(false);
  expect(activeTimesheet, JSON.stringify(timesheetHits)).toHaveLength(1);
  expect(activeTimesheet[0]?.index).toBe(person.index);
  expect(activeTimesheet[0]?.day25).toBe("+");
};

const dobrovolskyi: CancelRestoreCase = {
  name: "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович",
  id: "12840",
  index: "2110786",
  title: "Водій зенітного ракетного відділення",
  nameRe: /добровольськ/i,
  julyFromIndex: "2103239",
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
    workbook.sheet("sh").cell(2, 1).value(null);
    workbook.sheet("sh").cell(2, 2).value(null);
    workbook.sheet("sh").cell(2, 3).value(null);
    workbook.sheet("sh").cell(2, 4).value(null);
    workbook.sheet("sh").cell(2, 5).value(null);
    workbook.sheet("sh").cell(2, 6).value("Нема в sh");
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
