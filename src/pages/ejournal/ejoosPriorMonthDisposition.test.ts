import { describe, expect, it } from "vitest";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
  type ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { buildEjoosSyncPlan } from "./ejoosSyncPlan";
import { DEFAULT_STATUS_RULES } from "./ejoosRules";
import { buildTimesheetPreview, personChangesFromOps } from "./ejoosPersonDiff";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

if (typeof window === "undefined") {
  (globalThis as { window?: typeof globalThis }).window = globalThis;
}

if (typeof FileReader === "undefined") {
  class NodeFileReader {
    result: ArrayBuffer | null = null;
    onload: ((event: { target: NodeFileReader }) => void) | null = null;
    onloadend: ((event: { target: NodeFileReader }) => void) | null = null;
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

describe("prior-month РОЗПОРЯДЖ still on staff", () => {
  const person = {
    id: "99001",
    name: "БЕВЗ Артем Сергійович",
    rank: "солдат",
    index: "2103600",
    title: "Стрілець",
    dispositionDate: "25.08.2026",
    orderNumber: "450",
  };

  const buildFixtures = async () => {
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
    sh.cell(2, 1).value(person.id);
    sh.cell(2, 2).value(person.name);
    sh.cell(2, 3).value(person.rank);
    sh.cell(2, 4).value(person.index);
    sh.cell(2, 5).value(person.title);
    sh.cell(2, 6).value("В СТРОЮ");

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
    ruh.cell(2, 1).value("450");
    ruh.cell(2, 2).value("РОЗПОРЯДЖ");
    ruh.cell(2, 3).value("РОЗПОРЯДЖ");
    ruh.cell(2, 5).value(person.id);
    ruh.cell(2, 6).value(person.rank);
    ruh.cell(2, 7).value(person.name);
    ruh.cell(2, 8).value(person.index);
    ruh.cell(2, 10).value(
      ", який знаходиться у розпорядженні командира військової частини А4862",
    );
    ruh.cell(2, 19).value(person.orderNumber);
    ruh.cell(2, 20).value(person.dispositionDate);

    const pb = await snapshotOf(
      (await pbWorkbook.outputAsync("blob")) as Blob,
      "1ПБ (Рух, archive) серпень + вересень.xlsx",
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
    timesheet.cell(2, 9).value("Вересень 2026 р.");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(7, 6).value(person.rank);
    timesheet.cell(7, 7).value(person.name);
    timesheet.cell(7, 8).value(person.id);
    for (let day = 1; day <= 8; day += 1) {
      timesheet.cell(7, 8 + day).value("+");
    }

    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_08-09-2026.xlsx",
    );
    return { pb, ejoos };
  };

  it("blocks September apply but shows August disposition plan", async () => {
    const { pb, ejoos } = await buildFixtures();
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
      sourceAsOfDate: "08.09.2026",
    });
    const personOps = plan.ops.filter(
      (op) => op.personId === person.id || /бевз/i.test(op.fullName),
    );
    const disposition = personOps.find((op) => op.kind === "move_to_disposition");
    expect(disposition).toBeDefined();
    expect(disposition?.payload.journalMonthBlocked).toBe("1");
    expect(disposition?.payload.orderDate).toBe(person.dispositionDate);
    expect(disposition?.class).toBe("needs_input");

    const people = personChangesFromOps(plan.ops, plan.timesheetDay, {
      timesheetDayLabel: plan.timesheetDayLabel,
    });
    const bevz = people.find((item) => /бевз/i.test(item.fullName));
    expect(bevz).toBeDefined();
    expect(bevz?.timesheetPreview?.runs).toEqual([
      { from: 1, to: 24, mark: "+" },
      { from: 25, to: 25, mark: "ПЕРЕВ" },
      { from: 26, to: 31, mark: "-" },
    ]);
    expect(bevz?.ejoosWillDo.join("\n")).toMatch(/станом на.*25\.08\.2026/i);
  }, 30_000);

  it("allows August apply when sync as-of is 25.08.2026", async () => {
    const { pb, ejoos } = await buildFixtures();
    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
      sourceAsOfDate: "25.08.2026",
    });
    const personOps = plan.ops.filter(
      (op) => op.personId === person.id || /бевз/i.test(op.fullName),
    );
    expect(
      personOps.some((op) => op.payload.journalMonthBlocked === "1"),
    ).toBe(false);
    const disposition = personOps.find((op) => op.kind === "move_to_disposition");
    expect(disposition).toBeDefined();
    expect(disposition?.class).toBe("ready");
    expect(disposition?.payload.orderDate).toBe(person.dispositionDate);
    const preview = buildTimesheetPreview(
      disposition ? [disposition] : [],
      plan.timesheetDay,
      plan.timesheetDayLabel,
    );
    expect(preview?.runs).toEqual([
      { from: 1, to: 24, mark: "+" },
      { from: 25, to: 25, mark: "ПЕРЕВ" },
      { from: 26, to: 31, mark: "-" },
    ]);
  }, 30_000);

  it("shows August order day, then dashes through month end on 25.08 as-of", () => {
    const disposition = {
      id: "disp-bevz",
      kind: "move_to_disposition" as const,
      class: "ready" as const,
      sheet: "Табель",
      personId: person.id,
      fullName: person.name,
      positionIndex: person.index,
      rank: person.rank,
      before: "",
      after: "",
      sourceRef: "",
      why: "",
      confidence: "high" as const,
      payload: {
        orderDate: person.dispositionDate,
        orderNumber: person.orderNumber,
        destination: "у розпорядження командира",
        timesheetFound: "true",
        absenceCode: "РОЗПОРЯДЖЕННЯ",
      },
      checkedDefault: true,
    };
    const preview = buildTimesheetPreview(
      [disposition],
      25,
      "25.08.2026",
    );
    expect(preview?.runs).toEqual([
      { from: 1, to: 24, mark: "+" },
      { from: 25, to: 25, mark: "ПЕРЕВ" },
      { from: 26, to: 31, mark: "-" },
    ]);
  });
});
