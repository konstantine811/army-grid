import { describe, expect, it } from "vitest";
import type { ExcelWorkbookSnapshot } from "../../../../excelRoundTrip";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../../../../excelRoundTrip";
import { DEFAULT_STATUS_RULES } from "../../ejoosRules";
import { buildEjoosSyncPlan } from "../../ejoosSyncPlan";

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

/** Regression guard before splitting buildEjoosSyncPlan by op kind. */
describe("buildEjoosSyncPlan regression", () => {
  it("temp arrival closure: stable op kinds and placement payload", async () => {
    const person = {
      id: "21374",
      name: "МАССАЙ Вадим Євгенійович",
      index: "2103350",
      rank: "солдат",
    };
    const XlsxPopulate = await loadPopulate();

    const pbWorkbook = await XlsxPopulate.fromBlankAsync();
    const sh = pbWorkbook.sheet(0);
    sh.name("sh");
    const ruh = pbWorkbook.addSheet("РУХ");
    pbWorkbook.addSheet("archive");
    sh.cell(1, 1).value("ID");
    sh.cell(1, 2).value("ПІБ");
    sh.cell(1, 3).value("Звання");
    sh.cell(1, 4).value("Індекс посади");
    sh.cell(1, 6).value("Статус");
    sh.cell(2, 1).value(person.id);
    sh.cell(2, 2).value(person.name);
    sh.cell(2, 3).value(person.rank);
    sh.cell(2, 4).value(person.index);
    sh.cell(2, 6).value("21_МР");
    [
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
    ].forEach((header, index) => {
      ruh.cell(1, index + 1).value(header);
    });
    ruh.cell(2, 1).value("1");
    ruh.cell(2, 2).value("ПРИБУВ");
    ruh.cell(2, 5).value(person.id);
    ruh.cell(2, 6).value(person.rank);
    ruh.cell(2, 7).value(person.name);
    ruh.cell(2, 8).value("БРЕЗ");
    ruh.cell(2, 9).value(person.index);
    ruh.cell(2, 19).value("249");
    ruh.cell(2, 20).value("26.08.2026");
    const pb = await snapshotOf(
      (await pbWorkbook.outputAsync("blob")) as Blob,
      "1ПБ_07092026.xlsx",
    );

    const ejoosWorkbook = await XlsxPopulate.fromBlankAsync();
    const shpo = ejoosWorkbook.sheet(0);
    shpo.name("1. ШПО");
    ejoosWorkbook.addSheet("2. ООС");
    const arrivals = ejoosWorkbook.addSheet("4. Тимчасово прибулі");
    ejoosWorkbook.addSheet("5. Тимчасово відсутні");
    const timesheet = ejoosWorkbook.addSheet("6. Табель");
    shpo.cell(1, 1).value("1. ШПО");
    shpo.cell(7, 1).value(person.index);
    shpo.cell(7, 6).value(person.rank);
    shpo.cell(7, 7).value(person.name);
    shpo.cell(7, 8).value(person.id);
    arrivals.cell(1, 1).value("4. Тимчасово прибулі");
    arrivals.cell(6, 1).value(person.rank);
    arrivals.cell(6, 2).value(person.name);
    arrivals.cell(6, 6).value("БРЕЗ");
    arrivals.cell(6, 8).value("13.05.2026");
    timesheet.cell(1, 1).value("6. Табель");
    timesheet.cell(2, 9).value("Серпень 2026 р.");
    timesheet.cell(7, 2).value(person.index);
    timesheet.cell(7, 7).value(person.name);
    timesheet.cell(7, 8).value(person.id);
    timesheet.cell(7, 6).value(person.rank);
    const ejoos = await snapshotOf(
      (await ejoosWorkbook.outputAsync("blob")) as Blob,
      "ЄЖООС_станом_на_31-08-2026.xlsx",
    );

    const plan = buildEjoosSyncPlan(ejoos, pb, {
      statusRules: DEFAULT_STATUS_RULES,
    });

    const personOps = plan.ops.filter((op) => op.personId === person.id);
    expect(personOps.map((op) => op.kind).sort()).toMatchSnapshot();
    expect(
      personOps.find(
        (op) =>
          op.kind === "position_change" &&
          op.payload.isTempArrivalPlacement === "1",
      ),
    ).toMatchObject({
      class: "ready",
      payload: {
        arrivalExcelRow: "6",
        nextIndex: person.index,
      },
    });
    expect(
      personOps.some((op) => op.kind === "arrival"),
    ).toBe(false);
  });
});
