import { describe, expect, it } from "vitest";
import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  EJOOS_OOS_MILITARY_ID_COLUMN,
  EJOOS_OOS_RNOKPP_REFUSE_COLUMN,
} from "./ejoosStaffIndexFormat";
import {
  buildOosAnketaFieldColumns,
  resolveOosRnokppRefuseCleanup,
  shouldReplaceOosCellFromAnketa,
} from "./ejoosAnketaFill";

const oosSheet = (): ExcelSheetSnapshot => ({
  sheetName: "2. ООС",
  rawRows: [],
  columnIndexes: Array.from({ length: 40 }, (_, index) => index),
});

const buildOosHeaderRow = () => {
  const headerRow = Array.from({ length: 40 }, () => "");
  headerRow[21] = "РНОКПП (за наявності)";
  headerRow[22] = "Відмова від РНОКПП";
  headerRow[23] = "Назва документа, що посвідчує особу";
  headerRow[24] =
    "Серія (за наявності) та номер документа, що посвідчує особу";
  headerRow[25] = "Серія (за наявності) і номер військового квитка";
  return headerRow;
};

describe("buildOosAnketaFieldColumns", () => {
  it("maps RNOKPP and refuse separately, military ticket to column Z", () => {
    const headerRow = buildOosHeaderRow();
    const columns = buildOosAnketaFieldColumns(headerRow, oosSheet());

    expect(columns.get("rnokpp")).toBe(21);
    expect(columns.get("rnokppRefuse")).toBe(22);
    expect(columns.get("idDocumentNumber")).toBe(24);
    expect(columns.get("militaryId")).toBe(25);
    expect(columns.get("militaryId")).toBe(
      EJOOS_OOS_MILITARY_ID_COLUMN - 1,
    );
    expect(columns.get("rnokppRefuse")).toBe(
      EJOOS_OOS_RNOKPP_REFUSE_COLUMN - 1,
    );
  });

  it("does not map military ticket to refuse column W", () => {
    const headerRow = buildOosHeaderRow();
    const columns = buildOosAnketaFieldColumns(headerRow, oosSheet());

    expect(columns.get("militaryId")).not.toBe(columns.get("rnokppRefuse"));
  });

  it("keeps militaryId when only gap columns are requested", () => {
    const columns = buildOosAnketaFieldColumns(buildOosHeaderRow(), oosSheet(), {
      gapKeys: new Set(["militaryId"]),
    });

    expect([...columns.keys()]).toEqual(["militaryId"]);
  });
});

describe("shouldReplaceOosCellFromAnketa", () => {
  it("replaces AG in OOS when anketa has довідка", () => {
    expect(
      shouldReplaceOosCellFromAnketa(
        "fill",
        "militaryId",
        "АГ 484613",
        "довідка",
        new Set(),
      ),
    ).toBe(true);
  });

  it("replaces in merge mode for selected gap columns", () => {
    expect(
      shouldReplaceOosCellFromAnketa(
        "merge",
        "militaryId",
        "АГ 484613",
        "довідка",
        new Set(["militaryId"]),
      ),
    ).toBe(true);
  });
});

describe("resolveOosRnokppRefuseCleanup", () => {
  it.each([
    "АГ 484613",
    "АГ484613",
    "довідка",
    "відсутній",
    "ТПВ №4192",
    "будь-які інші помилкові дані",
  ])(
    "clears all existing values (%s) from RNOKPP refusal",
    (value) => {
      expect(resolveOosRnokppRefuseCleanup(value)).toBeNull();
    },
  );

  it("does not write an already empty cell again", () => {
    expect(resolveOosRnokppRefuseCleanup("")).toBeUndefined();
    expect(resolveOosRnokppRefuseCleanup(null)).toBeUndefined();
  });
});
