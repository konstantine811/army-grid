import { describe, expect, it } from "vitest";
import {
  buildVkTpvDovidkyNameIndex,
  extractMilitaryIdFromText,
  normalizeMilitaryIdCellValue,
} from "./vkTpvDovidkyImport";
import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { normalizeAnketaNameKey } from "../anketa-data/anketaPersonMatch";

describe("extractMilitaryIdFromText", () => {
  it("extracts classic AG/UN/AV ids", () => {
    expect(extractMilitaryIdFromText("АГ 565010")).toBe("АГ 565010");
    expect(extractMilitaryIdFromText("АВ 951365")).toBe("АВ 951365");
  });

  it("extracts MO/GG/NK ids from VK table", () => {
    expect(extractMilitaryIdFromText("МО 312448")).toBe("МО 312448");
    expect(extractMilitaryIdFromText("ГГ 198149")).toBe("ГГ 198149");
    expect(extractMilitaryIdFromText("НК 6332792")).toBe("НК 6332792");
    expect(extractMilitaryIdFromText("AB 909887")).toBe("AB 909887");
  });

  it("extracts TPV ids", () => {
    expect(extractMilitaryIdFromText("ТПВ №4160")).toBe("ТПВ №4160");
    expect(extractMilitaryIdFromText("ТВП №1204")).toBe("ТВП №1204");
  });

  it("returns empty for placeholders", () => {
    expect(extractMilitaryIdFromText("резерв+")).toBe("");
    expect(extractMilitaryIdFromText("довідка")).toBe("");
    expect(extractMilitaryIdFromText("посвідчення офіцера")).toBe("");
  });
});

describe("normalizeMilitaryIdCellValue", () => {
  it("returns extracted number when present", () => {
    expect(normalizeMilitaryIdCellValue("резерв+")).toBe("резерв+");
    expect(normalizeMilitaryIdCellValue("МО 312448")).toBe("МО 312448");
    expect(normalizeMilitaryIdCellValue("довідка МО 312448")).toBe("МО 312448");
  });
});

describe("buildVkTpvDovidkyNameIndex", () => {
  it("remembers the absent-sheet status when the same person also has a number elsewhere", () => {
    const fullName = "ДМИТРІЄВ Андрій Сергійович";
    const absentRow = Array.from({ length: 15 }, () => null) as unknown[];
    absentRow[4] = fullName;
    const snapshot = {
      sheets: [
        {
          sheetName: "ВК",
          rawRows: [["ПІБ", "ВК"], [fullName, "АГ 123456"]],
          headerRows: [],
          rows: [],
        },
        {
          sheetName: "ВІДСУТНІ",
          rawRows: [absentRow],
          headerRows: [],
          rows: [],
        },
      ],
    } as unknown as ExcelWorkbookSnapshot;

    const entry = buildVkTpvDovidkyNameIndex(snapshot).get(
      normalizeAnketaNameKey(fullName),
    );
    expect(entry?.militaryId).toBe("АГ 123456");
    expect(entry?.isAbsent).toBe(true);
  });

  it("uses the TPV number when a person is present on both VK and TPV sheets", () => {
    const fullName = "ЮРКО Артем Володимирович";
    const tpvHeader = Array.from({ length: 32 }, () => null) as unknown[];
    tpvHeader[0] = "№";
    tpvHeader[6] = "ПІБ";
    tpvHeader[13] = "№";
    tpvHeader[29] = "ВК№";
    const tpvRow = Array.from({ length: 32 }, () => null) as unknown[];
    tpvRow[0] = 97;
    tpvRow[6] = fullName;
    tpvRow[13] = 5759;
    const snapshot = {
      sheets: [
        {
          sheetName: "ВК",
          rawRows: [["ПІБ", "ВК"], [fullName, "АВ 906881"]],
          headerRows: [],
          rows: [],
        },
        {
          sheetName: "ТПВ",
          rawRows: [tpvHeader, tpvRow],
          headerRows: [],
          rows: [],
        },
      ],
    } as unknown as ExcelWorkbookSnapshot;

    const entry = buildVkTpvDovidkyNameIndex(snapshot).get(
      normalizeAnketaNameKey(fullName),
    );
    expect(entry?.militaryId).toBe("ТПВ №5759");
    expect(entry?.isTpv).toBe(true);
  });

  it("uses certificate status when a person is present on both VK and DOVIDKY sheets", () => {
    const fullName = "АБРАМЕНКО Костянтин Олександрович";
    const snapshot = {
      sheets: [
        {
          sheetName: "ВК",
          rawRows: [["ПІБ", "ВК"], [fullName, "АГ 123456"]],
          headerRows: [],
          rows: [],
        },
        {
          sheetName: "ДОВІДКИ",
          rawRows: [["ПІБ", "ВК№"], [fullName, null]],
          headerRows: [],
          rows: [],
        },
      ],
    } as unknown as ExcelWorkbookSnapshot;

    const entry = buildVkTpvDovidkyNameIndex(snapshot).get(
      normalizeAnketaNameKey(fullName),
    );
    expect(entry?.militaryId).toBe("довідка");
    expect(entry?.isCertificate).toBe(true);
  });
});
