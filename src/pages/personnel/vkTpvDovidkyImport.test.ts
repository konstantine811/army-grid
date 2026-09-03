import { describe, expect, it } from "vitest";
import { extractMilitaryIdFromText, normalizeMilitaryIdCellValue } from "./vkTpvDovidkyImport";

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
