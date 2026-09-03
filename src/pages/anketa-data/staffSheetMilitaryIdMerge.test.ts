import { describe, expect, it } from "vitest";
import {
  planStaffSheetMilitaryIdMerges,
  resolveStaffSheetMilitaryIdValue,
} from "./staffSheetMilitaryIdMerge";
import type { VkTpvDovidkyNameEntry } from "../personnel/vkTpvDovidkyImport";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";

const vkEntry = (militaryId: string): VkTpvDovidkyNameEntry => ({
  fullName: "Іванов Іван Іванович",
  nameKey: "іванов іван іванович",
  militaryId,
  rnokpp: "",
  sourceSheet: "ВК",
});

const rosterRow = (
  fullName: string,
  militaryId = "",
  excelRowNumber = 2,
): EjournalPreviewRow =>
  ({
    __rowNumber: excelRowNumber,
    column_14: fullName,
    column_11: militaryId,
  }) as EjournalPreviewRow;

describe("resolveStaffSheetMilitaryIdValue", () => {
  it("keeps existing staff value", () => {
    expect(
      resolveStaffSheetMilitaryIdValue("АГ 111111", vkEntry("АГ 222222")).action,
    ).toBe("kept_staff");
  });

  it("replaces placeholder staff value with VK number", () => {
    const result = resolveStaffSheetMilitaryIdValue(
      "резерв+",
      vkEntry("МО 312448"),
    );
    expect(result.action).toBe("from_vk");
    expect(result.value).toBe("МО 312448");
  });

  it("replaces note-like staff value with VK number", () => {
    const result = resolveStaffSheetMilitaryIdValue(
      "довідка",
      vkEntry("АГ 333333"),
    );
    expect(result.action).toBe("from_vk");
    expect(result.value).toBe("АГ 333333");
  });

  it("normalizes embedded staff number from messy cell text", () => {
    const result = resolveStaffSheetMilitaryIdValue("резерв+ МО 312448", undefined);
    expect(result.action).toBe("kept_staff");
    expect(result.value).toBe("МО 312448");
  });

  it("fills VK when staff cell is empty", () => {
    const result = resolveStaffSheetMilitaryIdValue("", vkEntry("АГ 333333"));
    expect(result.action).toBe("from_vk");
    expect(result.value).toBe("АГ 333333");
  });

  it("leaves empty when VK has no id", () => {
    const result = resolveStaffSheetMilitaryIdValue("", vkEntry(""));
    expect(result.action).toBe("unchanged");
    expect(result.value).toBe("");
  });

  it("leaves empty when person is missing in VK file", () => {
    const result = resolveStaffSheetMilitaryIdValue("", undefined);
    expect(result.action).toBe("unchanged");
    expect(result.value).toBe("");
  });

  it("ignores column header text in staff cell", () => {
    const result = resolveStaffSheetMilitaryIdValue(
      "Військовий квиток",
      vkEntry("МО 312448"),
    );
    expect(result.action).toBe("from_vk");
    expect(result.value).toBe("МО 312448");
  });
});

describe("planStaffSheetMilitaryIdMerges", () => {
  it("counts merge actions for roster rows", () => {
    const vkIndex = new Map<string, VkTpvDovidkyNameEntry>([
      ["іванов іван іванович", vkEntry("АГ 111111")],
      ["петров петро петрович", vkEntry("")],
    ]);
    const rows = [
      rosterRow("Іванов Іван Іванович"),
      rosterRow("Петров Петро Петрович", "АГ 999999"),
      rosterRow("Сидоренко Сидір Сидорович"),
    ];

    const report = planStaffSheetMilitaryIdMerges(rows, vkIndex);
    expect(report.fromVk).toBe(1);
    expect(report.keptStaff).toBe(1);
    expect(report.unchanged).toBe(1);
  });
});
