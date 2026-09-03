import { describe, expect, it } from "vitest";
import { createEmptyAnketaRow } from "./anketaSheet";
import {
  ANKETA_MILITARY_ID_ABSENT_VALUE,
  isAnketaRnokppReplaceable,
  planAnketaMilitaryIdMerges,
  resolveAnketaMilitaryIdValue,
  resolveAnketaRnokppValue,
} from "./anketaMilitaryIdImport";
import type { VkTpvDovidkyNameEntry } from "../personnel/vkTpvDovidkyImport";

const vkEntry = (
  militaryId: string,
  rnokpp = "",
  fullName = "Іванов Іван Іванович",
): VkTpvDovidkyNameEntry => ({
  fullName,
  nameKey: "іванов іван іванович",
  militaryId,
  rnokpp,
  sourceSheet: "ВК",
});

describe("resolveAnketaMilitaryIdValue", () => {
  it("uses VK number when both anketa and VK have values", () => {
    expect(
      resolveAnketaMilitaryIdValue("АГ 111111", vkEntry("АГ 222222")).action,
    ).toBe("from_vk");
    expect(
      resolveAnketaMilitaryIdValue("АГ 111111", vkEntry("АГ 222222")).value,
    ).toBe("АГ 222222");
  });

  it("keeps anketa value when VK has no military id", () => {
    const result = resolveAnketaMilitaryIdValue("АГ 111111", vkEntry(""));
    expect(result.action).toBe("kept_anketa");
    expect(result.value).toBe("АГ 111111");
  });

  it("marks absent when cell is empty and person is in VK file without id", () => {
    const result = resolveAnketaMilitaryIdValue("", vkEntry(""));
    expect(result.action).toBe("marked_absent");
    expect(result.value).toBe(ANKETA_MILITARY_ID_ABSENT_VALUE);
  });

  it("marks absent when cell is empty and person is not in VK file", () => {
    const result = resolveAnketaMilitaryIdValue("", undefined);
    expect(result.action).toBe("marked_not_found");
    expect(result.value).toBe(ANKETA_MILITARY_ID_ABSENT_VALUE);
  });

  it("fills VK number into empty anketa cell", () => {
    const result = resolveAnketaMilitaryIdValue("", vkEntry("АГ 333333"));
    expect(result.action).toBe("from_vk");
    expect(result.value).toBe("АГ 333333");
  });
});

describe("resolveAnketaRnokppValue", () => {
  it("keeps valid 10-digit INN from anketa", () => {
    const result = resolveAnketaRnokppValue(
      "2820405493",
      vkEntry("", "1234567890"),
    );
    expect(result.action).toBe("kept_anketa");
    expect(result.value).toBe("2820405493");
  });

  it("fills INN from VK when anketa cell is empty", () => {
    const result = resolveAnketaRnokppValue("", vkEntry("", "1234567890"));
    expect(result.action).toBe("from_vk");
    expect(result.value).toBe("1234567890");
  });

  it("fills INN from VK when anketa has placeholder text", () => {
    expect(isAnketaRnokppReplaceable("забув")).toBe(true);
    expect(isAnketaRnokppReplaceable("не має")).toBe(true);
    const result = resolveAnketaRnokppValue(
      "забув",
      vkEntry("", "9876543210"),
    );
    expect(result.action).toBe("from_vk");
    expect(result.value).toBe("9876543210");
  });

  it("leaves placeholder when VK has no INN", () => {
    const result = resolveAnketaRnokppValue("немає", vkEntry("", ""));
    expect(result.action).toBe("unchanged");
    expect(result.value).toBe("немає");
  });
});

describe("planAnketaMilitaryIdMerges", () => {
  it("plans edits for military id and rnokpp", () => {
    const rows = [
      {
        ...createEmptyAnketaRow(2),
        fullName: "Іванов Іван Іванович",
        militaryId: "",
        rnokpp: "забув",
      },
      {
        ...createEmptyAnketaRow(3),
        fullName: "Петров Петро Петрович",
        militaryId: "АГ 999999",
        rnokpp: "2820405493",
      },
      {
        ...createEmptyAnketaRow(4),
        fullName: "Сидоренко Сидір Сидорович",
        militaryId: "",
        rnokpp: "",
      },
    ];
    const vkIndex = new Map<string, VkTpvDovidkyNameEntry>([
      ["іванов іван іванович", vkEntry("АГ 111111", "1111111111")],
      ["петров петро петрович", vkEntry("АГ 222222", "2222222222")],
    ]);

    const { edits, report } = planAnketaMilitaryIdMerges(rows, vkIndex);

    expect(report.updatedFromVk).toBe(2);
    expect(report.markedNotFound).toBe(1);
    expect(report.rnokppUpdatedFromVk).toBe(1);
    expect(report.rnokppKeptAnketa).toBe(1);
    expect(edits).toHaveLength(4);
    expect(edits.find((edit) => edit.rowNumber === 2 && edit.columnId === "rnokpp")?.value).toBe(
      "1111111111",
    );
    expect(edits.find((edit) => edit.rowNumber === 3 && edit.columnId === "militaryId")?.value).toBe(
      "АГ 222222",
    );
    expect(
      edits.find((edit) => edit.rowNumber === 3 && edit.columnId === "rnokpp"),
    ).toBeUndefined();
  });
});
