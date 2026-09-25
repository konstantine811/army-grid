import { describe, expect, it } from "vitest";
import {
  makeUniqueZipEntryName,
  splitStaffQuestionnairePeople,
  STAFF_QUESTIONNAIRE_ZIP_PART_SIZE,
} from "./staffQuestionnairesZipExport";

describe("makeUniqueZipEntryName", () => {
  it("keeps the first file name unchanged", () => {
    const used = new Map<string, number>();
    expect(makeUniqueZipEntryName("Іваненко Іван.pdf", used)).toBe(
      "Іваненко Іван.pdf",
    );
    expect(used.get("іваненко іван.pdf")).toBe(1);
  });

  it("suffixes duplicate pdf names", () => {
    const used = new Map<string, number>();
    makeUniqueZipEntryName("Іваненко Іван.pdf", used);
    expect(makeUniqueZipEntryName("Іваненко Іван.pdf", used)).toBe(
      "Іваненко Іван (2).pdf",
    );
  });
});

describe("splitStaffQuestionnairePeople", () => {
  it("splits large lists into fixed-size parts", () => {
    const people = Array.from({ length: 160 }, (_, index) => index);
    const parts = splitStaffQuestionnairePeople(people, 75);
    expect(parts).toHaveLength(3);
    expect(parts[0]).toHaveLength(75);
    expect(parts[1]).toHaveLength(75);
    expect(parts[2]).toHaveLength(10);
  });

  it("uses default part size", () => {
    const people = Array.from(
      { length: STAFF_QUESTIONNAIRE_ZIP_PART_SIZE + 1 },
      (_, index) => index,
    );
    expect(splitStaffQuestionnairePeople(people)).toHaveLength(2);
  });
});
