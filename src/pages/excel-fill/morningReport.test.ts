import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { analyzeMorningReport } from "./ExcelFillPage";
import { rosterRowsToSourceSnapshot } from "./rosterSourceSnapshot";

const rosterPerson = (
  rowNumber: number,
  name: string,
  birthDate: string,
): EjournalPreviewRow => ({
  __dbRowId: `row-${rowNumber}`,
  __rowNumber: rowNumber,
  column_14: name,
  column_16: birthDate,
  column_21: "В строю",
  column_22: "Забезпечення",
  column_31: "ППД Вишневе",
});

describe("analyzeMorningReport", () => {
  it("deduplicates the same person by name and birth date", () => {
    const snapshot = rosterRowsToSourceSnapshot(
      [
        rosterPerson(
          285,
          "ШЕВЧЕНКО Олександр Володимирович (11.05.1981 р.н.)",
          "11.05.1981",
        ),
        rosterPerson(
          286,
          "ШЕВЧЕНКО Олександр Володимирович (11.05.1981 р.н.)",
          "11.05.1981",
        ),
        rosterPerson(
          287,
          "ШЕВЧЕНКО Олександр Володимирович",
          "12.05.1981",
        ),
      ],
      "Штатка.xlsx",
    );

    const report = analyzeMorningReport(snapshot);

    expect(report?.rows).toHaveLength(2);
    expect(report?.skippedRows).toBe(1);
    expect(report?.rows.map((row) => row.sourceRowNumber)).toEqual([285, 287]);
  });
});
