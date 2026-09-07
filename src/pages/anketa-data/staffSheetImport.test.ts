import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";

// Keep conversion helper behavior aligned with parse → DB import.
const previewRowToImportValues = (row: EjournalPreviewRow) => {
  const values: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(row)) {
    if (key.startsWith("__")) continue;
    if (value == null || String(value).trim() === "") continue;
    values[key] = value;
  }
  return values;
};

describe("staffSheetImport row parity", () => {
  it("keeps named preview rows when converting for DB import", () => {
    const previewRows: EjournalPreviewRow[] = [
      {
        __rowNumber: 5,
        column_14: "ПЕТРЕНКО Петро Петрович",
        column_2: "1 рота",
      } as EjournalPreviewRow,
      {
        __rowNumber: 6,
        column_2: "2 рота",
      } as EjournalPreviewRow,
    ];

    const dbRows = previewRows
      .map((row) => ({
        excelRowNumber: Number(row.__rowNumber) || 0,
        values: previewRowToImportValues(row),
      }))
      .filter((row) => Object.keys(row.values).length > 0);

    expect(dbRows).toHaveLength(2);
    expect(dbRows[0]?.values.column_14).toBe("ПЕТРЕНКО Петро Петрович");
  });
});
