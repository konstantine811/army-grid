import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  mergeBasisOrderLists,
  normalizeImportedBasisNumber,
  parseBasisOrderLine,
  parseBasisOrdersFromDocx,
  parseBasisOrdersFromText,
} from "./ubdBasisOrdersImport";

describe("ubdBasisOrdersImport", () => {
  it("parses a standard BR line", () => {
    expect(parseBasisOrderLine("№4862/ОКП/2223/дск від 01.08.2026")).toEqual({
      number: "4862/ОКП/2223/дск",
      date: "01.08.2026",
    });
  });

  it("normalizes missing slash before дск", () => {
    expect(parseBasisOrderLine("№4862/ОКП/2357дск від 12.08.2026")).toEqual({
      number: "4862/ОКП/2357/дск",
      date: "12.08.2026",
    });
    expect(normalizeImportedBasisNumber("4862/ОКП/2357дск")).toBe(
      "4862/ОКП/2357/дск",
    );
  });

  it("merges without duplicates", () => {
    const existing = [{ number: "4862/ОКП/2223/дск", date: "01.08.2026" }];
    const incoming = [
      { number: "4862/ОКП/2223/дск", date: "01.08.2026" },
      { number: "4862/ОКП/2236/дск", date: "02.08.2026" },
    ];
    const result = mergeBasisOrderLists(existing, incoming);
    expect(result.added).toBe(1);
    expect(result.skipped).toBe(1);
    expect(result.merged).toHaveLength(2);
  });

  it("parses sample 1ПБ docx", async () => {
    const samplePath =
      "/Volumes/KINGSTON/army_work/УБД/БР/1ПБ номера БР 08.2026.docx";
    let buffer: ArrayBuffer;
    try {
      buffer = readFileSync(samplePath).buffer.slice(0) as ArrayBuffer;
    } catch {
      return;
    }
    const rows = await parseBasisOrdersFromDocx(buffer);
    expect(rows.length).toBeGreaterThanOrEqual(16);
    expect(rows[0]).toEqual({
      number: "4862/ОКП/2223/дск",
      date: "01.08.2026",
    });
    expect(rows.at(-1)).toEqual({
      number: "4862/ОКП/2561/дск",
      date: "28.08.2026",
    });
  });

  it("extracts multiple entries from one text blob", () => {
    const rows = parseBasisOrdersFromText(
      "№4862/ОКП/2223/дск від 01.08.2026№4862/ОКП/2236/дск від 02.08.2026",
    );
    expect(rows).toHaveLength(2);
  });
});
