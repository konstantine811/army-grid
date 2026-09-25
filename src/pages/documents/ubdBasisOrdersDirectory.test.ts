import { describe, expect, it } from "vitest";
import {
  IMPORTED_BASIS_ORDERS_STORAGE_KEY,
  compactImportedBasisOrders,
  writeImportedBasisOrdersToLocalStorage,
} from "./ubdBasisOrdersDirectory";

describe("imported basis orders storage", () => {
  it("stores only number and date", () => {
    expect(
      compactImportedBasisOrders([
        {
          number: "4862/ОКП/1284/дск",
          date: "01.11.2025",
          note: "Полкові БР листопад 2025.docx",
        },
      ]),
    ).toEqual([{ number: "4862/ОКП/1284/дск", date: "01.11.2025" }]);
  });

  it("rewrites the local list when the previous value fills the quota", () => {
    let stored = "x".repeat(20);
    const storage = {
      setItem(_key: string, value: string) {
        if (stored.length > 10) throw new Error("quota");
        stored = value;
      },
      removeItem(key: string) {
        expect(key).toBe(IMPORTED_BASIS_ORDERS_STORAGE_KEY);
        stored = "";
      },
    };

    expect(
      writeImportedBasisOrdersToLocalStorage(storage, [
        { number: "4862/ОКП/1284/дск", date: "01.11.2025" },
      ]),
    ).toBe("rewritten");
    expect(stored).toContain("4862/ОКП/1284/дск");
  });
});
