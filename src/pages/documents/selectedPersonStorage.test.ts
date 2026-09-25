import { beforeEach, describe, expect, it } from "vitest";
import {
  compactPersonRowForStorage,
  storeSelectedPersonForDocuments,
} from "./selectedPersonStorage";

const memoryStorage = () => {
  const data = new Map<string, string>();
  return {
    getItem: (key: string) => data.get(key) ?? null,
    setItem: (key: string, value: string) => {
      data.set(key, value);
    },
    removeItem: (key: string) => {
      data.delete(key);
    },
    clear: () => data.clear(),
  };
};

describe("selectedPersonStorage", () => {
  beforeEach(() => {
    Object.defineProperty(globalThis, "localStorage", {
      configurable: true,
      value: memoryStorage(),
    });
    Object.defineProperty(globalThis, "sessionStorage", {
      configurable: true,
      value: memoryStorage(),
    });
    Object.defineProperty(globalThis, "window", {
      configurable: true,
      value: globalThis,
    });
  });

  it("drops oversized cell text before writing", () => {
    const compact = compactPersonRowForStorage({
      __dbRowId: "row-1",
      піб: "Іванов Іван",
      blob: "x".repeat(9_000),
    });
    expect(compact.__dbRowId).toBe("row-1");
    expect(compact["піб"]).toBe("Іванов Іван");
    expect(compact.blob).toBeUndefined();
  });

  it("does not throw when localStorage is full", () => {
    const setItem = localStorage.setItem.bind(localStorage);
    localStorage.setItem = () => {
      throw new DOMException("quota", "QuotaExceededError");
    };
    expect(() =>
      storeSelectedPersonForDocuments(
        { __dbRowId: "row-1", піб: "Іванов" },
        "ubdReport",
      ),
    ).not.toThrow();
    localStorage.setItem = setItem;
    expect(sessionStorage.getItem("army-grid:selected-person")).toContain(
      "Іванов",
    );
  });
});
