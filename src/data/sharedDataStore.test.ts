import { afterEach, describe, expect, it, vi } from "vitest";
import {
  clearSharedDataStore,
  deleteSharedDataEntry,
  getSharedDataEntry,
  setSharedDataEntry,
  sharedDataStore,
} from "./sharedDataStore";

afterEach(clearSharedDataStore);

describe("sharedDataStore", () => {
  it("shares the same payload reference and notifies subscribers", () => {
    const listener = vi.fn();
    const unsubscribe = sharedDataStore.subscribe(listener);
    const rows = [{ id: "1" }];

    setSharedDataEntry({ key: "personnel", value: rows, savedAt: 1 });

    expect(getSharedDataEntry("personnel")?.value).toBe(rows);
    expect(listener).toHaveBeenCalledTimes(1);

    deleteSharedDataEntry("personnel");
    expect(getSharedDataEntry("personnel")).toBeUndefined();
    expect(listener).toHaveBeenCalledTimes(2);
    unsubscribe();
  });
});
