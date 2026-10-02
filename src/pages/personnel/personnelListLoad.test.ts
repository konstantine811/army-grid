import { describe, expect, it } from "vitest";
import {
  emptyStaffSheetMessage,
  nextPersonnelRowId,
  personnelDatasetLoadError,
  personnelPreviewError,
} from "./personnelListLoad";

describe("personnel list load", () => {
  it("keeps the current person when the focused row has no id", () => {
    const rows = [{ __dbRowId: "12" }, { __dbRowId: "4246" }];
    expect(nextPersonnelRowId("4246", undefined, rows)).toBe("4246");
    expect(nextPersonnelRowId("missing", "12", rows)).toBe("12");
    expect(nextPersonnelRowId("", undefined, rows)).toBe("12");
  });

  it("uses the same status text as the page", () => {
    expect(personnelDatasetLoadError("ні")).toBe(
      "Не вдалося завантажити особовий склад.",
    );
    expect(personnelPreviewError(new Error("кеш"))).toBe("кеш");
    expect(emptyStaffSheetMessage).toBe("У штатці немає рядків.");
  });
});
