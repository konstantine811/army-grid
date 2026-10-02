import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  filterPersonnelScopeRows,
  staffFilterScopeChange,
} from "./personnelListScope";

const staff = {
  __dbRowId: "staff",
  roster__column_14: "КІЯНЕНКО",
} as EjournalPreviewRow;
const outside = { __dbRowId: "oos-only", ПІБ: "Поза штатом" } as EjournalPreviewRow;

describe("staffFilterScopeChange", () => {
  it("reloads the full list only when Усі has no rows outside the staff", () => {
    expect(staffFilterScopeChange("all", [staff])).toMatchObject({
      reload: true,
      publish: false,
    });
    expect(staffFilterScopeChange("all", [staff, outside])).toMatchObject({
      reload: false,
      publish: true,
    });
  });

  it("opens the archive without reloading", () => {
    expect(staffFilterScopeChange("archive", [staff])).toEqual({
      all: false,
      archive: true,
      publish: true,
      reload: false,
    });
  });
});

describe("filterPersonnelScopeRows", () => {
  it("keeps only staff rows until the full scope is open", () => {
    expect(
      filterPersonnelScopeRows([staff, outside], { all: false, archive: false }).map(
        (row) => row.__dbRowId,
      ),
    ).toEqual(["staff"]);
    expect(
      filterPersonnelScopeRows([staff, outside], { all: true, archive: true }),
    ).toHaveLength(2);
  });
});
