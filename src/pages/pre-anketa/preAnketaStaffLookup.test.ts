import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { createEmptyPreAnketaForm } from "./preAnketaFields";
import {
  applyPreAnketaStaffLookupHit,
  lookupPreAnketaStaffByFullName,
} from "./preAnketaStaffLookup";

const rosterRow = (
  name: string,
  extra: Record<string, unknown> = {},
): EjournalPreviewRow =>
  ({
    __dbRowId: name,
    column_14: name,
    column_13: "солдат",
    column_15: "СІМБА",
    ...extra,
  }) as EjournalPreviewRow;

describe("preAnketaStaffLookup", () => {
  it("finds callsign and rank by full name", () => {
    const hit = lookupPreAnketaStaffByFullName("СІБАРЦЕВ ДМИТРО МИХАЙЛОВИЧ", [
      rosterRow("СІБАРЦЕВ ДМИТРО МИХАЙЛОВИЧ"),
    ]);
    expect(hit).toMatchObject({
      callsign: "СІМБА",
      rank: "солдат",
    });
  });

  it("matches surname and given name without patronymic", () => {
    const hit = lookupPreAnketaStaffByFullName("СІБАРЦЕВ Дмитро", [
      rosterRow("СІБАРЦЕВ ДМИТРО МИХАЙЛОВИЧ"),
    ]);
    expect(hit?.callsign).toBe("СІМБА");
  });

  it("returns null for ambiguous matches", () => {
    const hit = lookupPreAnketaStaffByFullName("ІВАНОВ Іван Іванович", [
      rosterRow("ІВАНОВ Іван Іванович", { __dbRowId: "1" }),
      rosterRow("ІВАНОВ Іван Іванович", { __dbRowId: "2", column_15: "БУР" }),
    ]);
    expect(hit).toBeNull();
  });

  it("fills only empty callsign and rank fields", () => {
    const base = {
      ...createEmptyPreAnketaForm(),
      fullName: "СІБАРЦЕВ ДМИТРО МИХАЙЛОВИЧ",
      rank: "капітан",
    };
    const applied = applyPreAnketaStaffLookupHit(base, {
      callsign: "СІМБА",
      rank: "солдат",
      rosterName: "СІБАРЦЕВ ДМИТРО МИХАЙЛОВИЧ",
    });
    expect(applied.fields.callsign).toBe("СІМБА");
    expect(applied.fields.rank).toBe("капітан");
    expect(applied.changed).toBe(true);
  });
});
