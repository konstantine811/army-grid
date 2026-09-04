import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  createLostMilitaryIdFields,
  declinedInvestigator,
  investigatorFromPersonnelRow,
  mergeLostMilitaryIdFields,
  militaryUnitLabel,
  normalizeMilitaryUnitPhrase,
} from "./lostMilitaryIdReport";
import { buildPersonSummary } from "../personnel/personnelUtils";

const person = (name: string, extra: Record<string, unknown> = {}) =>
  ({
    __dbRowId: extra.__dbRowId ?? `row:${name}`,
    прізвище: name,
    звання: extra.rank ?? "солдат",
    повна_посада: extra.position ?? "командир відділення",
    id: extra.id ?? "11524",
    ...extra,
  }) as EjournalPreviewRow;

describe("normalizeMilitaryUnitPhrase", () => {
  it("adds prefix when the field holds only the unit number", () => {
    expect(normalizeMilitaryUnitPhrase("А4862")).toBe("військової частини А4862");
  });

  it("does not duplicate prefix when it is already present", () => {
    expect(normalizeMilitaryUnitPhrase("військової частини А4862")).toBe(
      "військової частини А4862",
    );
  });

  it("extracts the unit number for short references", () => {
    expect(militaryUnitLabel("військової частини А4862")).toBe("А4862");
    expect(militaryUnitLabel("А4862")).toBe("А4862");
  });
});

describe("investigatorFromPersonnelRow", () => {
  it("fills name, rank and dative position from Особовий склад", () => {
    const row = person("ГУНЬКО Олександр Олександрович", {
      rank: "солдат",
      position: "командир відділення — командир екіпажу безпілотних літальних комплексів",
      id: "2103111",
    });

    expect(investigatorFromPersonnelRow(row)).toMatchObject({
      investigatorFullName: "ГУНЬКО Олександр Олександрович",
      investigatorRank: "солдат",
      investigatorPosition:
        "Командиру відділення — командир екіпажу безпілотних літальних комплексів",
      investigatorPersonId: "2103111",
      investigatorManual: false,
    });
  });
});

describe("mergeLostMilitaryIdFields investigator mode", () => {
  it("opens old filled reports in manual mode so the typed name stays editable", () => {
    const defaults = createLostMilitaryIdFields(null, buildPersonSummary(null));
    const merged = mergeLostMilitaryIdFields(defaults, {
      investigatorFullName: "ГУНЬКО Олександр Олександрович",
      investigatorRank: "солдат",
    });
    expect(merged.investigatorManual).toBe(true);
  });

  it("keeps picker mode when the person was chosen from Особовий склад", () => {
    const defaults = createLostMilitaryIdFields(null, buildPersonSummary(null));
    const merged = mergeLostMilitaryIdFields(defaults, {
      investigatorFullName: "ГУНЬКО Олександр Олександрович",
      investigatorPersonId: "2103111",
      investigatorManual: false,
    });
    expect(merged.investigatorManual).toBe(false);
    expect(merged.investigatorPersonId).toBe("2103111");
  });

  it("capitalizes investigator position in the form and report text", () => {
    const defaults = createLostMilitaryIdFields(null, buildPersonSummary(null));
    const merged = mergeLostMilitaryIdFields(defaults, {
      investigatorPosition:
        "командиру відділення — командиру екіпажу безпілотних літальних комплексів",
    });
    expect(merged.investigatorPosition.startsWith("К")).toBe(true);
    expect(declinedInvestigator(merged).position.startsWith("К")).toBe(true);
  });
});
