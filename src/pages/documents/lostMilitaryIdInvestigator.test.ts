import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  actApprovalDateLine,
  createLostMilitaryIdFields,
  declinedInvestigator,
  instrumentalInvestigatorLine,
  investigatorFooterBlock,
  investigatorFooterLines,
  investigatorFromPersonnelRow,
  mergeLostMilitaryIdFields,
  militaryUnitLabel,
  normalizeMilitaryUnitPhrase,
  splitLostMilitaryIdSignatory,
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

describe("act document blocks", () => {
  it("uses instrumental case for investigator in act intro", () => {
    const fields = mergeLostMilitaryIdFields(
      createLostMilitaryIdFields(null, buildPersonSummary(null)),
      {
        investigatorFullName: "ПАЛЮХ Олег",
        investigatorRank: "старший лейтенант",
        investigatorPosition:
          "заступник командира роти вогневої підтримки з психологічної підтримки персоналу",
      },
    );
    const line = instrumentalInvestigatorLine(fields);
    expect(line).toMatch(/заступником/i);
    expect(line).toMatch(/старшим лейтенантом/i);
    expect(line).toMatch(/Олегом/i);
  });

  it("splits investigator footer into nominative position lines", () => {
    const fields = mergeLostMilitaryIdFields(
      createLostMilitaryIdFields(null, buildPersonSummary(null)),
      {
        investigatorPosition:
          "заступник командира роти вогневої підтримки з психологічної підтримки персоналу",
      },
    );
    expect(investigatorFooterLines(fields)).toEqual([
      "Заступник командира роти вогневої підтримки",
      "з психологічної підтримки персоналу:",
    ]);
  });

  it("fills approval block from signatory settings", () => {
    const parts = splitLostMilitaryIdSignatory({
      blockType: "APPROVAL",
      title: "Командир військової частини А4862\nстарший лейтенант",
      rank: "",
      fullName: "Дмитро СЕМЕНЮК",
    });
    expect(parts.titleLines).toEqual(["Командир військової частини А4862"]);
    expect(parts.rank).toBe("старший лейтенант");
    expect(parts.fullName).toMatch(/СЕМЕНЮК/i);
  });

  it("formats approval date from order date", () => {
    const fields = mergeLostMilitaryIdFields(
      createLostMilitaryIdFields(null, buildPersonSummary(null)),
      { orderDate: "15.11.2025" },
    );
    expect(actApprovalDateLine(fields)).toBe("«15»  листопада  2025 року");
  });
});
