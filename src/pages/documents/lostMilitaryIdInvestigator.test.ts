import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  actApprovalDateLine,
  approvalFooterBlock,
  buildLostMilitaryIdReportText,
  buildManualSignatoryDateLine,
  circumstancesText,
  createLostMilitaryIdFields,
  declinedInvestigator,
  formatReporterTitleLines,
  instrumentalInvestigatorLine,
  investigatorFooterLines,
  investigatorFromPersonnelRow,
  mergeLostMilitaryIdFields,
  militaryUnitLabel,
  normalizeMilitaryUnitPhrase,
  orderFooterBlock,
  splitLostMilitaryIdSignatory,
  usesMovementCircumstances,
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
  }) as unknown as EjournalPreviewRow;

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

  it("uses Взвод when посада is a number", () => {
    const row = person("ІВАНОВ Іван", {
      rank: "лейтенант",
      position: "29",
      взвод: "2 піхотний взвод 1 піхотної роти",
      id: "1",
    });
    expect(investigatorFromPersonnelRow(row).investigatorPosition).toMatch(
      /піхотний взвод/i,
    );
  });

  it("does not turn appointment date into dative position with trailing у", () => {
    const row = person("ІВАНОВ Іван", {
      rank: "лейтенант",
      position: "29.03.2026",
      повна_посада: "29.03.2026",
      id: "1",
    });
    expect(investigatorFromPersonnelRow(row).investigatorPosition).toBe("");
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

  it("clears corrupted date saved as investigator position", () => {
    const defaults = createLostMilitaryIdFields(null, buildPersonSummary(null));
    const merged = mergeLostMilitaryIdFields(defaults, {
      investigatorPosition: "29.03.2026у",
    });
    expect(merged.investigatorPosition).toBe("");
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

  it("builds approval footer from document signatory record", () => {
    const fields = mergeLostMilitaryIdFields(
      createLostMilitaryIdFields(null, buildPersonSummary(null), [
        {
          blockType: "APPROVAL",
          title: "Командир військової частини А4862",
          rank: "",
          fullName: "",
        },
      ]),
      { approvalDate: "29.08.2026" },
    );
    const footer = approvalFooterBlock(fields);
    expect(footer.titleLines).toEqual(["Командир військової частини А4862"]);
    expect(footer.rank).toBe("");
    expect(footer.name).toBe("");
    expect(actApprovalDateLine(fields)).toBe(
      buildManualSignatoryDateLine(new Date(2026, 7, 29)),
    );
  });

  it("builds order footer with full commander data, skipping placeholder signatories", () => {
    const fields = createLostMilitaryIdFields(null, buildPersonSummary(null), [
      {
        blockType: "SIGNER",
        title: "Командир __ взводу __ роти\n__ лейтенант __ (прізвище та ініціали)",
        rank: "",
        fullName: "",
      },
      {
        blockType: "APPROVAL",
        title: "Командир __ взводу __ роти",
        rank: "",
        fullName: "",
      },
    ]);
    const footer = orderFooterBlock(fields);
    expect(footer.titleLines).toEqual([
      "Тимчасово виконуючий обов’язки",
      "командира військової частини А4862",
    ]);
    expect(footer.rank).toBe("капітан");
    expect(footer.name).toMatch(/АДАМОВ/i);
  });

  it("prefers complete approval signatory for order footer", () => {
    const fields = createLostMilitaryIdFields(null, buildPersonSummary(null), [
      {
        blockType: "APPROVAL",
        title: "Командир 1 піхотного батальйону\nвійськової частини А4862",
        rank: "старший лейтенант",
        fullName: "Єгор СИДОРЕНКО",
      },
    ]);
    const footer = orderFooterBlock(fields);
    expect(footer.rank).toBe("старший лейтенант");
    expect(footer.name).toMatch(/СИДОРЕНКО/i);
    expect(footer.titleLines[0]).toMatch(/Командир 1 піхотного батальйону/i);
  });

  it("formats signatory date with day, month and year", () => {
    expect(buildManualSignatoryDateLine(new Date(2026, 8, 8))).toBe(
      "«08»  вересня  2026 року",
    );
  });

  it("strips date lines from multiline reporter title", () => {
    expect(
      formatReporterTitleLines(
        "Командир 1 піхотного батальйону\nвійськової частини А4862\n05.09.2026",
        "старший лейтенант",
      ),
    ).toEqual([
      "Командир 1 піхотного батальйону",
      "військової частини А4862",
    ]);
  });
});

describe("lost military id event circumstances", () => {
  const eventFields = (): ReturnType<typeof createLostMilitaryIdFields> => ({
    ...createLostMilitaryIdFields(
      person("ПЕТРЕНКО Іван Іванович", { rank: "солдат" }),
      buildPersonSummary(
        person("ПЕТРЕНКО Іван Іванович", { rank: "солдат" }),
      ),
    ),
    lossDate: "12.11.2025",
    isExactDate: true,
    circumstanceKind: "custom",
    lossLocation: "с. Гришене",
    customCircumstances:
      "на позицію потрапив каб та FPV дрон і все згоріло",
    searchConducted: false,
  });

  it("builds phrase from place and event details", () => {
    expect(circumstancesText(eventFields())).toBe(
      "с. Гришене, на позицію потрапив каб та FPV дрон і все згоріло",
    );
    expect(usesMovementCircumstances(eventFields())).toBe(false);
  });

  it("weaves event into report text without movement placeholders", () => {
    const text = buildLostMilitaryIdReportText(eventFields());
    expect(text).toContain("12.11.2025");
    expect(text).toContain("с. Гришене");
    expect(text).toContain("FPV дрон");
    expect(text).not.toContain("переміщення з ______");
  });

  it("includes unitLabel in report text after military unit phrase", () => {
    const fields = mergeLostMilitaryIdFields(eventFields(), {
      unitLabel: "2 піхотної роти",
    });
    const text = buildLostMilitaryIdReportText(fields);
    expect(text).toContain("військової частини А4862 2 піхотної роти");
  });

  it("does not wrap evacuation phrase in movement template", () => {
    const fields = mergeLostMilitaryIdFields(eventFields(), {
      circumstanceKind: "movement",
      fromLocation: "під час евакуації військовослужбовця до",
      toLocation: "до ПГХ 66 Петропавлівка",
      customCircumstances: "",
      lossLocation: "",
    });
    expect(circumstancesText(fields)).toBe(
      "під час евакуації військовослужбовця до ПГХ 66 Петропавлівка",
    );
    expect(buildLostMilitaryIdReportText(fields)).toContain(
      "під час евакуації військовослужбовця до ПГХ 66 Петропавлівка",
    );
    expect(buildLostMilitaryIdReportText(fields)).not.toContain("переміщення з");
  });

  it("keeps simple from-to movement template for short place names", () => {
    const fields = mergeLostMilitaryIdFields(eventFields(), {
      circumstanceKind: "movement",
      fromLocation: "с. Гришене",
      toLocation: "ПГХ 66 Петропавлівка",
      customCircumstances: "",
      lossLocation: "",
    });
    expect(circumstancesText(fields)).toBe(
      "під час переміщення з с. Гришене до ПГХ 66 Петропавлівка",
    );
  });

  it("uses full search result sentence without однак wrapper or double period", () => {
    const phrase =
      "місцезнаходження військового квитка після проведення евакуаційних заходів встановити не вдалося.";
    const fields = mergeLostMilitaryIdFields(eventFields(), {
      searchConducted: true,
      searchResult: phrase,
    });
    const text = buildLostMilitaryIdReportText(fields);
    expect(text).toContain(phrase);
    expect(text).not.toContain("однак");
    expect(text).not.toContain("не вдалося..");
  });

  it("requests military id restoration in report closing paragraph", () => {
    const text = buildLostMilitaryIdReportText(eventFields());
    expect(text).toContain(
      "У зв’язку з вищевикладеним прошу організувати проведення необхідних заходів щодо відновлення (оформлення нового) військового квитка",
    );
    expect(text).toContain("замість втраченого.");
    expect(text).not.toContain(
      "прошу призначити службове розслідування за фактом втрати",
    );
  });

  it("wraps short search result fragment in default template", () => {
    const fields = mergeLostMilitaryIdFields(eventFields(), {
      searchConducted: true,
      searchResult: "військовий квиток не знайдено",
    });
    expect(buildLostMilitaryIdReportText(fields)).toContain(
      "однак військовий квиток не знайдено.",
    );
  });
});
