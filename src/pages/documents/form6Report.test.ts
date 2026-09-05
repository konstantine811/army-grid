import { describe, expect, it } from "vitest";
import {
  formatForm6BasisText,
  form6BasisLineForWord,
  mergeForm6Fields,
  parseForm6BasisParts,
  stripForm6BasisLabel,
  type Form6ReportFields,
} from "./form6Report";

const fields = (
  extra: Partial<Form6ReportFields> = {},
): Form6ReportFields => ({
  commander: "Командиру 1 піхотного батальйону",
  fullName: "ВІПУЗОВ Владислав Андрійович",
  rank: "солдат",
  staffPosition: "водій-електрик",
  birthDate: "08.07.1997",
  idDocument: "Паспорт 002532584",
  rnokpp: "3561813015",
  address: "",
  phone: "095 247 65 10",
  taskPeriod: "з 07.06.2026-11.07.2026",
  taskPlace: "н.п. Сергіївка",
  basisNumber: "4862/ОКП/1158/дск",
  basisDate: "14.10.2025",
  basis:
    "Бойове розпорядження командира 425 ОШП «СКЕЛЯ» №4862/ОКП/1158/дск від 14.10.2025",
  date: "01.09.2026",
  formPurpose: "",
  folderName: "1ПБ РАПОРТ Форма 6 · ВІПУЗОВ Владислав Андрійович",
  signatories: [],
  statusNote: "",
  basisManual: false,
  ...extra,
});

describe("Form 6 basis from BR", () => {
  it("formats the СКЕЛЯ line with the picked order number and date", () => {
    expect(formatForm6BasisText("4862/ОКП/1471/дск", "07.06.2026")).toBe(
      "Бойове розпорядження командира 425 ОШП «СКЕЛЯ» №4862/ОКП/1471/дск від 07.06.2026",
    );
  });

  it("parses number and date from an old hardcoded Form 6 basis", () => {
    expect(
      parseForm6BasisParts(
        "Бойове розпорядження командира 425 ОШП «СКЕЛЯ» №4862/ОКП/1158/дск від 14.10.2025",
      ),
    ).toEqual({
      basisNumber: "4862/ОКП/1158/дск",
      basisDate: "14.10.2025",
    });
  });

  it("replaces the October 2025 default with the BR for the task period start", () => {
    const merged = mergeForm6Fields(
      fields({
        taskPeriod: "з 07.06.2026-11.07.2026",
        basisNumber: "",
        basisDate: "",
      }),
      fields(),
    );
    expect(merged.basisDate).toBe("07.06.2026");
    expect(merged.basisNumber).toBe("4862/ОКП/1471/дск");
    expect(merged.basis).toContain("№4862/ОКП/1471/дск");
    expect(merged.basis).toContain("від 07.06.2026");
  });

  it("keeps a manually entered BR instead of replacing it from the task period", () => {
    const merged = mergeForm6Fields(
      fields({
        taskPeriod: "з 07.06.2026-11.07.2026",
      }),
      fields({
        basisManual: true,
        basisNumber: "4862/ОКП/294/дск",
        basisDate: "01.02.2026",
        basis:
          "Бойове розпорядження командира 425 ОШП «СКЕЛЯ» №4862/ОКП/294/дск від 01.02.2026",
      }),
    );
    expect(merged.basisManual).toBe(true);
    expect(merged.basisNumber).toBe("4862/ОКП/294/дск");
    expect(merged.basisDate).toBe("01.02.2026");
    expect(merged.basis).toContain("№4862/ОКП/294/дск");
    expect(merged.basis).toContain("від 01.02.2026");
  });

  it("keeps a manually typed Підстава line without the Підстава: prefix", () => {
    const merged = mergeForm6Fields(
      fields(),
      fields({
        basisManual: true,
        basis:
          "Підстава: Бойове розпорядження командира 1ОШБр №12/дск від 01.02.2026",
      }),
    );
    expect(merged.basis).toBe(
      "Бойове розпорядження командира 1ОШБр №12/дск від 01.02.2026",
    );
  });

  it("puts Підстава: only once in the Word line", () => {
    expect(
      form6BasisLineForWord(
        "Підстава: Бойове розпорядження командира 425 ОШП «СКЕЛЯ» №12 від 01.02.2026",
      ),
    ).toBe(
      "Підстава: Бойове розпорядження командира 425 ОШП «СКЕЛЯ» №12 від 01.02.2026",
    );
  });

  it("does not refill cleared manual BR number and date on reopen", () => {
    const merged = mergeForm6Fields(
      fields({
        taskPeriod: "з 16.06.2014 по 11.09.2014",
        basisNumber: "4862/ОКП/294/дск",
        basisDate: "01.02.2026",
      }),
      fields({
        basisManual: true,
        basisNumber: "",
        basisDate: "",
        basis:
          'шифротелеграма Командувача Сухопутних військ ЗС України від "04" червня 2014 року №05',
      }),
    );
    expect(merged.basisNumber).toBe("");
    expect(merged.basisDate).toBe("");
    expect(merged.basis).toContain("шифротелеграма");
  });

  it("keeps a leading space in the Підстава field", () => {
    expect(stripForm6BasisLabel(" шифротелеграма")).toBe(" шифротелеграма");
    const merged = mergeForm6Fields(
      fields(),
      fields({
        basisManual: true,
        basis: " шифротелеграма",
      }),
    );
    expect(merged.basis).toBe(" шифротелеграма");
    expect(form6BasisLineForWord(merged.basis)).toBe(
      "Підстава: шифротелеграма",
    );
    expect(form6BasisLineForWord("шифротелеграма")).toBe(
      "Підстава: шифротелеграма",
    );
  });
});

describe("Form 6 id document merge", () => {
  it("keeps a valid document RNOKPP when personnel contains a placeholder", () => {
    const merged = mergeForm6Fields(
      fields({ rnokpp: "втрачено" }),
      fields({ rnokpp: "3142223156" }),
    );

    expect(merged.rnokpp).toBe("3142223156");
  });

  it("does not wipe a saved passport number with the default document title", () => {
    const merged = mergeForm6Fields(
      fields({ idDocument: "Паспорт громадянина України" }),
      fields({
        idDocument: "Паспорт громадянина України Серія КВ №001828",
      }),
    );
    expect(merged.idDocument).toBe(
      "Паспорт громадянина України Серія КВ №001828",
    );
  });

  it("uses personnel id document when it already has a number", () => {
    const merged = mergeForm6Fields(
      fields({ idDocument: "Паспорт 998877665" }),
      fields({ idDocument: "Паспорт 002532584" }),
    );
    expect(merged.idDocument).toBe("Паспорт 998877665");
  });

  it("keeps the saved Для чого форма value", () => {
    const merged = mergeForm6Fields(
      fields(),
      fields({ formPurpose: "для соцзахисту" }),
    );
    expect(merged.formPurpose).toBe("для соцзахисту");
  });
});
