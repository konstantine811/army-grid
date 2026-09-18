import { describe, expect, it } from "vitest";
import { createEmptyPreAnketaForm } from "./preAnketaFields";
import {
  buildPreAnketaOcrProposals,
  mapOcrFieldsToPreAnketaValues,
} from "./preAnketaOcr";

describe("preAnketaOcr", () => {
  it("maps direct fullName from Gemini gap-fill response", () => {
    const mapped = mapOcrFieldsToPreAnketaValues([
      {
        key: "fullName",
        label: "П.І.Б.",
        value: "КОВАЛЕНКО Олексій Петрович",
        confidence: "high",
        source: "паспорт",
      },
    ]);
    expect(mapped.fullName?.value).toBe("КОВАЛЕНКО Олексій Петрович");
  });

  it("builds fullName from passport name parts", () => {
    const mapped = mapOcrFieldsToPreAnketaValues([
      { key: "surname", label: "Прізвище", value: "ІВАНОВ", confidence: "high" },
      { key: "firstName", label: "Ім'я", value: "Іван", confidence: "high" },
      {
        key: "patronymic",
        label: "По батькові",
        value: "Іванович",
        confidence: "high",
        source: "паспорт",
      },
    ]);
    expect(mapped.fullName?.value).toBe("ІВАНОВ Іван Іванович");
    expect(mapped.fullName?.source).toBe("паспорт");
  });

  it("proposes fullName when the form field is empty", () => {
    const proposals = buildPreAnketaOcrProposals(
      [
        {
          key: "fullName",
          label: "П.І.Б.",
          value: "ПЕТРЕНКО Петро Петрович",
          confidence: "high",
        },
      ],
      createEmptyPreAnketaForm(),
    );
    expect(proposals).toEqual([
      expect.objectContaining({
        field: "fullName",
        value: "ПЕТРЕНКО Петро Петрович",
        selected: true,
      }),
    ]);
  });
});
