import { describe, expect, it } from "vitest";
import { matchesQuestionnairePerson } from "./personnelQuestionnaireFile";

describe("questionnaire person matching", () => {
  const name = "ШЕВЧЕНКО Олександр Володимирович (11.05.1981 р.н.)";

  it("requires the full name and birth date, even for the stored filename", () => {
    expect(matchesQuestionnairePerson({ fileName: "ШЕВЧЕНКО Олександр Володимирович (МОТО).pdf" }, name)).toBe(false);
    expect(matchesQuestionnairePerson({ fileName: "ШЕВЧЕНКО Олександр Володимирович (11.05.1982 р.н.).pdf" }, name)).toBe(false);
    expect(matchesQuestionnairePerson({ fileName: "ШЕВЧЕНКО Олександр Васильович (11.05.1981 р.н.).pdf" }, name)).toBe(false);
    expect(matchesQuestionnairePerson({ fileName: `${name}.pdf` }, name)).toBe(true);
  });

  it("finds other people without requiring a date absent from their name", () => {
    expect(matchesQuestionnairePerson({ fileName: "ШЕВЧЕНКО Олександр Володимирович (МОТО).pdf" }, "ШЕВЧЕНКО Олександр Володимирович")).toBe(true);
    expect(matchesQuestionnairePerson({ fileName: "КОВАЛЬ Іван Петрович.pdf" }, "КОВАЛЬ Іван Петрович")).toBe(true);
  });

  it("accepts the date already in the name from a containing directory", () => {
    expect(matchesQuestionnairePerson({ fileName: "шевченко_олександр_володимирович.pdf", relativePath: "11.05.1981/шевченко_олександр_володимирович.pdf" }, name)).toBe(true);
  });
});
