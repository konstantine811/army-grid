import { describe, expect, it } from "vitest";
import {
  overviewNameMatchesQuery,
  parseOverviewNameQueries,
} from "./overviewNameSearch";

describe("parseOverviewNameQueries", () => {
  it("keeps seven pasted names", () => {
    const queries = parseOverviewNameQueries(`ДОРОШЕНКО Дмитро Сергійович
ЛОБОДА Олексій Миколайович
ГОМЕНЮК Віктор Савич
ПОЛЬОВИЙ Олексій Геннадійович
ВЕРЕНКО Роман Григорович
БОРИШПОЛЕЦЬ Роман Юрійович
МОРОЗОВ Віктор Вікторович (27.03.1988 р.н.)`);
    expect(queries).toHaveLength(7);
    expect(queries.some((name) => /БОРИШПОЛЕЦЬ/i.test(name))).toBe(true);
  });
});

describe("overviewNameMatchesQuery", () => {
  it("finds Боришполець when the roster dropped the soft sign", () => {
    expect(
      overviewNameMatchesQuery(
        "БОРИШПОЛЕЦ Роман Юрійович",
        "БОРИШПОЛЕЦЬ Роман Юрійович",
      ),
    ).toBe(true);
  });

  it("does not treat Віктор as Вікторович", () => {
    expect(
      overviewNameMatchesQuery(
        "МОРОЗОВ Олексій Вікторович",
        "МОРОЗОВ Віктор Вікторович",
      ),
    ).toBe(false);
  });

  it("still matches a truncated given name", () => {
    expect(
      overviewNameMatchesQuery(
        "ВЕРЕНКО Роман Григорович",
        "ВЕРЕНКО Ром Григорович",
      ),
    ).toBe(true);
  });
});
