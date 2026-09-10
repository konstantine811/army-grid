import { describe, expect, it } from "vitest";
import {
  buildOverviewRowSearchText,
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

describe("buildOverviewRowSearchText", () => {
  it("normalizes searchable fields once", () => {
    const text = buildOverviewRowSearchText(
      {
        name: "ПОЛЬОВИЙ Олексій",
        rank: "солдат",
        unit: "1 рота",
        statusLabel: "В строю",
      },
      "Форма 6",
    );
    expect(text).toContain("польовий");
    expect(text).toContain("форма");
    expect(text).toContain("6");
  });
});
