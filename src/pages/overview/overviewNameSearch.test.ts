import { describe, expect, it } from "vitest";
import {
  buildOverviewRowSearchText,
  overviewNameMatchesQuery,
  parseOverviewNameQueries,
} from "./overviewNameSearch";

describe("parseOverviewNameQueries", () => {
  it("merges surname and given name split by mobile paste newline", () => {
    expect(parseOverviewNameQueries("Рубановський\nГеоргій")).toEqual([
      "Рубановський Георгій",
    ]);
    expect(parseOverviewNameQueries("Рубановський; Георгій")).toEqual([
      "Рубановський Георгій",
    ]);
  });

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

  it("strips rank and callsign columns from a staff-table paste", () => {
    const queries = parseOverviewNameQueries(
      [
        "РОДРІГЕС ФОНСЕКА Ігор Лукас\tТОГІ",
        "ПАЛАЦОЛЛІ Габріел Енріке\tСМАЙЛ",
        "АЛЧАПАР ДЕ ОЛІВЕЙРА Луан Леннон\tАУЧАПА",
        "ДЕ ХЕСУС ЛІМА АВЕЛЛАР Жосе Еверальдо\tАВЕЛАР",
        "старший солдат\tКОВАЛЬ Анатолій Анатолійович\tБЕР",
      ].join("\n"),
    );
    expect(queries).toEqual([
      "РОДРІГЕС ФОНСЕКА Ігор Лукас",
      "ПАЛАЦОЛЛІ Габріел Енріке",
      "АЛЧАПАР ДЕ ОЛІВЕЙРА Луан Леннон",
      "ДЕ ХЕСУС ЛІМА АВЕЛЛАР Жосе Еверальдо",
      "КОВАЛЬ Анатолій Анатолійович",
    ]);
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
    expect(
      overviewNameMatchesQuery(
        "МОРОЗОВ Віктор Вікторович",
        "МОРОЗОВ Олексій Вікторович",
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

  it("does not match a different given name with the same surname", () => {
    expect(
      overviewNameMatchesQuery(
        "СЕРГІЄНКО Фелікс Миколайович",
        "СЕРГІЄНКО Денис Олегович",
      ),
    ).toBe(false);
  });

  it("requires every token for a two-word name query", () => {
    expect(
      overviewNameMatchesQuery(
        "РУБАНОВСЬКИЙ ГЕОРГІЙ ВОЛОДИМИРОВИЧ",
        "Рубановський Георгій",
      ),
    ).toBe(true);
    expect(
      overviewNameMatchesQuery(
        "ІВАНОВ ГЕОРГІЙ ПЕТРОВИЧ",
        "Рубановський Георгій",
      ),
    ).toBe(false);
  });

  it("finds long foreign names even with rank or callsign in the query", () => {
    expect(
      overviewNameMatchesQuery(
        "РОДРІГЕС ФОНСЕКА Ігор Лукас",
        "РОДРІГЕС ФОНСЕКА Ігор Лукас ТОГІ",
      ),
    ).toBe(true);
    expect(
      overviewNameMatchesQuery(
        "ДЕ ХЕСУС ЛІМА АВЕЛЛАР Жосе Еверальдо",
        "ДЕ ХЕСУС ЛІМА АВЕЛЛАР Жосе Еверальдо",
      ),
    ).toBe(true);
    expect(
      overviewNameMatchesQuery(
        "КОВАЛЬ Анатолій Анатолійович",
        "старший солдат КОВАЛЬ Анатолій Анатолійович БЕР",
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
    expect(
      buildOverviewRowSearchText(
        { name: "КОВАЛЬ Анатолій", callSign: "БЕР" },
        "",
      ),
    ).toContain("бер");
  });
});
