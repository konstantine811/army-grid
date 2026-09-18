import { describe, expect, it } from "vitest";
import {
  personnelSearchEmptyHint,
  personnelSearchMatchesQuery,
  personnelSearchMatchesWildcardQuery,
} from "./personnelSearch";

describe("personnelSearchMatchesQuery", () => {
  it("does not reduce a full name query to surname only", () => {
    expect(
      personnelSearchMatchesQuery(
        "сергієнко фелікс миколайович дельфін",
        "сергієнко денис олегович",
      ),
    ).toBe(false);
  });

  it("matches all words in a multi-word query", () => {
    expect(
      personnelSearchMatchesQuery(
        "сергієнко денис олегович",
        "сергієнко денис олегович",
      ),
    ).toBe(true);
  });

  it("still supports a surname-only query", () => {
    expect(
      personnelSearchMatchesQuery(
        "сергієнко фелікс миколайович",
        "сергієнко",
      ),
    ).toBe(true);
  });

  it("matches Ukrainian і when the query was typed with и", () => {
    expect(
      personnelSearchMatchesQuery(
        "ліга андрій петрович архів",
        "лига",
      ),
    ).toBe(true);
  });

  it("finds ЗАКАЛЮЖНИЙ when the query uses і", () => {
    expect(
      personnelSearchMatchesQuery(
        "солдат закалюжний іван олегович 771",
        "закалюжній іван олегович",
      ),
    ).toBe(true);
  });

  it("matches surname prefixes with * like Word substring search", () => {
    expect(
      personnelSearchMatchesQuery("абрамкін костянтин", "а*", {
        primaryText: "абрамкін костянтин",
      }),
    ).toBe(true);
    expect(
      personnelSearchMatchesQuery("бондаренко іван", "а*", {
        primaryText: "бондаренко іван",
      }),
    ).toBe(false);
    expect(
      personnelSearchMatchesQuery("аліїєв рустам", "а*", {
        primaryText: "аліїєв рустам",
      }),
    ).toBe(true);
  });

  it("matches * anywhere in the PIB cell like Word", () => {
    expect(
      personnelSearchMatchesQuery("абрамкін костянтин", "кос*", {
        primaryText: "абрамкін костянтин",
      }),
    ).toBe(true);
    expect(
      personnelSearchMatchesQuery("рябко сергій анатолійович", "а*", {
        primaryText: "рябко сергій анатолійович",
      }),
    ).toBe(false);
    expect(
      personnelSearchMatchesQuery("рябко сергій анатолійович", "*ол*", {
        primaryText: "рябко сергій анатолійович",
      }),
    ).toBe(true);
  });

  it("matches Latin keyboard A* as Cyrillic А*", () => {
    expect(
      personnelSearchMatchesQuery("абрамкін костянтин", "A*", {
        primaryText: "абрамкін костянтин",
      }),
    ).toBe(true);
  });

  it("matches Word-style ? in the middle of a pattern", () => {
    expect(
      personnelSearchMatchesQuery("костянтин", "к?с", {
        primaryText: "костянтин",
      }),
    ).toBe(true);
    expect(
      personnelSearchMatchesQuery("абрамкін костянтин", "к?т", {
        primaryText: "абрамкін костянтин",
      }),
    ).toBe(false);
  });

  it("finds imya fragments without wildcards", () => {
    expect(
      personnelSearchMatchesQuery("абрамкін костянтин", "кос"),
    ).toBe(true);
  });

  it("supports ? as one character anywhere in PIB or call sign", () => {
    expect(
      personnelSearchMatchesWildcardQuery("аб б", "а?", {
        primaryText: "аб б",
      }),
    ).toBe(true);
    expect(
      personnelSearchMatchesQuery(
        "галінецьких олег валерійович",
        "а?",
        {
          primaryText: "галінецьких олег валерійович",
          callSignText: "ас",
        },
      ),
    ).toBe(true);
  });

  it("suggests * when a single-letter ? mask finds nothing", () => {
    expect(personnelSearchEmptyHint("а?")).toContain("А*");
    expect(personnelSearchEmptyHint("A?")).toContain("А*");
    expect(personnelSearchEmptyHint("абрамкін")).toBe("");
  });

  it("does not let ? match a space between surname and first name", () => {
    expect(
      personnelSearchMatchesQuery("кіяненко андрій олександрович", "к?я", {
        primaryText: "кіяненко андрій олександрович",
      }),
    ).toBe(true);
    expect(
      personnelSearchMatchesQuery("челак ян иванович", "к?я", {
        primaryText: "челак ян иванович",
      }),
    ).toBe(false);
  });

  it("matches call sign patterns with *", () => {
    expect(
      personnelSearchMatchesQuery("рябко сергій", "св*", {
        primaryText: "рябко сергій",
        callSignText: "свін",
      }),
    ).toBe(true);
  });
});
