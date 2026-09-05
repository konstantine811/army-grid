import { describe, expect, it } from "vitest";
import { personnelSearchMatchesQuery } from "./personnelSearch";

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
});
