import { describe, expect, it } from "vitest";
import { normalizeAnketaPlaceText } from "./anketaPlaceFormat";

describe("normalizeAnketaPlaceText", () => {
  it("orders settlement, district, region and puts phone last", () => {
    expect(
      normalizeAnketaPlaceText(
        "Дніпропетровська обл Софіївський район смт Софіївка +380 662033264",
      ),
    ).toBe(
      "смт Софіївка, Софіївський р-н, Дніпропетровська обл, тел: +380662033264",
    );
    expect(
      normalizeAnketaPlaceText(
        "Івано-Франківська обл Калуський район м.Калуш",
      ),
    ).toBe("м. Калуш, Калуський р-н, Івано-Франківська обл");
  });

  it("keeps city street details after the settlement", () => {
    expect(
      normalizeAnketaPlaceText("м. Павлоград Радянська 87 кв 13"),
    ).toBe("м. Павлоград, вул. Радянська, буд. 87, кв. 13");
  });

  it("title-cases a lowercase city and boulevard", () => {
    expect(
      normalizeAnketaPlaceText(
        "м. дніпро, бульвар європейський, буд. 2, кв. 43",
      ),
    ).toBe("м. Дніпро, бульвар Європейський, буд. 2, кв. 43");
  });

  it("does not rewrite military or free-text locations", () => {
    expect(normalizeAnketaPlaceText("ППД Вишневе")).toBe("ППД Вишневе");
    expect(normalizeAnketaPlaceText("2 ШБ")).toBe("2 ШБ");
  });
});
