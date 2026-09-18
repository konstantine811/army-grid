import { describe, expect, it } from "vitest";
import {
  looksLikeConscriptionOffice,
  parseAnketaConscriptionText,
} from "./anketaConscription";

describe("parseAnketaConscriptionText", () => {
  it("splits RTCC office and date from one handwritten line", () => {
    expect(
      parseAnketaConscriptionText(
        "Калинівський РТЦК та СП 14.11.2023 Калинівський РТЦК та СП 14.11.2023",
      ),
    ).toEqual({
      when: "14.11.2023",
      by: "Калинівський РТЦК та СП",
      militaryId: "",
    });
  });

  it("does not treat military id series as the recruiting office", () => {
    expect(parseAnketaConscriptionText("АГ 040571")).toEqual({
      when: "",
      by: "",
      militaryId: "АГ 040571",
    });
  });

  it("keeps office, date and ticket when they are on one line", () => {
    expect(
      parseAnketaConscriptionText(
        "АГ 040571 Калинівський РТЦК та СП 14.11.2023",
      ),
    ).toEqual({
      when: "14.11.2023",
      by: "Калинівський РТЦК та СП",
      militaryId: "АГ 040571",
    });
  });

  it("detects recruiting-office text", () => {
    expect(looksLikeConscriptionOffice("Калинівський РТЦК та СП")).toBe(true);
    expect(looksLikeConscriptionOffice("слюсар ТОВ Промінь")).toBe(false);
  });
});
