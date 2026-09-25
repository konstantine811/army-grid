import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { indexKryvyiRihPlaces, isKryvyiRihAreaText } from "./kryvyiRihStaff";

describe("isKryvyiRihAreaText", () => {
  it("matches the city, the raion, and a Kryvyi Rih TCK", () => {
    expect(isKryvyiRihAreaText("м. Кривий Ріг, вул. Відродження буд. 11")).toBe(
      true,
    );
    expect(
      isKryvyiRihAreaText("с. Лозуватка, Криворізький р-н, Дніпропетровська обл."),
    ).toBe(true);
    expect(isKryvyiRihAreaText("Криворізький РТЦК та СП")).toBe(true);
    expect(isKryvyiRihAreaText("м. Апостолове")).toBe(true);
  });

  it("does not treat a street in another city as Kryvyi Rih", () => {
    expect(isKryvyiRihAreaText("вул. Криворізька, 5, м. Дніпро")).toBe(false);
    expect(isKryvyiRihAreaText("м. Київ")).toBe(false);
    expect(isKryvyiRihAreaText("с. Широке, Харківська обл.")).toBe(false);
  });
});

describe("indexKryvyiRihPlaces", () => {
  it("keeps in-staff people whose address or birth place is Kryvyi Rih", () => {
    const inStaff = {
      __dbRowId: "roster:1",
      roster__column_14: "Шевченко Тарас Григорович",
      адреса_проживання: "м. Кривий Ріг, вул. Гагаріна 1",
    } as EjournalPreviewRow;
    const archive = {
      __dbRowId: "oos:2",
      __rosterArchive: true,
      прізвище: "Франко Іван Якович",
      місце_народження: "м. Кривий Ріг",
    } as EjournalPreviewRow;
    const other = {
      __dbRowId: "roster:3",
      roster__column_14: "Леся Українка",
      адреса_проживання: "м. Київ",
    } as EjournalPreviewRow;
    const onlyCalledBy = {
      __dbRowId: "roster:4",
      roster__column_14: "Коцюбинський Михайло Михайлович",
      ким_призваний: "Криворізький РТЦК та СП",
      місце_народження: "м. Львів",
    } as EjournalPreviewRow;
    const registered = {
      __dbRowId: "roster:5",
      roster__column_14: "Франко Іван Якович",
      місце_реєстрації: "м. Кривий Ріг, вул. Святослава 2",
    } as EjournalPreviewRow;

    const index = indexKryvyiRihPlaces([
      inStaff,
      archive,
      other,
      onlyCalledBy,
      registered,
    ]);

    expect(index["шевченко тарас григорович"]).toContain("Кривий Ріг");
    expect(index["франко іван якович"]).toContain("Кривий Ріг");
    expect(index["леся українка"]).toBeUndefined();
    expect(index["коцюбинський михайло михайлович"]).toBeUndefined();
  });
});
