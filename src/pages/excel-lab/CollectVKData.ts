import type { EntityKSPVKSheets } from "./ExcelKSPVKData";
import type { EntityEJOOSSheets } from "./ExcelLabEJOOS";
import type { EntityCardsSheets } from "./ExcelLabMilitaryServiceCards";
import type { EntityStateSheets } from "./ExcelLabStateParser";
import type { ExcelLabProcessedData } from "./parseUploadedExcel";

export type CollectedVKPerson = {
  staff: EntityStateSheets[string];
  ejoos: EntityEJOOSSheets[string] | undefined;
  militaryCards: EntityCardsSheets[string] | undefined;
  ksp: EntityKSPVKSheets[string] | undefined;
};

export type CollectedVKSheets = Record<string, CollectedVKPerson>;

export const collectVKData = (
  cards: ExcelLabProcessedData,
): CollectedVKSheets => {
  const data: CollectedVKSheets = {};
  const dataNotVK: CollectedVKSheets = {};
  const dataVKOnKSP: CollectedVKSheets = {};
  const dataVKOnMilitaryCardsAndKSP: CollectedVKSheets = {};
  const dataVKOnMilitaryCards: CollectedVKSheets = {};
  Object.entries(cards.staff).forEach(([key, staff]) => {
    if (
      (!cards.ejoos[key] || cards.ejoos[key].cardNumber === "дані відсутні") &&
      !cards.militaryCards[key] &&
      !cards.ksp[key] &&
      staff.status !== "СЗЧ"
    ) {
      dataNotVK[key] = {
        staff,
        ejoos: cards.ejoos[key],
        militaryCards: cards.militaryCards[key],
        ksp: cards.ksp[key],
      };
    }
    if (cards.ksp[key] && !cards.militaryCards[key]) {
      dataVKOnKSP[key] = {
        staff,
        ejoos: cards.ejoos[key],
        militaryCards: cards.militaryCards[key],
        ksp: cards.ksp[key],
      };
    }
    if (cards.ksp[key] && cards.militaryCards[key]) {
      dataVKOnMilitaryCardsAndKSP[key] = {
        staff,
        ejoos: cards.ejoos[key],
        militaryCards: cards.militaryCards[key],
        ksp: cards.ksp[key],
      };
    }

    if (cards.militaryCards[key]) {
      dataVKOnMilitaryCards[key] = {
        staff,
        ejoos: cards.ejoos[key],
        militaryCards: cards.militaryCards[key],
        ksp: cards.ksp[key],
      };
    }

    data[key] = {
      staff,
      ejoos: cards.ejoos[key],
      militaryCards: cards.militaryCards[key],
      ksp: cards.ksp[key],
    };
  });

  console.log("[excel-lab] collected data not vk:", dataNotVK);
  console.log(
    "[excel-lab] collected vk data on military cards:",
    dataVKOnMilitaryCards,
  );
  console.log("[excel-lab] collected vk data:", data);
  console.log("[excel-lab] collected vk data on ksp:", dataVKOnKSP);
  console.log(
    "[excel-lab] collected vk data on military cards and ksp:",
    dataVKOnMilitaryCardsAndKSP,
  );
  return data;
};

export const describeCollectedVKData = (data: CollectedVKSheets) =>
  `Оброблено осіб: ${Object.keys(data).length}. Редагуй collectVKData у CollectVKData.ts.`;
