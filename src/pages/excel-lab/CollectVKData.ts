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

export type CollectedVKData = {
  all: CollectedVKSheets;
  notVK: CollectedVKSheets;
  vkOnMilitaryCards: CollectedVKSheets;
  vkOnKSP: CollectedVKSheets;
  vkOnMilitaryCardsAndKSP: CollectedVKSheets;
};

export const collectVKData = (
  cards: ExcelLabProcessedData,
): CollectedVKData => {
  const data: CollectedVKSheets = {};
  const dataNotVK: CollectedVKSheets = {};
  const dataVKOnKSP: CollectedVKSheets = {};
  const dataVKOnMilitaryCardsAndKSP: CollectedVKSheets = {};
  const dataVKOnMilitaryCards: CollectedVKSheets = {};
  Object.entries(cards.staff).forEach(([key, staff]) => {
    if (
      (!cards.ejoos[key] ||
        cards.ejoos[key].cardNumber === "дані відсутні" ||
        cards.ejoos[key].cardNumber === "відсутній") &&
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
    if (
      cards.ksp[key] &&
      (!cards.militaryCards[key] ||
        !cards.militaryCards[key]?.ВК?.status ||
        !cards.militaryCards[key]?.ДОВІДКИ?.status ||
        !cards.militaryCards[key]?.ТПВ?.status)
    ) {
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

  return {
    all: data,
    notVK: dataNotVK,
    vkOnMilitaryCards: dataVKOnMilitaryCards,
    vkOnKSP: dataVKOnKSP,
    vkOnMilitaryCardsAndKSP: dataVKOnMilitaryCardsAndKSP,
  };
};

export const describeCollectedVKData = (collected: CollectedVKData) =>
  [
    `Усі: ${Object.keys(collected.all).length}`,
    `Без ВК: ${Object.keys(collected.notVK).length}`,
    `ВК у картках: ${Object.keys(collected.vkOnMilitaryCards).length}`,
    `ВК на КСП: ${Object.keys(collected.vkOnKSP).length}`,
    `ВК картки+КСП: ${Object.keys(collected.vkOnMilitaryCardsAndKSP).length}`,
  ].join("\n");
