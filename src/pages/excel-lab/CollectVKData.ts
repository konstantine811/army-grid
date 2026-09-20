import type { EntityEJOOSSheets } from "./ExcelLabEJOOS";
import type { EntityCardsSheets, ExcelLabStateCards } from "./ExcelLabMilitaryServiceCards";
import type { EntityStateSheets } from "./ExcelLabStateParser";
import type { ExcelLabProcessedData } from "./parseUploadedExcel";

export type CollectedVKPerson = {
  staff: EntityStateSheets[string];
  ejoos: EntityEJOOSSheets[string] | undefined;
  militaryCards: EntityCardsSheets[string] | undefined;
};

export type CollectedVKSheets = Record<string, CollectedVKPerson>;

export const collectVKData = (
  cards: ExcelLabProcessedData,
): CollectedVKSheets => {
  const data: CollectedVKSheets = {};

  Object.entries(cards.staff).forEach(([key, staff]) => {
    data[key] = {
      staff,
      ejoos: cards.ejoos[key],
      militaryCards: cards.militaryCards[key],
    };
  });

  console.log("[excel-lab] collected vk data:", data);
  return data;
};
