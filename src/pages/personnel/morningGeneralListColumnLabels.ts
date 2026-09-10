/**
 * Назви колонок «1.ОС Загальний список» за Excel-індексом (1-based).
 * Окремий модуль без залежностей — уникає circular import між staffSheet і personnelUtils.
 */
export const MORNING_GENERAL_LIST_COLUMN_LABELS: Record<number, string> = {
  1: "№",
  2: "Підрозділ",
  3: "Взвод",
  4: "Відділення",
  5: "Посада",
  6: "ВОС",
  7: "Повна посада",
  8: "ШПК факт",
  9: "Категорія складу",
  10: "Анкета",
  11: "Військовий квиток",
  12: "Мобілізація/контракт",
  13: "Звання",
  14: "ПІБ",
  15: "Позивний",
  16: "Дата народження",
  17: "Рік",
  18: "Повних років",
  19: "ІПН",
  20: "Група крові",
  21: "Статус",
  22: "Тип В\\С",
  23: "Статус БГ",
  24: "БЗВП/БРЕЗ",
  25: "Наявність БЗВП",
  26: "Курс БЗВП",
  27: "Відрядження (БРЕЗ)",
  28: "Обмеження",
  29: "В якому підрозділі",
  31: "Місце перебування",
  32: "Примітки",
  33: "Напрямок",
  34: "Примітка 3",
  35: "Місце перебування (уточнення)",
  // Колонки без заголовка в Excel (часто списки / дублікаты значень)
  37: "Статус",
  38: "Тип В\\С",
  39: "БЗВП/БРЕЗ",
  40: "Місце перебування",
  41: "Обмеження",
  42: "Статус БГ",
};

const parseGenericRosterColumnNumber = (key: string) => {
  const match = key.trim().match(/^column_(\d+)(?:_\d+)?$/i);
  if (!match) return null;
  const number = Number(match[1]);
  return Number.isFinite(number) && number > 0 ? number : null;
};

export const resolveMorningGeneralListColumnLabel = (
  sourceKey: string,
  fallback = "",
) => {
  const columnNumber = parseGenericRosterColumnNumber(sourceKey);
  if (columnNumber != null) {
    const known = MORNING_GENERAL_LIST_COLUMN_LABELS[columnNumber];
    if (known) return known;
    return fallback || `Колонка ${columnNumber}`;
  }
  return fallback;
};
