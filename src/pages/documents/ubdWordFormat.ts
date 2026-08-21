/**
 * Форматування рапорту УБД.
 *
 * Редагуйте лише цей файл і оновіть сторінку рапорту.
 * Програма підставить числа в Word-шаблон.
 *
 * Одиниці:
 * - відступи / поля / ширини — twips (1 мм ≈ 56.7, зручно: mm(15))
 * - шрифт — half-points (14 пт = 28, зручно: fontPt(14))
 * - підпис-картинка — EMU (1 мм = 36000, зручно: emuMm(27.5))
 */

/** Міліметри → twips */
export const mm = (value: number) => Math.round(value * 56.7);

/** Сантиметри → twips */
export const cm = (value: number) => mm(value * 10);

/** Пункти → twips (інтервал рядка, spacing before/after) */
export const pt = (value: number) => Math.round(value * 20);

/** Пункти шрифту → half-points (w:sz) */
export const fontPt = (value: number) => Math.round(value * 2);

/** Міліметри → EMU (розмір/зсув картинки підпису) */
export const emuMm = (value: number) => Math.round(value * 36000);

export const ubdWordFormat = {
  page: {
    /** A4 ширина */
    width: 11906,
    /** A4 висота. Не зменшуйте — інакше Word поріже текст на короткі сторінки. */
    height: 16838,
    marginTop: mm(15),
    marginRight: mm(15),
    marginBottom: mm(11),
    marginLeft: mm(30),
    headerDistance: 709,
    footerDistance: 709,
    gutter: 0,
    columnSpace: 720,
    linePitch: 360,
    defaultTabStop: 708,
  },

  font: {
    name: "Times New Roman",
    /** Основний текст */
    body: fontPt(14),
    /** Заголовки колонок таблиці */
    tableHeader: fontPt(12),
    /** Міжрядковий інтервал: 240 = одинарний */
    line: 240,
    /** Колір тексту (РНОКПП теж чорний, не червоний) */
    color: "000000",
  },

  classification: {
    /** Відступ зліва для «Відкрита інформація» */
    indentLeft: 5815,
    spacingBefore: 74,
    spacingAfter: 40,
  },

  commander: {
    spacingBefore: pt(12),
    spacingAfter: 0,
  },

  title: {
    spacingBefore: 0,
    spacingAfter: 0,
  },

  body: {
    spacingBefore: 0,
    spacingAfter: 0,
  },

  table: {
    /** Видимі 7 колонок */
    columns: [1300, 2097, 1134, 1453, 1099, 1134, 1276],
    /** Службова 8-ма колонка еталона (не видима) */
    phantomColumn: 2114,
    indent: 113,
    borderSize: 4,
    /** Було 870 / 1005 — нижчі заголовки, щоб рапорт вміщався на 1 сторінку. */
    headerRow1Height: mm(10),
    headerRow2Height: mm(10),
    /** Як в еталоні (416). Не ставте 1500+ — виштовхує «ЗАТВЕРДЖУЮ» на 2-гу сторінку. */
    dataRowHeight: 416,
    cellPaddingTop: 0,
    cellPaddingBottom: 0,
    cellPaddingLeft: 108,
    cellPaddingRight: 108,
  },

  basis: {
    indentLeft: 142,
    firstLine: 426,
    spacingBefore: 0,
    spacingAfter: 0,
  },

  annex: {
    indentLeft: 142,
    firstLine: 426,
    spacingBefore: 0,
    spacingAfter: 0,
  },

  signature: {
    /** 0 = ширина рядка підпису = ширина тексту сторінки */
    nameTabRight: 0,
    spacingAfter: 0,
    imageOffsetX: 0,
    imageOffsetY: 0,
    /**
     * Картинка в комірці таблиці (не absolute).
     * imageWidth: 0 — зберегти пропорції, підігнати лише висоту.
     */
    imageWidth: 0,
    imageHeight: emuMm(24),
    /** Відступ після «Додаток» перед блоком підпису */
    spacingBefore: pt(10),
  },

  approval: {
    spacingBefore: pt(2),
    spacingAfter: 0,
    indentRight: 1132,
  },

  page2: {
    /** Word сам переносить «ВІДКРИТА ІНФОРМАЦІЯ» за вмістом рапорту. */
    pageBreakBefore: false,
    titleSpacingBefore: 9,
    titleIndentRight: 276,
    noteIndentLeft: 20,
  },
};

export type UbdWordFormat = typeof ubdWordFormat;
