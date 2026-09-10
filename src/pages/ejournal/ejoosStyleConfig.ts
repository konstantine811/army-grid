/**
 * Єдине місце для налаштування шрифту записів у ЄЖООС.
 *
 * Після змін:
 *   npm test -- src/pages/ejournal/ejoosTimesheetStyle.test.ts
 *   npm run ejoos:style-probe -- <файл.xlsx> --sheet шпо --rows 95-99
 */
export const EJOOS_WRITE_STYLE = {
  fontName: "Times New Roman",
  fontSize: 12,
  /** Записи даних ніколи не робимо bold — лише шаблонні заголовки лишаються як є. */
  bold: false,
} as const;
