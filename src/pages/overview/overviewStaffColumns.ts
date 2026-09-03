import { FIGHTER_STATUS_FALLBACK_HEADERS } from "../excel-fill/staffSheet";
import { MORNING_GENERAL_LIST_COLUMN_LABELS } from "../personnel/personnelUtils";

const rosterCol = (columnNumber: number) =>
  MORNING_GENERAL_LIST_COLUMN_LABELS[columnNumber] ?? "";

/** Заголовки колонок Огляду = назви з «1.ОС Загальний список» та «Статус бійців». */
export const OVERVIEW_STAFF_COLUMN_HEADERS = {
  name: rosterCol(14),
  unit: rosterCol(2),
  rank: rosterCol(13),
  positionTitle: rosterCol(7),
  status: rosterCol(21),
  questionnaire: rosterCol(10),
  fighterDirection: rosterCol(33) || FIGHTER_STATUS_FALLBACK_HEADERS[1],
  fighterExitDate: FIGHTER_STATUS_FALLBACK_HEADERS[6],
  fighterReturnDate: FIGHTER_STATUS_FALLBACK_HEADERS[7],
  fighterTotalDays: FIGHTER_STATUS_FALLBACK_HEADERS[8],
  fighterStatus: "Статус 200/300/500",
} as const;
