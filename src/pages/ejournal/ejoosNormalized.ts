import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { buildEjoosLiveView } from "./ejoosLiveViews";

export type EjoosNormPerson = {
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  dayCode: string;
  isVacant: boolean;
};

export type EjoosNormAbsence = {
  personId: string;
  fullName: string;
  positionIndex: string;
  ground: string;
  place: string;
  departDate: string;
  actualReturn: string;
  isOpen: boolean;
};

export type EjoosNormTimesheet = {
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  day: number;
  dayLabel: string;
  code: string;
};

export type EjoosNormalizedSnapshot = {
  unitLabel: string;
  versionId?: string;
  asOfDate?: string | null;
  persons: EjoosNormPerson[];
  absences: EjoosNormAbsence[];
  timesheet: EjoosNormTimesheet[];
};

export const buildNormalizedSnapshotFromWorkbook = (input: {
  workbook: ExcelWorkbookSnapshot;
  unitLabel?: string;
  versionId?: string;
  asOfDate?: string | null;
  pbFileName?: string | null;
}): EjoosNormalizedSnapshot => {
  const view = buildEjoosLiveView({
    workbook: input.workbook,
    asOfDate: input.asOfDate,
    pbFileName: input.pbFileName,
  });

  return {
    unitLabel: input.unitLabel || "1ПБ",
    versionId: input.versionId,
    asOfDate: input.asOfDate || view.timesheetDayLabel,
    persons: view.roster.map((row) => ({
      personId: row.personId,
      fullName: row.fullName,
      rank: row.rank,
      positionIndex: row.positionIndex,
      dayCode: row.dayCode,
      isVacant: row.isVacant,
    })),
    absences: [
      ...view.absentsOpen.map((row) => ({
        personId: row.personId,
        fullName: row.fullName,
        positionIndex: row.positionIndex,
        ground: row.ground,
        place: row.place,
        departDate: row.departDate,
        actualReturn: row.actualReturn,
        isOpen: true,
      })),
      ...view.absentsClosed.map((row) => ({
        personId: row.personId,
        fullName: row.fullName,
        positionIndex: row.positionIndex,
        ground: row.ground,
        place: row.place,
        departDate: row.departDate,
        actualReturn: row.actualReturn,
        isOpen: false,
      })),
    ],
    timesheet: view.timesheet.map((row) => ({
      personId: row.personId,
      fullName: row.fullName,
      rank: row.rank,
      positionIndex: row.positionIndex,
      day: view.timesheetDay,
      dayLabel: view.timesheetDayLabel,
      code: row.dayValue,
    })),
  };
};
