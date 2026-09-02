import { isInternalStaffIndexHop } from "./ejoosMovementRules";

export type TimesheetEpisodeMove = {
  type: string;
  destination: string;
  changeText: string;
  previousIndex?: string;
  nextIndex?: string;
  note: string;
};

/**
 * З якої дати фарбувати активний рядок Табеля.
 * Внутрішня зміна посади (і повернення на стару) не починає новий епізод.
 */
export const resolveTimesheetEpisodeStart = (input: {
  monthStartLabel: string;
  appointmentDate: string;
  inboundDate: string;
  hasMonthStartAbsence: boolean;
  hasDepartureEvidence: boolean;
  leftUnitThisMonth: boolean;
  ownUnitMoves: TimesheetEpisodeMove[];
}) => {
  if (input.hasMonthStartAbsence) return input.monthStartLabel;
  if (input.leftUnitThisMonth || input.hasDepartureEvidence) {
    return input.appointmentDate || input.inboundDate || "";
  }
  const hops = input.ownUnitMoves;
  if (!hops.length || hops.every(isInternalStaffIndexHop)) {
    return input.monthStartLabel;
  }
  return input.appointmentDate || input.inboundDate || "";
};
