export type PbShPerson = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  positionTitle: string;
  status: string;
  /** Якщо є в sh — «Звідки прибув». */
  arrivedFrom: string;
};

export type PbArchivePeriod = {
  excelRow: number;
  periodNumber: string;
  personId: string;
  fullName: string;
  rank: string;
  positionTitle: string;
  absenceType: string;
  departDate: string;
  place: string;
  orderNumber: string;
  orderDate: string;
  plannedReturn: string;
  /** Фактична / дата повернення, якщо є. */
  returnDate: string;
  returnOrderNumber: string;
  returnOrderDate: string;
};

export type PbMovement = {
  excelRow: number;
  movementNumber: string;
  type: string;
  personId: string;
  fullName: string;
  rank: string;
  previousIndex: string;
  nextIndex: string;
  destination: string;
  orderNumber: string;
  orderDate: string;
  basisNumber: string;
  basisDate: string;
  changeText: string;
  status: string;
  note: string;
  /** Колонка «Звідки» / «Звідки прибув», якщо є в Рух. */
  arrivedFrom: string;
};
