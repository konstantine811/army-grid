export type EjoosAbsentRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  ground: string;
  place: string;
  departDate: string;
  actualReturn: string;
};

export type EjoosTimesheetRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  dayValue: string;
};

/** Повний рядок Табеля: штатний або історичний, з усіма днями місяця. */
export type EjoosTimesheetPersonScan = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  hasDepartureText: boolean;
  firstDepartureDay: number;
  plusDays: number[];
  /** 1-based day → нормалізована позначка (або «вибув»). */
  dayCodes: string[];
  departureText?: string;
};

export type EjoosShpoRow = {
  excelRow: number;
  positionIndex: string;
  personId: string;
  fullName: string;
  rank: string;
};

export type EjoosArrivalRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  arriveDate: string;
  fromUnit: string;
};

export type EjoosOosRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  serviceType: string;
  contractFrom: string;
  contractTo: string;
};

export type EjoosExcludedRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  orderNumber: string;
  orderDate: string;
  destination: string;
  note: string;
};

export type JournalTimesheetDay = {
  day: number;
  label: string;
  sourceDateUnknown?: boolean;
};
