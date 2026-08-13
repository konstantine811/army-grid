import type { CellValue } from "../../excelRoundTrip";
import type { AnalyticsMetric } from "../analytics/analyticsData";

export type BchsAnalyticsRow = {
  rowNumber: number;
  unit: string;
  staff: number;
  staffOfficers: number;
  staffSergeants: number;
  staffSoldiers: number;
  listed: number;
  listedOfficers: number;
  listedSergeants: number;
  listedSoldiers: number;
  staffedPercent: CellValue;
  available: number;
  availableOfficers: number;
  availableSergeants: number;
  availableSoldiers: number;
  shortage: number;
  shortagePercent: CellValue;
  absent: number;
  businessTrip: number;
  training: number;
  hospitalWounded: number;
  hospitalIllness: number;
  vacation: number;
  awol: number;
  missing: number;
  killed: number;
  medWounded: number;
  medIllness: number;
  inRanksActually: number;
  actualPercent: CellValue;
  combatComponent: number;
};

export type BchsComparisonRow = BchsAnalyticsRow & {
  actualOfficers: number;
  actualSergeants: number;
  actualSoldiers: number;
  awayInOtherUnits: number;
  awayOfficers: number;
  awaySergeants: number;
  awaySoldiers: number;
  awayDestinationsText: string;
  attachedFromOtherUnits: number;
  attachedOfficers: number;
  attachedSergeants: number;
  attachedSoldiers: number;
  attachedSourcesText: string;
  unassignedNewcomers: number;
  noBzvp: number;
  levelPercent: number;
  balanceActual: number;
  assaultReady: number;
  assaultRecovery: number;
  assaultExecution: number;
  assaultTotal: number;
  droneCrew: number;
  vehicleCrew: number;
  crewServedWeapons: number;
  commandCombat: number;
  supportCombat: number;
};

export type BchsAnalyticsTableColumn = {
  key: string;
  letter: string;
  label: string;
  isPercent?: boolean;
};

export type BchsAnalyticsTableRow = {
  rowNumber: number;
  values: Record<string, string | number>;
};

export type BchsSupplementRow = {
  rowNumber: number;
  battalion: string;
  unit: string;
  staff: number;
  listed: number;
  available: number;
  staffedPercent: number;
  combatTask: number;
  replacementReserve: number;
  taskReserve: number;
  commanderReserve: number;
  absent: number;
  businessTrip: number;
  training: number;
  hospitalWounded: number;
  hospitalIllness: number;
  vacation: number;
  awol: number;
  missing: number;
  killed: number;
  medWounded: number;
  medIllness: number;
  detached: number;
  attached: number;
  newcomers: number;
  inRanks: number;
  assaultReady: number;
  assaultRecovery: number;
  assaultExecution: number;
  noBzvp: number;
  assaultTotal: number;
  vehicleCrew: number;
  droneCrew: number;
  crewServedWeapons: number;
  commandCombat: number;
  supportCombat: number;
  bzvpBuckets: AnalyticsMetric[];
  totalBzvp: number;
};

export type BchsSupplementSnapshot = {
  kind: "personnel-bzvp" | "appendix";
  title: string;
  reportDate: string;
  total: BchsSupplementRow;
  rows: BchsSupplementRow[];
  totals: BchsSupplementRow[];
  absenceReasons: AnalyticsMetric[];
  combatCategories: AnalyticsMetric[];
  reserveMetrics: AnalyticsMetric[];
  bzvpBuckets: AnalyticsMetric[];
};

export type BchsAnalyticsSnapshot = {
  reportDate: string;
  total: BchsComparisonRow;
  rows: BchsComparisonRow[];
  comparisonRows: BchsComparisonRow[];
  table?: {
    columns: BchsAnalyticsTableColumn[];
    rows: BchsAnalyticsTableRow[];
  };
  detachedDestinations: AnalyticsMetric[];
  attachedSources: AnalyticsMetric[];
  absenceReasons: AnalyticsMetric[];
  dataIssues?: BchsDataIssue[];
  supplement?: BchsSupplementSnapshot;
};

export type BchsDataIssue = {
  fullName: string;
  rosterUnit: string;
  status: string;
  destination?: string;
  reason: string;
};

export type BchsPersonnelAwayPerson = {
  battalion: string;
  rosterUnit: string;
  rankCategory: string;
  rankTitle: string;
  fullName: string;
  status: string;
  roleType: string;
  combatReadiness: string;
  bzvpStatus: string;
  destination: string;
  medicalPlace: string;
  medicalNote: string;
};

export type BchsUnitAwayStats = {
  officers: number;
  sergeants: number;
  soldiers: number;
  total: number;
  destinations: Map<string, number>;
  destinationText: string;
};

export type BchsUnitAttachedStats = {
  officers: number;
  sergeants: number;
  soldiers: number;
  total: number;
  sources: Map<string, number>;
  sourcesText: string;
};

export type GeneratedWorkbook = any;

export type GeneratedSheet = any;
