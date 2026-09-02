export type RankGroup = "officer" | "sergeant" | "soldier";
export type ServiceType = "mobilized" | "contract";

export type RankServiceBucket =
  | "officerMobilized"
  | "officerContract"
  | "sergeantMobilized"
  | "sergeantContract"
  | "soldierMobilized"
  | "soldierContract";

export type Sex = "male" | "female" | "unknown";

export type MaritalStatus = "married" | "unmarried" | "civil" | "unknown";

export type ArrivalSource =
  | "tck"
  | "trainingCenter"
  | "recruiting"
  | "brez"
  | "unitTransfer"
  | "other"
  | "unknown";

export type NationalityKey =
  | "ukraine"
  | "poland"
  | "usa"
  | "britain"
  | "portugal"
  | "argentina"
  | "russia"
  | "other";

export type RegionKey =
  | "kyivCity"
  | "kyiv"
  | "vinnytsia"
  | "volyn"
  | "dnipro"
  | "donetskControlled"
  | "donetskOccupied"
  | "zhytomyr"
  | "zakarpattia"
  | "ivanoFrankivsk"
  | "kirovohrad"
  | "luhansk"
  | "lviv"
  | "mykolaiv"
  | "odesa"
  | "poltava"
  | "rivne"
  | "sumy"
  | "ternopil"
  | "zaporizhzhiaControlled"
  | "zaporizhzhiaOccupied"
  | "kharkivControlled"
  | "kharkivOccupied"
  | "khersonControlled"
  | "khersonOccupied"
  | "khmelnytskyi"
  | "cherkasy"
  | "chernivtsi"
  | "chernihiv"
  | "crimea"
  | "unknown";

export type AgeBand = "18-25" | "26-30" | "31-40" | "41-50" | "50+" | "unknown";

export type ExitBand =
  | "none"
  | "1-4"
  | "5-10"
  | "11-15"
  | "16-20"
  | "21-25"
  | "26-30"
  | "30+";

export type PassportMetricId =
  | "staff"
  | "listed"
  | "present"
  | "disposition"
  | "arrived"
  | "fromTck"
  | "fromTrainingCenter"
  | "fromRecruiting"
  | "fromBrez"
  | "male"
  | "female"
  | "age18_25"
  | "age26_30"
  | "age31_40"
  | "age41_50"
  | "age50plus"
  | "natUkraine"
  | "natPoland"
  | "natUsa"
  | "natBritain"
  | "natPortugal"
  | "natArgentina"
  | "natOther"
  | "natRussia"
  | "exitsNone"
  | "exits1_4"
  | "exits5_10"
  | "exits11_15"
  | "exits16_20"
  | "exits21_25"
  | "exits26_30"
  | "exits30plus"
  | RegionKey
  | "married"
  | "unmarried"
  | "civil"
  | "childrenUnder18"
  | "children3plus"
  | "relativesServing"
  | "relativesAbroad"
  | "relativesHostile"
  | "ubd"
  | "idp";

export type BucketCounts = Record<RankServiceBucket, number>;

export type PassportRowKind = "section" | "metric";

export type PassportTableRow = {
  id: string;
  kind: PassportRowKind;
  number: string;
  label: string;
  metricId?: PassportMetricId;
  counts?: BucketCounts;
  total?: number;
};

export type ParsedRelatives = {
  marital: MaritalStatus;
  childCount: number;
  childrenUnder18: number;
  relativesServing: boolean;
  relativesAbroad: boolean;
  relativesHostile: boolean;
  notes: string[];
};

export type SocPerson = {
  id: string;
  name: string;
  normalizedName: string;
  shortName: string;
  callsign: string;
  position: string;
  positionIndex: string;
  rank: string;
  staffRank: string;
  rankGroup: RankGroup;
  serviceType: ServiceType;
  sex: Sex;
  birthDate: string;
  age: number | null;
  ageBand: AgeBand;
  birthPlace: string;
  region: RegionKey;
  regionLabel: string;
  regionOccupied: boolean;
  nationality: NationalityKey;
  marital: MaritalStatus;
  childrenUnder18: number;
  children3plus: boolean;
  relativesServing: boolean;
  relativesAbroad: boolean;
  relativesHostile: boolean;
  relativesRaw: string;
  extraRaw: string;
  arrivedFrom: string;
  calledBy: string;
  arrivalSource: ArrivalSource;
  hasUbd: boolean;
  ubdNumber: string;
  ubdRosterStatus: "submitted" | "notSubmitted" | null;
  /** Вручну виключені з «Не виконували» → мін. 1 бойовий вихід у статистиці. */
  staticCombatExitOverride: boolean;
  oosDislocation: string;
  combatDutyEvidence: string[];
  isIdp: boolean;
  morningStatus: string;
  morningAbsenceNotes: string;
  morningDestination: string;
  morningLocation: string;
  isTransiter: boolean;
  bzvpStatus: string;
  brezAssignment: string;
  onStaff: boolean;
  onList: boolean;
  present: boolean;
  inDisposition: boolean;
  morningExitCount: number;
  jbdExitCount: number;
  exitCount: number;
  exitBand: ExitBand;
  match: {
    oos: boolean;
    morning: boolean;
    exits: boolean;
    jbdExits: boolean;
    bplaExits: boolean;
    ubdRoster: boolean;
    tempArrival: boolean;
  };
  parseNotes: string[];
};

export type SocStaffSlot = {
  position: string;
  positionIndex: string;
  staffRank: string;
  rankGroup: RankGroup;
  occupied: boolean;
  name: string;
};

export type SocPassportSummary = {
  staffSlots: number;
  occupied: number;
  vacant: number;
  oosMatched: number;
  morningMatched: number;
  exitsMatched: number;
  combatDutyMatched: number;
  relativesParsed: number;
  unknownRegion: number;
  unknownAge: number;
};

export type SocPassportResult = {
  people: SocPerson[];
  staffSlots: SocStaffSlot[];
  rows: PassportTableRow[];
  summary: SocPassportSummary;
  warnings: string[];
  sheets: {
    shpo: string;
    oos: string;
    morning?: string;
    fighterStatus?: string;
    tempArrived?: string;
    jbdExits?: string;
    bplaExits?: string;
    ubdRoster?: string;
    housingIdp?: string;
  };
};
