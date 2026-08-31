/** Editable ЕЖООС operator rules: status map + field authority sources. */

export type EjoosTimesheetCode =
  | "+"
  | "вдр"
  | "від"
  | "ВП"
  | "лік"
  | "ЛП"
  | "СЗЧ"
  | "ЗБ"
  | "пол"
  | "заг"
  | "пом";

export const EJOOS_TIMESHEET_CODES: EjoosTimesheetCode[] = [
  "+",
  "вдр",
  "від",
  "ВП",
  "лік",
  "ЛП",
  "СЗЧ",
  "ЗБ",
  "пол",
  "заг",
  "пом",
];

export type EjoosStatusConfidence = "high" | "review" | "manual";

export type EjoosStatusRule = {
  id: string;
  enabled: boolean;
  priority: number;
  label: string;
  /** Any of these substrings (normalized) */
  matchAny: string[];
  /** All of these must also match */
  matchAll?: string[];
  /** Skip rule if any of these match */
  excludeAny?: string[];
  timesheetCode: EjoosTimesheetCode | null;
  absenceGround: string | null;
  confidence: EjoosStatusConfidence;
  reason: string;
};

export type EjoosFieldSource =
  | "pb_sh"
  | "ejoos"
  | "personnel_card"
  | "anketa";

export type EjoosFieldAuthority = {
  field:
    | "status"
    | "fullName"
    | "rank"
    | "rnokpp"
    | "positionIndex"
    | "positionTitle"
    | "phones"
    | "callSign";
  label: string;
  primary: EjoosFieldSource;
  fallbacks: EjoosFieldSource[];
  note: string;
};

export type EjoosOperatorSettings = {
  version: 1;
  unitLabel: string;
  updatedAt: string;
  statusRules: EjoosStatusRule[];
  fieldAuthorities: EjoosFieldAuthority[];
};

export const FIELD_SOURCE_LABELS: Record<EjoosFieldSource, string> = {
  pb_sh: "1ПБ (sh)",
  ejoos: "ЕЖООС",
  personnel_card: "Картка особового складу",
  anketa: "Анкета",
};

export const DEFAULT_STATUS_RULES: EjoosStatusRule[] = [
  {
    id: "on_duty",
    enabled: true,
    priority: 10,
    label: "В строю",
    matchAny: ["в строю", "присут", "+"],
    excludeAny: [
      "сзч",
      "самовіл",
      "відпуст",
      "лік",
      "безвіст",
      "полон",
      "загиб",
      "не в строю",
      "не присут",
      "неприсут",
    ],
    timesheetCode: "+",
    absenceGround: null,
    confidence: "high",
    reason: "СТАТУС вказує на присутність → код табеля «+»",
  },
  {
    id: "own_unit",
    enabled: true,
    priority: 11,
    label: "У підрозділі 1ПБ",
    matchAny: ["_5 1пб", "_4 1пб", "_3 1пб", "_2 1пб", "_1 1пб"],
    timesheetCode: "+",
    absenceGround: null,
    confidence: "high",
    reason: "СТАТУС — підрозділ 1ПБ (`_5 1ПБ`) → код табеля «+»",
  },
  {
    id: "tdy",
    enabled: true,
    priority: 20,
    label: "Відрядження",
    matchAny: ["відряд", "вдр"],
    timesheetCode: "вдр",
    absenceGround: "відрядження",
    confidence: "high",
    reason: "СТАТУС містить відрядження → код «вдр» + період у «Тимч. відсутні»",
  },
  {
    id: "leave_wounded",
    enabled: true,
    priority: 30,
    label: "Відпустка після поранення",
    matchAny: ["відпуст"],
    matchAll: ["поран"],
    timesheetCode: "ВП",
    absenceGround: "відпустка для лікування після поранення",
    confidence: "high",
    reason: "Відпустка після поранення → код «ВП»",
  },
  {
    id: "leave_trauma",
    enabled: true,
    priority: 31,
    label: "Відпустка після травми",
    matchAny: ["відпуст"],
    matchAll: ["травм"],
    timesheetCode: "ВП",
    absenceGround: "відпустка для лікування після поранення",
    confidence: "high",
    reason: "Відпустка після травми → код «ВП»",
  },
  {
    id: "leave_contusion",
    enabled: true,
    priority: 32,
    label: "Відпустка після контузії",
    matchAny: ["відпуст"],
    matchAll: ["контуз"],
    timesheetCode: "ВП",
    absenceGround: "відпустка для лікування після поранення",
    confidence: "high",
    reason: "Відпустка після контузії → код «ВП»",
  },
  {
    id: "leave",
    enabled: true,
    priority: 40,
    label: "Відпустка",
    matchAny: ["відпуст", "від"],
    timesheetCode: "від",
    absenceGround: "відпустка",
    confidence: "high",
    reason: "СТАТУС містить відпустку → код «від»",
  },
  {
    id: "med_wounded",
    enabled: true,
    priority: 50,
    label: "Лікування після поранення",
    matchAny: ["лік", "шпит", "мед"],
    matchAll: ["поран"],
    timesheetCode: "ЛП",
    absenceGround: "лікування після поранення",
    confidence: "high",
    reason: "Лікування після поранення → код «ЛП»",
  },
  {
    id: "med_trauma",
    enabled: true,
    priority: 51,
    label: "Лікування після травми",
    matchAny: ["лік", "шпит", "мед"],
    matchAll: ["травм"],
    timesheetCode: "ЛП",
    absenceGround: "лікування після поранення",
    confidence: "high",
    reason: "Лікування після травми → код «ЛП»",
  },
  {
    id: "med",
    enabled: true,
    priority: 60,
    label: "Лікування",
    matchAny: ["лік", "шпит", "хвороб", "медич", "медрот"],
    timesheetCode: "лік",
    absenceGround: "лікування",
    confidence: "high",
    reason: "СТАТУС містить лікування/шпиталь → код «лік»",
  },
  {
    id: "awol",
    enabled: true,
    priority: 70,
    label: "СЗЧ",
    matchAny: ["сзч", "самовіль"],
    timesheetCode: "СЗЧ",
    absenceGround: "СЗЧ",
    confidence: "high",
    reason: "СТАТУС СЗЧ → код «СЗЧ»",
  },
  {
    id: "mia",
    enabled: true,
    priority: 80,
    label: "Зниклий безвісти",
    matchAny: ["безвіст", "зник", " зб"],
    timesheetCode: "ЗБ",
    absenceGround: "зниклий безвісти",
    confidence: "review",
    reason: "Ймовірно ЗБ — перевірте формулювання статусу",
  },
  {
    id: "pow",
    enabled: true,
    priority: 90,
    label: "Полон",
    matchAny: ["полон", "пол "],
    timesheetCode: "пол",
    absenceGround: "полон",
    confidence: "review",
    reason: "Ймовірно полон — потрібна перевірка",
  },
  {
    id: "kia",
    enabled: true,
    priority: 100,
    label: "Загибель",
    matchAny: ["загиб", "заг "],
    timesheetCode: "заг",
    absenceGround: null,
    confidence: "manual",
    reason: "Загибель лише в дату події + окремий контур документів",
  },
  {
    id: "deceased",
    enabled: true,
    priority: 110,
    label: "Смерть",
    matchAny: ["помер", "смерт"],
    timesheetCode: "пом",
    absenceGround: null,
    confidence: "manual",
    reason: "Смерть лише в дату події + окремий контур документів",
  },
  {
    id: "absent_archive",
    enabled: true,
    priority: 120,
    label: "Відсутній в архіві",
    matchAny: ["відсут"],
    matchAll: ["архів"],
    timesheetCode: null,
    absenceGround: null,
    confidence: "review",
    reason:
      "Статус «ВІДСУТНІЙ в АРХІВІ» — звірити з вкладкою archive / Тимч. відсутні",
  },
];

export const DEFAULT_FIELD_AUTHORITIES: EjoosFieldAuthority[] = [
  {
    field: "status",
    label: "Статус / код табеля",
    primary: "pb_sh",
    fallbacks: ["ejoos"],
    note: "Оперативний статус завжди з 1ПБ; ЕЖООС отримує код після підтвердження.",
  },
  {
    field: "fullName",
    label: "ПІБ",
    primary: "personnel_card",
    fallbacks: ["pb_sh", "ejoos", "anketa"],
    note: "Канонічне ПІБ з картки ОС; 1ПБ/ЕЖООС — для звірки.",
  },
  {
    field: "rank",
    label: "Звання",
    primary: "pb_sh",
    fallbacks: ["ejoos", "personnel_card"],
    note: "Звання з 1ПБ для зайнятості ШПО/Табель.",
  },
  {
    field: "rnokpp",
    label: "РНОКПП",
    primary: "anketa",
    fallbacks: ["personnel_card"],
    note: "РНОКПП не береться з 1ПБ/ЕЖООС — лише з анкети/картки.",
  },
  {
    field: "positionIndex",
    label: "Індекс посади",
    primary: "pb_sh",
    fallbacks: ["ejoos"],
    note: "Індекс посади з sh для оновлення ШПО.",
  },
  {
    field: "positionTitle",
    label: "Посада (текст)",
    primary: "personnel_card",
    fallbacks: ["pb_sh", "ejoos"],
    note: "Повна назва посади з картки ОС.",
  },
  {
    field: "phones",
    label: "Телефони",
    primary: "personnel_card",
    fallbacks: ["anketa"],
    note: "Контакти не з 1ПБ.",
  },
  {
    field: "callSign",
    label: "Позивний",
    primary: "personnel_card",
    fallbacks: ["anketa", "pb_sh"],
    note: "Позивний з картки; 1ПБ — резерв.",
  },
];

export const buildDefaultOperatorSettings = (
  unitLabel = "1ПБ",
): EjoosOperatorSettings => ({
  version: 1,
  unitLabel,
  updatedAt: new Date().toISOString(),
  statusRules: DEFAULT_STATUS_RULES.map((rule) => ({ ...rule })),
  fieldAuthorities: DEFAULT_FIELD_AUTHORITIES.map((item) => ({
    ...item,
    fallbacks: [...item.fallbacks],
  })),
});
