import { formatTimesheetDeparture } from "./ejoosTimesheetText";

/**
 * Карта колонок «2. ООС» → «3. Виключені» для історичної частини A:AA.
 *
 * «Виключені» не формується з Рух з нуля: готову картку особи переносимо з
 * ООС, а подію виключення дописуємо окремо в AB:AF.
 */
export const OOS_TO_EXCLUDED_BASE: Array<[number, number]> = [
  [1, 1], // A→A звання
  [2, 2], // B→B ПІБ
  [3, 3], // C→C ID
  [4, 4], // D→D індекс / історія посад
  [5, 5], // E→E дати прийняття посад
  [7, 6], // G→F звідки прибув
  [8, 7], // H→G дата зарахування
  [9, 8], // I→H наказ про зарахування: дата
  [10, 9], // J→I наказ про зарахування: номер
  [11, 10], // K→J наказ на посаду: номер
  [12, 11], // L→K наказ на посаду: дата
  [15, 12], // O→L наказ на останнє звання: дата
  [16, 13], // P→M наказ на останнє звання: номер
  [19, 14], // S→N вид служби
  [20, 15], // T→O початок контракту
  [21, 16], // U→P закінчення контракту
  [22, 17], // V→Q РНОКПП
  [24, 18], // X→R документ: номер
  [25, 19], // Y→S документ: тип
  [26, 20], // Z→T дата народження
  [27, 21], // AA→U місце народження
  [28, 22], // AB→V стать
  [30, 23], // AD→W ким призваний
  [29, 24], // AC→X коли призваний
  [31, 25], // AE→Y освіта
  [32, 26], // AF→Z родичі
  [33, 27], // AG→AA додаткова інформація
];

export const OOS_TO_EXCLUDED_ALL: Array<[number, number]> = [
  ...OOS_TO_EXCLUDED_BASE,
];

/** Зворотне перенесення картки з «Виключені» в ООС після скасованого вибуття. */
export const EXCLUDED_TO_OOS_BASE: Array<[number, number]> = OOS_TO_EXCLUDED_BASE.map(
  ([oosCol, exclCol]) => [exclCol, oosCol],
);

/** Колонки Виключені з VLOOKUP у шаблоні (1-based). AB–AF не чіпаємо. */
export const EXCLUDED_VLOOKUP_COLUMNS_1BASED = [
  ...new Set(OOS_TO_EXCLUDED_BASE.map(([, excl]) => excl)),
];

/**
 * Фіксовані колонки «3. Виключені» ← анкета (за змістом заголовків аркуша).
 * N–P у шаблоні без VLOOKUP — їх теж дописуємо з анкети.
 */
export const EXCLUDED_ANKETA_FIXED_COLUMNS: Array<{
  column: number;
  anketaKey:
    | "rank"
    | "externalId"
    | "positionIndex"
    | "positionDates"
    | "positionOrderNumber"
    | "arrivedFrom"
    | "enlistDate"
    | "enlistOrderDate"
    | "enlistOrderNumber"
    | "appointmentOrderNumber"
    | "appointmentOrderDate"
    | "serviceType"
    | "contractFrom"
    | "contractTo"
    | "rnokpp"
    | "rnokppRefuse"
    | "idDocumentNumber"
    | "idDocumentName"
    | "birthDate"
    | "birthPlace"
    | "sex"
    | "conscriptedWhen"
    | "conscriptedBy"
    | "education"
    | "relatives"
    | "additionalInfo";
}> = [
  { column: 1, anketaKey: "rank" },
  { column: 3, anketaKey: "externalId" },
  { column: 4, anketaKey: "positionIndex" },
  { column: 5, anketaKey: "positionDates" },
  { column: 6, anketaKey: "arrivedFrom" },
  { column: 7, anketaKey: "enlistDate" },
  { column: 8, anketaKey: "enlistOrderDate" },
  { column: 9, anketaKey: "enlistOrderNumber" },
  { column: 10, anketaKey: "appointmentOrderNumber" },
  { column: 11, anketaKey: "appointmentOrderDate" },
  { column: 14, anketaKey: "serviceType" },
  { column: 15, anketaKey: "contractFrom" },
  { column: 16, anketaKey: "contractTo" },
  { column: 17, anketaKey: "rnokpp" },
  { column: 18, anketaKey: "idDocumentNumber" },
  { column: 19, anketaKey: "idDocumentName" },
  { column: 20, anketaKey: "birthDate" },
  { column: 21, anketaKey: "birthPlace" },
  { column: 22, anketaKey: "sex" },
  { column: 23, anketaKey: "conscriptedBy" },
  { column: 24, anketaKey: "conscriptedWhen" },
  { column: 25, anketaKey: "education" },
  { column: 26, anketaKey: "relatives" },
  { column: 27, anketaKey: "additionalInfo" },
];

/** Типовий перший рядок даних на Виключені / ООС (після заголовків і нумерації). */
export const EJOOS_PERSON_DATA_START_ROW = 6;

/** Z «Дані про родичів» і AF «Підстава виключення» — багаторядковий текст. */
export const isExcludedWrapColumn = (column: number) =>
  column === 26 || column === 32;

type TransferUnitPayload = {
  type?: string;
  exclusionReason?: string;
  destination?: string;
  documentsDest?: string;
  timesheetDestination?: string;
  changeText?: string;
};

const TRANSFER_FIELDS = (payload: TransferUnitPayload) => [
  payload.destination,
  payload.changeText,
  payload.timesheetDestination,
  payload.documentsDest,
  payload.exclusionReason,
];

const extractTransferUnitLine = (raw: string) => {
  const text = String(raw || "").replace(/\s+/g, " ").trim();
  if (!text) return "";

  const companyBn = text.match(
    /_+\s*(\d+)\s+(\d+)\s*(ПБ|ШБ|РБ|ТБ)\b/iu,
  );
  if (companyBn) {
    return `_ ${companyBn[1]} ${companyBn[2]}${companyBn[3].toUpperCase()}`;
  }
  const underscored = text.match(/_+\s*(\d+)\s*(ПБ|ШБ|РБ|ТБ)\b/iu);
  if (underscored) {
    return `_ ${underscored[1]} ${underscored[2].toUpperCase()}`;
  }
  const companyBnBare = text.match(/\b(\d+)\s+(\d+)\s*(ПБ|ШБ|РБ|ТБ)\b/iu);
  if (companyBnBare) {
    return `_ ${companyBnBare[1]} ${companyBnBare[2]}${companyBnBare[3].toUpperCase()}`;
  }
  const storm = text.match(
    /(\d+)\s*штурмов(?:ий|ого|ому|им)?\s+батальйон/iu,
  );
  if (storm) return `_ ${storm[1]} ШБ`;
  const infantry = text.match(
    /(\d+)\s*піхотн(?:ий|ого|ому|им)?\s+батальйон/iu,
  );
  if (infantry) return `_ ${infantry[1]} ПБ`;
  const compact = text.match(/\b(\d+)[_ ](БП|ПБ|ШБ|РБ|ТБ)\b/iu);
  if (compact) return `_ ${compact[1]}_${compact[2].toUpperCase()}`;
  const abbrev = text.match(/\b(\d+)\s*(ПБ|ШБ|РБ|ТБ)\b/iu);
  if (abbrev) return `_ ${abbrev[1]} ${abbrev[2].toUpperCase()}`;
  if (/бат(?:альйон)?(?:у|і)?\s+зв['’ʼ]?яз/iu.test(text)) return `_ БАТ ЗВ'ЯЗ`;
  const code = text.match(/[АAаa]\s*(\d{4})(?!\d)/iu);
  if (code) return `_ А${code[1]}`;
  return "";
};

/** Коротке «куди» для колонки «Підстава виключення»: _ 5 2ПБ, _ А4784. */
export const shortExcludedTransferUnit = (payload: TransferUnitPayload) => {
  const patterns: Array<(value: string) => string> = [
    (value) => {
      const hit = value.match(/_+\s*(\d+)\s+(\d+)\s*(ПБ|ШБ|РБ|ТБ)\b/iu);
      return hit ? `_ ${hit[1]} ${hit[2]}${hit[3].toUpperCase()}` : "";
    },
    (value) => {
      const hit = value.match(/_+\s*(\d+)\s*(ПБ|ШБ|РБ|ТБ)\b/iu);
      return hit ? `_ ${hit[1]} ${hit[2].toUpperCase()}` : "";
    },
    extractTransferUnitLine,
  ];
  for (const match of patterns) {
    for (const field of TRANSFER_FIELDS(payload)) {
      const unit = match(String(field || ""));
      if (unit) return unit;
    }
  }
  return "";
};

/**
 * Підстава виключення (колонка AF): коротко, як у шаблоні.
 * ПЕРЕВЕДЕННЯ
 * _ А4784
 */
export const formatExcludedListBasis = (payload: TransferUnitPayload) => {
  const type = String(payload.type || "");
  const raw = String(payload.exclusionReason || "").replace(/\s+/g, " ").trim();
  const isDisposition = /розпорядж/i.test(type) || /розпорядж/i.test(raw);
  const title = isDisposition ? "РОЗПОРЯДЖЕННЯ" : "ПЕРЕВЕДЕННЯ";
  const unit = shortExcludedTransferUnit(payload);
  return unit ? `${title}\n${unit}` : title;
};

/** Колонка AE «Куди вибув» — lowercase як у шаблоні. */
export const formatExcludedDestinationText = (value: string) =>
  value
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

/**
 * День вибуття в Табелі: «вибув до/у» + фраза з «Яка зміна» без назви посади.
 * Не плутати з короткою підставою у Виключених (ПЕРЕВЕДЕННЯ / _ А4784).
 */
export const formatTimesheetTransferMark = (payload: TransferUnitPayload) => {
  const dest = [
    payload.timesheetDestination,
    payload.documentsDest,
    payload.destination,
    payload.changeText,
  ]
    .map((value) => String(value || "").replace(/\s+/g, " ").trim())
    .find(Boolean) || "";
  return formatTimesheetDeparture(dest);
};
