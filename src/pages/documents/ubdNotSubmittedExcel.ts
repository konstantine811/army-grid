import writeXlsxFile, {
  type CellObject,
  type SheetData,
} from "write-excel-file/browser";
import dayjs from "dayjs";
import type { BackendPersonDocument, BackendPersonnelProfile } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import {
  extractPhones,
  formatExcelDateDisplay,
  formatUaPhoneDisplay,
  getPersonDisplayName,
  getPersonExternalId,
  getPersonFieldValue,
  previewValueToDisplay,
  resolvePersonCallSign,
} from "../personnel/personnelUtils";
import { getRosterValue } from "../personnel/personnelRosterMerge";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import { splitFullNameParts } from "./serviceCharacteristicReport";
import type { AnketaRow } from "../anketa-data/anketaSheet";

/**
 * Колонки як у реєстрі УБД («Подавалися» / «Не подавалися»),
 * без зведеної колонки «Прізвище Власне ім’я По батькові».
 */
export const UBD_NOT_SUBMITTED_HEADERS = [
  "Позивний",
  "Прізвище",
  "Власне ім’я",
  "По батькові (за наявності)",
  "Стать",
  "Дата народження",
  "Серія документу",
  "Номер документу",
  "Реєстраційний номер облікової картки платника податків",
  "Номер контактного телефону",
  "Військове (спеціальне) звання (за наявності)",
  "Н.П. бойових дій",
  "Період участі у бойових діях",
  "посада",
  "БР",
  "Статус",
] as const;

export const UBD_STATUS_HEADERS = [
  "Позивний",
  "Прізвище",
  "Власне ім’я",
  "По батькові (за наявності)",
  "Статус",
  "Статус УБД",
  "Документи",
  "Рапорт",
  "Відправлено",
] as const;

export type UbdNotSubmittedExcelRow = {
  callSign: string;
  lastName: string;
  firstName: string;
  patronymic: string;
  sex: string;
  birthDate: string;
  idSeries: string;
  idNumber: string;
  rnokpp: string;
  phone: string;
  rank: string;
  combatPlace: string;
  combatPeriod: string;
  position: string;
  combatOrder: string;
  personnelStatus: string;
  ubdStatus: string;
  documents: string;
  raport: string;
  sent: string;
};

const HEADER_GREEN = "#1F4E3D";
const BORDER = "#A9BDB4";
const ZEBRA = "#F4F8F6";
const WHITE = "#FFFFFF";
const TEXT = "#1A1A14";

const border = {
  borderColor: BORDER,
  borderStyle: "thin" as const,
};

const cell = (
  value: string | number | Date | null,
  extra: Omit<CellObject, "value"> = {},
): CellObject => ({
  value: value ?? undefined,
  fontFamily: "Calibri",
  fontSize: 11,
  alignVertical: "center",
  ...border,
  ...extra,
});

const text = (value: string) => String(value ?? "").trim();

/**
 * Профіль БД часто віддає `{ id, values: { ...поля } }` — як у картці особи.
 * Без розгортання `values` позивний / серія документа не читаються.
 */
const asRow = (value: unknown): EjournalPreviewRow | null => {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const nested = record.values;
  if (nested && typeof nested === "object" && !Array.isArray(nested)) {
    const id = String(record.id ?? record.__dbRowId ?? "").trim();
    return {
      ...(id ? { __dbRowId: id } : {}),
      ...(nested as EjournalPreviewRow),
    };
  }
  return record as EjournalPreviewRow;
};

/** Зливає анкету / ООС / штатку, щоб позивний і серія документа не губились. */
const mergePersonRows = (
  ...rows: Array<EjournalPreviewRow | null>
): EjournalPreviewRow | null => {
  const merged: EjournalPreviewRow = {};
  let hasAny = false;
  for (const row of rows) {
    if (!row) continue;
    for (const [key, value] of Object.entries(row)) {
      if (key === "values" || value == null) continue;
      const asText = typeof value === "string" ? value.trim() : value;
      if (asText === "") continue;
      const current = merged[key];
      if (current == null || current === "") {
        merged[key] = value;
        hasAny = true;
      }
    }
  }
  return hasAny ? merged : null;
};

const pickPersonRow = (
  profile: BackendPersonnelProfile | null | undefined,
  rosterRow?: EjournalPreviewRow | null,
): EjournalPreviewRow | null =>
  mergePersonRows(
    asRow(profile?.person),
    asRow(profile?.ejournal?.oosRow),
    asRow(profile?.roster?.row),
    rosterRow ?? null,
  );

export const formatUbdRegistrySex = (value: string) => {
  const normalized = text(value).toLocaleLowerCase("uk-UA");
  if (!normalized) return "Ч";
  if (
    normalized === "ч" ||
    normalized.startsWith("чол") ||
    normalized === "male" ||
    normalized === "m"
  ) {
    return "Ч";
  }
  if (
    normalized === "ж" ||
    normalized.startsWith("жін") ||
    normalized === "female" ||
    normalized === "f"
  ) {
    return "Ж";
  }
  return "Ч";
};

/** Серія/номер з картки: «КС 691264», «КС691264» або лише номер «12165720». */
export const splitUbdIdDocument = (value: string) => {
  const raw = text(value).replace(/\s+/g, " ");
  if (!raw) return { series: "", number: "" };
  const match = raw.match(/^([A-Za-zА-ЯІЇЄҐа-яіїєґ]{2})\s*([0-9].*)$/u);
  if (match) {
    return {
      series: match[1].toLocaleUpperCase("uk-UA"),
      number: match[2].replace(/\s+/g, ""),
    };
  }
  return { series: "", number: raw.replace(/\s+/g, "") };
};

/** Префікс колонки БР у Excel УБД. */
export const UBD_COMBAT_ORDER_PREFIX = "425 ОШП «СКЕЛЯ» №";

/** @deprecated використайте UBD_COMBAT_ORDER_PREFIX */
export const UBD_COMBAT_ORDER_STUB = UBD_COMBAT_ORDER_PREFIX;

const stripBasisNumber = (value: string) =>
  String(value ?? "")
    .trim()
    .replace(/^№\s*/, "");

const parseBasisFromCombined = (value: string) => {
  const textValue = String(value ?? "").trim();
  const match = textValue.match(
    /№\s*(\S+)\s+від\s+(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/i,
  );
  if (!match) {
    const numberOnly = textValue.match(/№\s*(\S+)/i);
    return {
      number: numberOnly ? stripBasisNumber(numberOnly[1]) : "",
      date: "",
    };
  }
  return {
    number: stripBasisNumber(match[1]),
    date: match[2].replaceAll("/", ".").replaceAll("-", "."),
  };
};

/** БР у Excel: `425 ОШП «СКЕЛЯ» №4862/…/дск` (+ дата, якщо є). */
export const formatUbdCombatOrder = (
  basisNumber?: string,
  basisDate?: string,
  basisCombined?: string,
) => {
  let number = stripBasisNumber(basisNumber ?? "");
  let date = String(basisDate ?? "").trim();
  if (!number) {
    const parsed = parseBasisFromCombined(basisCombined ?? "");
    number = parsed.number;
    if (!date) date = parsed.date;
  }
  if (!number) return UBD_COMBAT_ORDER_PREFIX;
  const withNumber = `${UBD_COMBAT_ORDER_PREFIX}${number}`;
  return date ? `${withNumber} від ${date}` : withNumber;
};

const formatCombatPeriod = (value: string) => {
  const raw = text(value).replace(/^з\s+/i, "").trim();
  if (!raw) return "";
  const match = raw.match(
    /(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})\s*[-–—]\s*(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/,
  );
  if (!match) return raw;
  const start = match[1].replaceAll("/", ".").replaceAll("-", ".");
  const end = match[2].replaceAll("/", ".").replaceAll("-", ".");
  return `${start} - ${end}`;
};

const readField = (
  fields: Record<string, unknown> | null | undefined,
  key: string,
) => {
  const value = fields?.[key];
  return typeof value === "string" ? value.trim() : "";
};

const yesIf = (condition: boolean) => (condition ? "так" : "");

const resolveUbdWorkflowFlags = (document: BackendPersonDocument) => {
  const steps = [
    "document",
    "account",
    "scan",
    "ready",
    "sentReport",
    "sentScans",
    "received",
    "handed",
  ];
  const workflow =
    document.workflow && typeof document.workflow === "object"
      ? (document.workflow as Record<string, unknown>)
      : {};
  const completed =
    workflow.completed && typeof workflow.completed === "object"
      ? (workflow.completed as Record<string, unknown>)
      : {};
  let status = String(
    document.status || workflow.currentStatus || "document",
  ).trim();
  if (status === "sent") status = "sentReport";
  const index = Math.max(0, steps.indexOf(status));
  const files =
    document.files && typeof document.files === "object"
      ? (document.files as { ubdScans?: unknown[] })
      : {};
  const hasScans = Array.isArray(files.ubdScans) && files.ubdScans.length > 0;
  const note = readField(document.fields as Record<string, unknown>, "statusNote");
  const sentMarked =
    completed.sent === true ||
    completed.sentReport === true ||
    completed.sentScans === true ||
    index >= steps.indexOf("sentReport");
  return {
    documents: yesIf(
      completed.scan === true ||
        index >= steps.indexOf("scan") ||
        hasScans,
    ),
    raport: yesIf(
      completed.document === true || index >= steps.indexOf("document"),
    ),
    sent: note || yesIf(sentMarked),
  };
};

const readIdDocumentRaw = (personRow: EjournalPreviewRow | null) => {
  const fromKnown =
    getPersonFieldValue(personRow, [
      "серія",
      "номер",
      "документа",
      "посвідчує",
    ]) ||
    getPersonFieldValue(personRow, ["серія", "номер", "документа"]) ||
    getPersonFieldValue(personRow, ["серія", "номер"]) ||
    getPersonFieldValue(personRow, ["серія_номер_документа"]);
  if (fromKnown) return fromKnown;
  if (!personRow) return "";

  // Fallback: будь-яке поле картки на кшталт «Серія/номер документа».
  for (const [key, raw] of Object.entries(personRow)) {
    if (key.startsWith("__")) continue;
    const normalized = key.toLocaleLowerCase("uk-UA");
    if (!normalized.includes("сері") || !normalized.includes("номер")) continue;
    if (normalized.includes("посвідченн") && !normalized.includes("документ")) {
      continue;
    }
    const value = previewValueToDisplay(raw).trim();
    if (value) return value;
  }
  return "";
};

const readCallSign = (
  personRow: EjournalPreviewRow | null,
  profile: BackendPersonnelProfile | null | undefined,
  fields: Record<string, unknown>,
  rosterCallSign?: string,
) => {
  const fromRosterHint = text(rosterCallSign || "");
  if (fromRosterHint) return fromRosterHint;

  const fromRow = resolvePersonCallSign(personRow);
  if (fromRow) return fromRow;

  // Штатка: колонка 15 «Позивний» часто як column_15 без слова «позив» у ключі.
  if (personRow) {
    const fromColumn = text(readRosterColumnValue(personRow, 15));
    if (fromColumn) return fromColumn;
  }

  const fromField = readField(fields, "callSign");
  if (fromField) return fromField;

  const ticket = profile?.documents?.find(
    (item) => item.type === "temporaryMilitaryId",
  );
  return readField(
    ticket?.fields as Record<string, unknown> | null | undefined,
    "callSign",
  );
};

/** Індекс позивних зі Штатки / Загального списку для експорту УБД. */
export const buildUbdRosterCallSignIndex = (rows: EjournalPreviewRow[]) => {
  const byId = new Map<string, { row: EjournalPreviewRow; callSign: string }>();
  const byName = new Map<string, { row: EjournalPreviewRow; callSign: string }>();

  for (const row of rows) {
    const callSign =
      resolvePersonCallSign(row) ||
      text(readRosterColumnValue(row, 15)) ||
      text(getRosterValue(row, ["позив"]));
    if (!callSign) continue;

    const entry = { row, callSign };
    const id = getPersonExternalId(row);
    if (id && !byId.has(id)) byId.set(id, entry);

    const name = normalizeRosterMatchText(
      getPersonDisplayName(row) ||
        getRosterValue(row, ["піб"]) ||
        getRosterValue(row, ["прізвище"]),
    );
    if (name && !byName.has(name)) byName.set(name, entry);
  }

  return { byId, byName };
};

export const lookupUbdRosterCallSign = (
  index: ReturnType<typeof buildUbdRosterCallSignIndex>,
  personExternalId: string,
  fullName: string,
) => {
  const id = text(personExternalId);
  if (id) {
    const byId = index.byId.get(id);
    if (byId) return byId;
  }

  const name = normalizeRosterMatchText(fullName);
  if (!name) return null;
  const exact = index.byName.get(name);
  if (exact) return exact;

  const surname = name.split(" ")[0] || "";
  if (surname.length < 3) return null;
  for (const [rowName, entry] of index.byName) {
    if (
      rowName.startsWith(surname) &&
      name.split(" ").every((part) => !part || rowName.includes(part))
    ) {
      return entry;
    }
  }
  return null;
};

/** Коротка «Посада» (не «Повна посада»). */
const readShortPosition = (
  personRow: EjournalPreviewRow | null,
  fields: Record<string, unknown>,
) => {
  if (personRow) {
    const fromRosterColumn = text(readRosterColumnValue(personRow, 5));
    if (fromRosterColumn) return fromRosterColumn;

    let shortFallback = "";
    for (const [key, raw] of Object.entries(personRow)) {
      if (key.startsWith("__")) continue;
      const value = previewValueToDisplay(raw).trim();
      if (!value) continue;
      const norm = key
        .replace(/^roster__/i, "")
        .toLocaleLowerCase("uk-UA")
        .replace(/[\s-]+/g, "_");
      if (norm.includes("повна") && norm.includes("посада")) continue;
      if (
        (norm === "посада" ||
          norm.endsWith("_посада") ||
          /^column_5(_|$)/i.test(key)) &&
        !norm.includes("індекс") &&
        !norm.includes("прийняття") &&
        !norm.includes("наказу")
      ) {
        return value;
      }
      if (
        !shortFallback &&
        norm.includes("посада") &&
        !norm.includes("індекс") &&
        !norm.includes("прийняття") &&
        !norm.includes("наказу")
      ) {
        shortFallback = value;
      }
    }
    if (shortFallback) return shortFallback;
  }

  const fromDocument = readField(fields, "staffPosition");
  // Якщо в рапорті вже збережена «повна посада» — не тягнемо її в колонку «посада».
  if (
    fromDocument &&
    !/батальйон|рот[аии]|взвод|відділен/i.test(fromDocument)
  ) {
    return fromDocument;
  }
  return "";
};

export const buildUbdNotSubmittedRowFromDocument = ({
  document,
  profile,
  personnelStatus,
  rosterRow,
  rosterCallSign,
  anketaRow,
}: {
  document: BackendPersonDocument;
  profile?: BackendPersonnelProfile | null;
  personnelStatus?: string;
  rosterRow?: EjournalPreviewRow | null;
  rosterCallSign?: string;
  anketaRow?: AnketaRow | null;
}): UbdNotSubmittedExcelRow => {
  const fields = (document.fields || {}) as Record<string, unknown>;
  const personRow = pickPersonRow(profile, rosterRow);
  const anketa = anketaRow ?? null;
  const summaryName =
    readField(fields, "fullName") ||
    text(profile?.fullName || "") ||
    text(anketa?.fullName || "") ||
    getPersonFieldValue(personRow, ["прізвище"]) ||
    getPersonFieldValue(personRow, ["піб"]) ||
    `ID ${document.personExternalId}`;

  // Прізвище = лише перше слово ПІБ; імʼя / по батькові — наступні.
  const { lastName, firstName, patronymic } = splitFullNameParts(summaryName);

  const idFromAnketa = text(anketa?.idDocumentNumber || "");
  const { series: idSeries, number: idNumber } = splitUbdIdDocument(
    readIdDocumentRaw(personRow) || idFromAnketa,
  );

  const additionalInfo =
    getPersonFieldValue(personRow, ["додаткова_інформація"]) ||
    text(anketa?.additionalInfo || "");
  const phones = extractPhones(additionalInfo);
  const phone = phones.map(formatUaPhoneDisplay).join("\n");

  const birthFromFields = readField(fields, "birthDate");
  const birthFromPerson = formatExcelDateDisplay(
    getPersonFieldValue(personRow, ["дата_народження"]),
  );
  const birthFromAnketa = formatExcelDateDisplay(anketa?.birthDate || "");
  const birthDate = birthFromFields || birthFromPerson || birthFromAnketa;

  const workflowFlags = resolveUbdWorkflowFlags(document);

  return {
    callSign: readCallSign(personRow, profile, fields, rosterCallSign),
    lastName,
    firstName,
    patronymic,
    sex: formatUbdRegistrySex(
      getPersonFieldValue(personRow, ["стать"]) || text(anketa?.sex || ""),
    ),
    birthDate,
    idSeries,
    idNumber,
    rnokpp:
      readField(fields, "rnokpp") ||
      getPersonFieldValue(personRow, ["рнокпп_за_наявності"]) ||
      getPersonFieldValue(personRow, ["рнокпп"]) ||
      text(anketa?.rnokpp || ""),
    phone,
    rank:
      readField(fields, "rank") ||
      getPersonFieldValue(personRow, ["звання"]) ||
      text(anketa?.rank || ""),
    combatPlace: readField(fields, "taskPlace"),
    combatPeriod: formatCombatPeriod(readField(fields, "taskPeriod")),
    position: readShortPosition(personRow, fields),
    combatOrder: formatUbdCombatOrder(
      readField(fields, "basisNumber"),
      readField(fields, "basisDate"),
      readField(fields, "basis"),
    ),
    personnelStatus: text(personnelStatus || ""),
    ubdStatus: "",
    documents: workflowFlags.documents,
    raport: workflowFlags.raport,
    sent: workflowFlags.sent,
  };
};

const LEFT_ALIGN = new Set([1, 9, 13, 14]);

const buildHeaderRow = (labels: readonly string[]) =>
  labels.map((label) =>
    cell(label, {
      fontWeight: "bold",
      fontSize: 10,
      textColor: WHITE,
      backgroundColor: HEADER_GREEN,
      align: "center",
      wrap: true,
      height: 36,
    }),
  );

const buildSheet = (rows: UbdNotSubmittedExcelRow[]): SheetData => [
  buildHeaderRow(UBD_NOT_SUBMITTED_HEADERS),
  ...rows.map((row, index) => {
    const zebra =
      index % 2 === 1 ? { backgroundColor: ZEBRA } : { backgroundColor: WHITE };
    const values: Array<string | number | null> = [
      row.callSign,
      row.lastName,
      row.firstName,
      row.patronymic,
      row.sex,
      row.birthDate,
      row.idSeries,
      row.idNumber,
      row.rnokpp,
      row.phone,
      row.rank,
      row.combatPlace,
      row.combatPeriod,
      row.position,
      row.combatOrder,
      row.personnelStatus,
    ];
    return values.map((value, columnIndex) =>
      cell(value || null, {
        textColor: TEXT,
        wrap: true,
        align: LEFT_ALIGN.has(columnIndex) ? "left" : "center",
        ...zebra,
      }),
    );
  }),
];

const STATUS_LEFT_ALIGN = new Set([1, 8]);

const buildStatusSheet = (rows: UbdNotSubmittedExcelRow[]): SheetData => [
  buildHeaderRow(UBD_STATUS_HEADERS),
  ...rows.map((row, index) => {
    const zebra =
      index % 2 === 1 ? { backgroundColor: ZEBRA } : { backgroundColor: WHITE };
    const values: Array<string | number | null> = [
      row.callSign,
      row.lastName,
      row.firstName,
      row.patronymic,
      row.personnelStatus,
      row.ubdStatus,
      row.documents,
      row.raport,
      row.sent,
    ];
    return values.map((value, columnIndex) =>
      cell(value || null, {
        textColor: TEXT,
        wrap: true,
        align: STATUS_LEFT_ALIGN.has(columnIndex) ? "left" : "center",
        ...zebra,
      }),
    );
  }),
];

export const exportUbdNotSubmittedExcel = async ({
  rows,
}: {
  rows: UbdNotSubmittedExcelRow[];
  periodFilterLabel?: string;
}) => {
  const fileName = `1ПБ Рапорти для статусу УБД ${rows.length}.xlsx`;

  const columns = UBD_NOT_SUBMITTED_HEADERS.map((_, index) => {
    if (index === 13 || index === 14) return { width: 32 };
    if (index === 9) return { width: 18 };
    if (index === 8) return { width: 18 };
    if (index === 11 || index === 12) return { width: 28 };
    if (index === 1) return { width: 20 };
    return { width: 16 };
  });

  await writeXlsxFile(
    [
      {
        sheet: "Не подавалися",
        data: buildSheet(rows),
        orientation: "landscape",
        stickyRowsCount: 1,
        stickyColumnsCount: 2,
        showGridLines: true,
        columns,
      },
    ],
    { fontFamily: "Calibri", fontSize: 11 },
  ).toFile(fileName);

  return { fileName, rowCount: rows.length };
};

export const exportUbdStatusExcel = async ({
  rows,
  periodFilterLabel,
}: {
  rows: UbdNotSubmittedExcelRow[];
  periodFilterLabel?: string;
}) => {
  const exportedAt = dayjs();
  const periodSuffix =
    periodFilterLabel && periodFilterLabel !== "Усі місяці"
      ? ` ${periodFilterLabel}`
      : "";
  const fileName = `1ПБ УБД статус${periodSuffix} ${exportedAt.format("DD.MM.YYYY")}.xlsx`;

  await writeXlsxFile(
    [
      {
        sheet: "Статус",
        data: buildStatusSheet(rows),
        orientation: "landscape",
        stickyRowsCount: 1,
        stickyColumnsCount: 2,
        showGridLines: true,
        columns: [
          { width: 14 },
          { width: 20 },
          { width: 14 },
          { width: 16 },
          { width: 16 },
          { width: 14 },
          { width: 12 },
          { width: 12 },
          { width: 40 },
        ],
      },
    ],
    { fontFamily: "Calibri", fontSize: 11 },
  ).toFile(fileName);

  return { fileName, rowCount: rows.length };
};
