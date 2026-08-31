import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  formatUaPhoneDisplay,
  getPersonFieldValue,
  getPersonFullPositionTitle,
} from "../personnel/personnelUtils";
import {
  buildFighterTaskPeriodText,
  getFighterTaskPlace,
} from "../personnel/fighterStatusImport";

export type Form6Signatory = {
  blockType: "SIGNER" | "APPROVAL";
  title: string;
  rank: string;
  fullName: string;
  signatureData?: string | null;
};

export type Form6ReportFields = {
  commander: string;
  fullName: string;
  rank: string;
  staffPosition: string;
  birthDate: string;
  idDocument: string;
  rnokpp: string;
  address: string;
  phone: string;
  taskPeriod: string;
  taskPlace: string;
  basis: string;
  date: string;
  folderName: string;
  signatories: Form6Signatory[];
  statusNote: string;
};

const DEFAULT_FORM6_COMMANDER =
  "Командиру 1 піхотного батальйону військової частини А4862";
const DEFAULT_FORM6_BASIS =
  "Бойове розпорядження командира 425 ОШП «СКЕЛЯ» №4862/ОКП/1158/дск від 14.10.2025";

export const buildForm6FolderName = (fullName: string) => {
  const name = fullName.trim();
  return name ? `1ПБ РАПОРТ Форма 6 · ${name}` : "1ПБ РАПОРТ Форма 6";
};

const GIVEN_NAME_GENITIVE: Record<string, string> = {
  анатолій: "Анатолія",
  андрій: "Андрія",
  богдан: "Богдана",
  василь: "Василя",
  віктор: "Віктора",
  володимир: "Володимира",
  дмитро: "Дмитра",
  іван: "Івана",
  ігор: "Ігоря",
  максим: "Максима",
  микола: "Миколи",
  михайло: "Михайла",
  олег: "Олега",
  олександр: "Олександра",
  олексій: "Олексія",
  павло: "Павла",
  петро: "Петра",
  роман: "Романа",
  сергій: "Сергія",
  степан: "Степана",
  тарас: "Тараса",
  юрій: "Юрія",
  ярослав: "Ярослава",
};

export const form6WorkflowSteps = [
  { key: "document", title: "Заповнили рапорт" },
  { key: "account", title: "Фото 3x4 та ІНН" },
  { key: "scan", title: "Скани паспорта, ІНН та статусу" },
  { key: "ready", title: "Готово до відправки" },
  { key: "sent", title: "Відправили" },
  { key: "received", title: "Отримали" },
  { key: "handed", title: "Вручили" },
];

const titleCaseUk = (value: string) =>
  value
    ? value.charAt(0).toLocaleUpperCase("uk-UA") +
      value.slice(1).toLocaleLowerCase("uk-UA")
    : "";

const formatForm6Date = (value: Date) => {
  const pad = (part: number) => String(part).padStart(2, "0");
  return `${pad(value.getDate())}.${pad(value.getMonth() + 1)}.${value.getFullYear()}`;
};

export const toUkrainianGenitiveFullName = (fullName: string) => {
  const parts = fullName
    .trim()
    .replace(/\s*\([^)]+\)\s*$/, "")
    .split(/\s+/)
    .filter(Boolean);
  if (!parts.length) return "";

  const [surname, given, patronymic] = parts;
  let genitiveSurname = surname.toLocaleUpperCase("uk-UA");
  if (/КО$/i.test(genitiveSurname)) {
    genitiveSurname = `${genitiveSurname.slice(0, -1)}А`;
  } else if (/(УК|ЮК|АК|ИК)$/i.test(genitiveSurname)) {
    genitiveSurname = `${genitiveSurname}А`;
  } else if (/ИЙ$/i.test(genitiveSurname)) {
    genitiveSurname = `${genitiveSurname.slice(0, -2)}ОГО`;
  }

  const givenKey = (given ?? "").toLocaleLowerCase("uk-UA");
  const genitiveGiven = given
    ? GIVEN_NAME_GENITIVE[givenKey] ||
      (givenKey.endsWith("й")
        ? `${titleCaseUk(given.slice(0, -1))}я`
        : `${titleCaseUk(given)}а`)
    : "";

  let genitivePatronymic = patronymic ?? "";
  if (/ович$/i.test(genitivePatronymic)) {
    genitivePatronymic = genitivePatronymic.replace(/ович$/i, "овича");
  } else if (/івна$/i.test(genitivePatronymic) || /ївна$/i.test(genitivePatronymic)) {
    genitivePatronymic = genitivePatronymic.replace(/вна$/i, "вни");
  }

  return [genitiveSurname, genitiveGiven, genitivePatronymic]
    .filter(Boolean)
    .join(" ");
};

export const createForm6Fields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
  signatories: Form6Signatory[] = [],
): Form6ReportFields => {
  const fullName = summary.name !== "Особа не вибрана" ? summary.name : "";
  const staffPosition = getPersonFullPositionTitle(row);
  const documentName =
    getPersonFieldValue(row, ["назва", "документа", "посвідчує"]) ||
    "Паспорт громадянина України";
  const documentNumber = getPersonFieldValue(row, [
    "серія",
    "номер",
    "документа",
    "посвідчує",
  ]);
  const address =
    getPersonFieldValue(row, ["адреса", "проживан"]) ||
    getPersonFieldValue(row, ["зареєстров"]) ||
    getPersonFieldValue(row, ["місце", "проживан"]) ||
    getPersonFieldValue(row, ["адреса"]) ||
    "";
  const phone =
    summary.phonesDisplay[0] ||
    (summary.phones[0] ? formatUaPhoneDisplay(summary.phones[0]) : "");

  return {
    commander: DEFAULT_FORM6_COMMANDER,
    fullName,
    rank: summary.rank || "",
    staffPosition,
    birthDate: summary.birthDate || "",
    idDocument: [documentName, documentNumber].filter(Boolean).join(" "),
    rnokpp: summary.rnokpp || "",
    address,
    phone,
    taskPeriod: buildFighterTaskPeriodText(row),
    taskPlace: getFighterTaskPlace(row),
    basis: DEFAULT_FORM6_BASIS,
    date: formatForm6Date(new Date()),
    folderName: buildForm6FolderName(fullName),
    signatories,
    statusNote: "",
  };
};

export const mergeForm6Fields = (
  defaults: Form6ReportFields,
  value: unknown,
): Form6ReportFields => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return defaults;
  }
  const next = value as Partial<Form6ReportFields>;
  const merged = {
    ...defaults,
    ...next,
    signatories: Array.isArray(next.signatories)
      ? next.signatories
      : defaults.signatories,
  };
  // Personnel wins when filled; otherwise keep what was entered in the report.
  const pick = (personnel: string, document: string) =>
    String(personnel ?? "").trim() || String(document ?? "").trim();
  const fullName = pick(defaults.fullName, merged.fullName);
  const currentFolder = String(merged.folderName ?? "").trim();
  const folderName = (() => {
    if (!currentFolder) return buildForm6FolderName(fullName);
    if (/^1ПБ\s+РАПОРТ\s+Форма\s+6\s*$/i.test(currentFolder)) {
      return buildForm6FolderName(fullName);
    }
    if (/^Форма\s+6\b/i.test(currentFolder)) return buildForm6FolderName(fullName);
    if (
      /^1ПБ\s+РАПОРТ\s+Форма\s+6\b/i.test(currentFolder) &&
      fullName &&
      !currentFolder.includes(fullName)
    ) {
      return buildForm6FolderName(fullName);
    }
    return currentFolder;
  })();
  return {
    ...merged,
    fullName,
    rank: pick(defaults.rank, merged.rank),
    staffPosition: pick(defaults.staffPosition, merged.staffPosition),
    birthDate: pick(defaults.birthDate, merged.birthDate),
    idDocument: pick(defaults.idDocument, merged.idDocument),
    rnokpp: pick(defaults.rnokpp, merged.rnokpp),
    address: pick(defaults.address, merged.address),
    phone: pick(defaults.phone, merged.phone),
    folderName,
  };
};
