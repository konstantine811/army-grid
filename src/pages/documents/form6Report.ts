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
import {
  pickUbdBasisOrderForTaskPeriod,
  resolveUbdBasisForTask,
  UBD_BASIS_ORDER_OPTIONS,
} from "./ubdBasisOrders";
import { capitalizeReportPosition } from "./reportPosition";

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
  basisNumber: string;
  basisDate: string;
  basis: string;
  /** БР введено вручну — не підставляти зі списку за періодом завдань. */
  basisManual: boolean;
  date: string;
  /** Навіщо роблять цю форму — колонка в Excel-експорті журналу. */
  formPurpose: string;
  folderName: string;
  signatories: Form6Signatory[];
  statusNote: string;
};

const DEFAULT_FORM6_COMMANDER =
  "Командиру 1 піхотного батальйону військової частини А4862";
const FALLBACK_FORM6_BASIS = pickUbdBasisOrderForTaskPeriod("") ?? {
  number: "4862/ОКП/1162/дск",
  date: "09.05.2026",
};

const stripForm6BasisNumber = (value: string) =>
  String(value ?? "")
    .trim()
    .replace(/^№\s*/, "");

const form6IdDocumentHasNumber = (value: unknown) =>
  String(value ?? "").replace(/\D/g, "").length >= 6;

/** Якщо в картці лише назва паспорта, не затирати збережений номер у Формі 6. */
const pickForm6IdDocument = (personnel: string, document: string) => {
  const fromPersonnel = String(personnel ?? "").trim();
  const fromDocument = String(document ?? "").trim();
  if (form6IdDocumentHasNumber(fromPersonnel)) return fromPersonnel;
  if (form6IdDocumentHasNumber(fromDocument)) return fromDocument;
  return fromPersonnel || fromDocument;
};

/** Текст після «Підстава:» у Word — саме слово «Підстава:» відкидаємо, пробіл на початку лишаємо. */
export const stripForm6BasisLabel = (value: string) =>
  String(value ?? "").replace(/^\s*Підстава:\s*/i, "");

export const extractForm6BasisParts = (value: string) => {
  const text = stripForm6BasisLabel(value);
  const match = text.match(
    /№\s*(\S+)\s+від\s+(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/i,
  );
  if (!match) {
    return { basisNumber: "", basisDate: "" };
  }
  return {
    basisNumber: match[1],
    basisDate: match[2].replaceAll("/", ".").replaceAll("-", "."),
  };
};

/** Підстава Форми 6: той самий БР, що й для УБД, у формулюванні рапорту Ф6. */
export const formatForm6BasisText = (number: string, date: string) => {
  const orderNumber = stripForm6BasisNumber(number);
  const orderDate = String(date ?? "").trim();
  const reference = [
    orderNumber ? `№${orderNumber}` : "",
    orderDate ? `від ${orderDate}` : "",
  ]
    .filter(Boolean)
    .join(" ");
  return ["Бойове розпорядження командира 425 ОШП «СКЕЛЯ»", reference]
    .filter(Boolean)
    .join(" ");
};

export const parseForm6BasisParts = (value: string) => {
  const extracted = extractForm6BasisParts(value);
  if (!extracted.basisNumber && !extracted.basisDate) {
    return {
      basisNumber: FALLBACK_FORM6_BASIS.number,
      basisDate: FALLBACK_FORM6_BASIS.date,
    };
  }
  return extracted;
};

export const form6BasisLineForWord = (basis: string) => {
  const body = stripForm6BasisLabel(basis);
  if (!body.trim()) return "Підстава: ______";
  return /^\s/.test(body) ? `Підстава:${body}` : `Підстава: ${body}`;
};

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
  const taskPeriod = buildFighterTaskPeriodText(row);
  const taskPlace = getFighterTaskPlace(row);
  const resolved = resolveUbdBasisForTask(taskPeriod, taskPlace);
  const basis = resolved ?? FALLBACK_FORM6_BASIS;

  return {
    commander: DEFAULT_FORM6_COMMANDER,
    fullName,
    rank: summary.rank || "",
    staffPosition: capitalizeReportPosition(staffPosition),
    birthDate: summary.birthDate || "",
    idDocument: [documentName, documentNumber].filter(Boolean).join(" "),
    rnokpp: summary.rnokpp || "",
    address,
    phone,
    taskPeriod,
    taskPlace,
    basisNumber: basis.number,
    basisDate: basis.date,
    basis: formatForm6BasisText(basis.number, basis.date),
    basisManual: false,
    date: formatForm6Date(new Date()),
    formPurpose: "",
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
  const extracted = extractForm6BasisParts(
    String(next.basisNumber ? "" : next.basis ?? defaults.basis),
  );
  const explicitBasisNumber = stripForm6BasisNumber(
    String(next.basisNumber ?? ""),
  );
  const explicitBasisDate = String(next.basisDate ?? "").trim();
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
  const taskPeriod = pick(defaults.taskPeriod, merged.taskPeriod);
  const taskPlace = pick(defaults.taskPlace, merged.taskPlace);
  const basisManual =
    next.basisManual === true || String(next.basisManual) === "true";
  const resolved = resolveUbdBasisForTask(taskPeriod, taskPlace);
  const picked = resolved
    ? { number: resolved.number, date: resolved.date }
    : pickUbdBasisOrderForTaskPeriod(taskPeriod, undefined, taskPlace);
  let basisNumber = explicitBasisNumber;
  let basisDate = explicitBasisDate;
  if (!basisManual) {
    const savedBasisNumber =
      explicitBasisNumber ||
      extracted.basisNumber ||
      FALLBACK_FORM6_BASIS.number;
    const savedBasisDate =
      explicitBasisDate || extracted.basisDate || FALLBACK_FORM6_BASIS.date;
    basisNumber = savedBasisNumber;
    basisDate = savedBasisDate;
    if (resolved?.matches.length) {
      const savedIsMatch =
        Boolean(savedBasisNumber) &&
        (resolved.number === savedBasisNumber ||
          resolved.matches.some((item) => item.number === savedBasisNumber));
      if (savedIsMatch) {
        basisNumber = savedBasisNumber;
        basisDate = savedBasisDate || resolved.date;
      } else {
        basisNumber = resolved.number;
        basisDate = resolved.date;
      }
    } else if (picked) {
      const savedMatchesPickedDate =
        savedBasisDate.replaceAll("/", ".").replaceAll("-", ".") === picked.date;
      const savedIsKnownForDate = UBD_BASIS_ORDER_OPTIONS.some(
        (item) =>
          item.date === picked.date && item.number === savedBasisNumber,
      );
      if (savedMatchesPickedDate && savedIsKnownForDate) {
        basisNumber = savedBasisNumber;
        basisDate = picked.date;
      } else {
        basisNumber = picked.number;
        basisDate = picked.date;
      }
    }
  }
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
    staffPosition: capitalizeReportPosition(
      pick(defaults.staffPosition, merged.staffPosition),
    ),
    birthDate: pick(defaults.birthDate, merged.birthDate),
    idDocument: pickForm6IdDocument(defaults.idDocument, merged.idDocument),
    rnokpp: pick(defaults.rnokpp, merged.rnokpp),
    address: pick(defaults.address, merged.address),
    phone: pick(defaults.phone, merged.phone),
    taskPeriod,
    taskPlace,
    basisNumber,
    basisDate,
    basis: basisManual
      ? (() => {
          const typed = stripForm6BasisLabel(
            String(next.basis ?? merged.basis ?? ""),
          );
          if (typed.trim()) return typed;
          if (!basisNumber && !basisDate) return typed;
          return formatForm6BasisText(basisNumber, basisDate);
        })()
      : formatForm6BasisText(basisNumber, basisDate),
    basisManual,
    formPurpose: String(merged.formPurpose ?? ""),
    folderName,
  };
};
