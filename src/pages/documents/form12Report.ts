import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  getPersonFullPositionTitle,
} from "../personnel/personnelUtils";
import { toUkrainianGenitiveFullName } from "./form6Report";
import {
  formatGivenSurname,
  formatPositionTitleBlock,
} from "./ubdRestoreReport";

export type Form12Signatory = {
  blockType: "SIGNER" | "APPROVAL";
  title: string;
  rank: string;
  fullName: string;
  signatureData?: string | null;
};

export type Form12ReportFields = {
  commander: string;
  fullName: string;
  rank: string;
  staffPosition: string;
  date: string;
  signatureData: string;
  signatureFileName: string;
  folderName: string;
  signatories: Form12Signatory[];
  statusNote: string;
};

export const form12WorkflowSteps = [
  { key: "document", title: "Заповнили Рапорт" },
  { key: "fighterSign", title: "Підпис службовця" },
  { key: "sent", title: "Відправили" },
  { key: "received", title: "Отримали" },
  { key: "handed", title: "Вручили" },
];

const DEFAULT_FORM12_COMMANDER = "Командиру 1 піхотного батальйону";

export const buildForm12FolderName = (fullName: string) => {
  const name = fullName.trim();
  return name ? `1ПБ РАПОРТ Форма 12 · ${name}` : "1ПБ РАПОРТ Форма 12";
};

const RANK_DATIVE: Record<string, string> = {
  солдат: "солдату",
  "старший солдат": "старшому солдату",
  "молодший сержант": "молодшому сержанту",
  сержант: "сержанту",
  "старший сержант": "старшому сержанту",
  "головний сержант": "головному сержанту",
  "штаб-сержант": "штаб-сержанту",
  "майстер-сержант": "майстер-сержанту",
  "старший майстер-сержант": "старшому майстер-сержанту",
  "головний майстер-сержант": "головному майстер-сержанту",
  "молодший лейтенант": "молодшому лейтенанту",
  лейтенант: "лейтенанту",
  "старший лейтенант": "старшому лейтенанту",
  капітан: "капітану",
  майор: "майору",
  підполковник: "підполковнику",
  полковник: "полковнику",
};

const GIVEN_NAME_DATIVE: Record<string, string> = {
  анатолій: "Анатолію",
  андрій: "Андрію",
  богдан: "Богдану",
  вадим: "Вадиму",
  василь: "Василю",
  віктор: "Віктору",
  володимир: "Володимиру",
  дмитро: "Дмитру",
  іван: "Івану",
  ігор: "Ігорю",
  максим: "Максиму",
  микола: "Миколі",
  михайло: "Михайлу",
  олег: "Олегу",
  олександр: "Олександру",
  олексій: "Олексію",
  павло: "Павлу",
  петро: "Петру",
  роман: "Роману",
  сергій: "Сергію",
  степан: "Степану",
  тарас: "Тарасу",
  юрій: "Юрію",
  ярослав: "Ярославу",
};

const titleCaseUk = (value: string) =>
  value
    ? value.charAt(0).toLocaleUpperCase("uk-UA") +
      value.slice(1).toLocaleLowerCase("uk-UA")
    : "";

const formatForm12Date = (value: Date) => {
  const pad = (part: number) => String(part).padStart(2, "0");
  return `${pad(value.getDate())}.${pad(value.getMonth() + 1)}.${value.getFullYear()}`;
};

const toDativeSurname = (surname: string) => {
  const value = surname.toLocaleUpperCase("uk-UA");
  if (/АГА$/u.test(value) || /ОГА$/u.test(value)) {
    return `${value.slice(0, -2)}ЗІ`;
  }
  if (/КА$/u.test(value)) return `${value.slice(0, -1)}І`;
  if (/А$/u.test(value)) return `${value.slice(0, -1)}І`;
  if (/КО$/u.test(value)) return `${value.slice(0, -1)}У`;
  if (/ИЙ$/u.test(value)) return `${value.slice(0, -2)}ОМУ`;
  if (/(УК|ЮК|АК|ИК|ОВ|ЕВ|ЄВ|ІН|ИН)$/u.test(value)) return `${value}У`;
  return `${value}У`;
};

const toDativeGiven = (given: string) => {
  const key = given.toLocaleLowerCase("uk-UA");
  if (GIVEN_NAME_DATIVE[key]) return GIVEN_NAME_DATIVE[key];
  if (key.endsWith("й")) return `${titleCaseUk(given.slice(0, -1))}ю`;
  if (key.endsWith("о")) return `${titleCaseUk(given.slice(0, -1))}у`;
  return `${titleCaseUk(given)}у`;
};

const toDativePatronymic = (patronymic: string) => {
  if (/ович$/iu.test(patronymic)) {
    return patronymic.replace(/ович$/iu, "овичу");
  }
  if (/івна$/iu.test(patronymic) || /ївна$/iu.test(patronymic)) {
    return patronymic.replace(/вна$/iu, "вні");
  }
  return patronymic;
};

export const toUkrainianDativeRank = (rank: string) => {
  const key = rank.trim().toLocaleLowerCase("uk-UA");
  if (!key) return "";
  if (RANK_DATIVE[key]) return RANK_DATIVE[key];
  return key
    .split(/\s+/)
    .map((word) => {
      if (word.endsWith("ий") || word.endsWith("ій")) {
        return `${word.slice(0, -2)}ому`;
      }
      if (/[нтдк]$/i.test(word)) return `${word}у`;
      return word;
    })
    .join(" ");
};

export const toUkrainianDativeFullName = (fullName: string) => {
  const parts = fullName
    .trim()
    .replace(/\s*\([^)]+\)\s*$/, "")
    .split(/\s+/)
    .filter(Boolean);
  if (!parts.length) return "";
  const [surname, given, patronymic] = parts;
  return [toDativeSurname(surname), given ? toDativeGiven(given) : "", patronymic ? toDativePatronymic(patronymic) : ""]
    .filter(Boolean)
    .join(" ");
};

export const form12FighterSignName = (fullName: string) =>
  formatGivenSurname(fullName);

export const form12PositionLines = (position: string) =>
  formatPositionTitleBlock(position)
    .split(/\n/)
    .map((line) => line.trim())
    .filter(Boolean);

const isDateLine = (value: string) =>
  /^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value.trim());

const splitEmbeddedDate = (line: string) => {
  const match = line.match(/^(.*?)(?:\s+)(\d{1,2}\.\d{1,2}\.\d{4})$/);
  if (match?.[1]?.trim()) return [match[1].trim(), match[2]];
  return [line];
};

const stripTrailingRank = (line: string, rank: string) => {
  if (!rank) return line;
  const escaped = rank.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return line.replace(new RegExp(`\\s+${escaped}$`, "iu"), "").trim();
};

export const splitForm12Signatory = (
  signatory: Form12Signatory | null,
  lineCount: number,
) => {
  if (!signatory) {
    return {
      titleLines: Array.from({ length: lineCount }, () => ""),
      rank: "",
      date: "",
      fullName: "",
    };
  }

  const rawLines = signatory.title
    .replace(/^ЗАТВЕРДЖУЮ[\s:–—-]*/iu, "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .flatMap(splitEmbeddedDate);

  const date = rawLines.find(isDateLine) ?? "";
  const rankFromTitle =
    [...rawLines].reverse().find(
      (line) =>
        !isDateLine(line) &&
        !/командир|тимчасово|військової частини|затверджую/i.test(line),
    ) ?? "";
  const rank = signatory.rank.trim() || rankFromTitle;
  const titleLines = rawLines
    .filter(
      (line) =>
        line !== date &&
        line.toLocaleLowerCase("uk-UA") !== rank.toLocaleLowerCase("uk-UA"),
    )
    .map((line) => stripTrailingRank(line, rank))
    .filter(Boolean);

  const padded = [...titleLines];
  while (padded.length < lineCount) padded.push("");
  if (padded.length > lineCount) {
    padded.splice(
      lineCount - 1,
      padded.length,
      padded.slice(lineCount - 1).join(" "),
    );
  }

  return {
    titleLines: padded.slice(0, lineCount),
    rank,
    date,
    fullName: signatory.fullName.trim(),
  };
};

export const toUkrainianDativePosition = (position: string) => {
  const text = position.trim();
  if (!text) return "";
  const [first, ...rest] = text.split(/\s+/);
  const key = first.toLocaleLowerCase("uk-UA");
  let dativeFirst = key;
  if (key.endsWith("ець")) dativeFirst = `${key.slice(0, -3)}цю`;
  else if (key.endsWith("ий") || key.endsWith("ій")) {
    dativeFirst = `${key.slice(0, -2)}ому`;
  } else if (!/(у|ю|і|ї)$/u.test(key)) dativeFirst = `${key}у`;
  return [dativeFirst, ...rest].join(" ");
};

export const form12PleaText = (fields: Form12ReportFields) => {
  const dativeRank = toUkrainianDativeRank(fields.rank) || "______";
  const dativeName =
    toUkrainianDativeFullName(fields.fullName) ||
    toUkrainianGenitiveFullName(fields.fullName) ||
    "______";
  const position = toUkrainianDativePosition(
    fields.staffPosition.replace(/[.,;:\s]+$/u, ""),
  ) || "______";
  return `Прошу Вашого клопотання перед вищим командуванням на надання мені, ${dativeRank} ${dativeName}, ${position}, довідки Ф-12 (довідки про безпосередню участь у бойових діях).`;
};

export const createForm12Fields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
  signatories: Form12Signatory[] = [],
): Form12ReportFields => {
  const fullName = summary.name !== "Особа не вибрана" ? summary.name : "";
  const staffPosition = getPersonFullPositionTitle(row);

  return {
    commander: DEFAULT_FORM12_COMMANDER,
    fullName,
    rank: summary.rank || "",
    staffPosition,
    date: formatForm12Date(new Date()),
    signatureData: "",
    signatureFileName: "",
    folderName: buildForm12FolderName(fullName),
    signatories,
    statusNote: "",
  };
};

export const mergeForm12Fields = (
  defaults: Form12ReportFields,
  value: unknown,
): Form12ReportFields => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return defaults;
  }
  const next = value as Partial<Form12ReportFields>;
  const merged = {
    ...defaults,
    ...next,
    signatories: Array.isArray(next.signatories)
      ? next.signatories
      : defaults.signatories,
  };
  // Prefer the saved document name so opening another card cannot rewrite identity.
  const fullName =
    String(merged.fullName ?? "").trim() ||
    String(defaults.fullName ?? "").trim();
  const currentFolder = String(merged.folderName ?? "").trim();
  const folderName = (() => {
    if (!currentFolder) return buildForm12FolderName(fullName);
    if (/^1ПБ\s+РАПОРТ\s+Форма\s+12\s*$/i.test(currentFolder)) {
      return buildForm12FolderName(fullName);
    }
    if (/^Форма\s+12\b/i.test(currentFolder)) {
      return buildForm12FolderName(fullName);
    }
    if (
      /^1ПБ\s+РАПОРТ\s+Форма\s+12\b/i.test(currentFolder) &&
      fullName &&
      !currentFolder.includes(fullName)
    ) {
      return buildForm12FolderName(fullName);
    }
    return currentFolder;
  })();
  return {
    ...merged,
    fullName: fullName || merged.fullName,
    folderName,
  };
};
