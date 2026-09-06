import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  getPersonFullPositionTitle,
} from "../personnel/personnelUtils";
import { toUkrainianGenitiveFullName } from "./form6Report";
import { capitalizeReportPosition } from "./reportPosition";

export type UbdRestoreSignatory = {
  blockType: "SIGNER" | "APPROVAL";
  title: string;
  rank: string;
  fullName: string;
  signatureData?: string | null;
};

export type UbdRestoreReportFields = {
  commander: string;
  fullName: string;
  rank: string;
  staffPosition: string;
  signerTitle: string;
  certificateSeries: string;
  circumstances: string;
  requestText: string;
  coveringCommander: string;
  date: string;
  signatureData: string;
  signatureFileName: string;
  folderName: string;
  signatories: UbdRestoreSignatory[];
  statusNote: string;
};

export const DEFAULT_UBD_RESTORE_COMMANDER = "Командиру батальйону розвідки";
export const DEFAULT_UBD_RESTORE_COVERING_COMMANDER =
  "Командиру військової частини А4862";
export const DEFAULT_UBD_RESTORE_CIRCUMSTANCES =
  "під час виконання службових обов’язків у темну пору доби, за складних погодних умов та обмеженої видимості, мною було пошкоджено посвідчення учасника бойових дій";
export const DEFAULT_UBD_RESTORE_REQUEST =
  "Прошу Вашого клопотання перед вищим командуванням про призначення за даним фактом службового розслідування, оформлення та видачу мені нового посвідчення учасника бойових дій у встановленому порядку.";

export const ubdRestoreWorkflowSteps = [
  { key: "document", title: "Заповнити Рапорт" },
  { key: "fighterSign", title: "Підпис службовця" },
  { key: "photo", title: "ФОТО" },
  { key: "scan", title: "Скани" },
  { key: "sent", title: "Відправили" },
  { key: "approved", title: "Погоджено" },
  { key: "transferred", title: "Передано ФОТО + Скани" },
  { key: "received", title: "Отримали" },
  { key: "handed", title: "Вручили" },
];

const RANK_GENITIVE: Record<string, string> = {
  солдат: "солдата",
  "старший солдат": "старшого солдата",
  "молодший сержант": "молодшого сержанта",
  сержант: "сержанта",
  "старший сержант": "старшого сержанта",
  "головний сержант": "головного сержанта",
  "штаб-сержант": "штаб-сержанта",
  "майстер-сержант": "майстер-сержанта",
  "старший майстер-сержант": "старшого майстер-сержанта",
  "головний майстер-сержант": "головного майстер-сержанта",
  "молодший лейтенант": "молодшого лейтенанта",
  лейтенант: "лейтенанта",
  "старший лейтенант": "старшого лейтенанта",
  капітан: "капітана",
  майор: "майора",
  підполковник: "підполковника",
  полковник: "полковника",
};

const titleCaseUk = (value: string) =>
  value
    ? value.charAt(0).toLocaleUpperCase("uk-UA") +
      value.slice(1).toLocaleLowerCase("uk-UA")
    : "";

const formatRestoreDate = (value: Date) => {
  const pad = (part: number) => String(part).padStart(2, "0");
  return `${pad(value.getDate())}.${pad(value.getMonth() + 1)}.${value.getFullYear()}`;
};

export const formatGivenSurname = (fullName: string) => {
  const parts = fullName
    .trim()
    .replace(/\s*\([^)]+\)\s*$/, "")
    .split(/\s+/)
    .filter(Boolean);
  if (!parts.length) return "";
  if (parts.length === 1) return parts[0].toLocaleUpperCase("uk-UA");

  const first = parts[0];
  const last = parts[parts.length - 1];
  const surnameFirst =
    first === first.toLocaleUpperCase("uk-UA") && /[А-ЯІЇЄҐA-Z]/.test(first);
  if (surnameFirst) {
    return `${titleCaseUk(parts[1] ?? first)} ${first.toLocaleUpperCase("uk-UA")}`;
  }
  return `${titleCaseUk(first)} ${last.toLocaleUpperCase("uk-UA")}`;
};

export const toUkrainianGenitiveRank = (rank: string) => {
  const key = rank.trim().toLocaleLowerCase("uk-UA");
  if (!key) return "";
  if (RANK_GENITIVE[key]) return RANK_GENITIVE[key];
  return key
    .split(/\s+/)
    .map((word) => {
      if (word.endsWith("ий") || word.endsWith("ій")) {
        return `${word.slice(0, -2)}ого`;
      }
      if (/[нтдк]$/i.test(word)) return `${word}а`;
      return word;
    })
    .join(" ");
};

export const toUkrainianGenitiveGivenSurname = (fullName: string) => {
  const genitive = toUkrainianGenitiveFullName(fullName);
  const parts = genitive.split(/\s+/).filter(Boolean);
  if (parts.length >= 2) return `${parts[1]} ${parts[0]}`;
  return genitive;
};

export const formatPositionTitleBlock = (position: string) => {
  const text = capitalizeReportPosition(position);
  if (!text) return "";
  const unit = text.match(/^(.*?)[, ]+(військової частини\s+\S+)$/i);
  if (!unit) return text;
  const head = unit[1].trim();
  const company = head.match(/^(.*?)\s+(роти\s+.*)$/i);
  if (company) return `${company[1]}\n${company[2]}\n${unit[2]}`;
  return `${head}\n${unit[2]}`;
};

/** Перший рядок блоку посади — з великої (другий рядок «військової частини…» без змін). */
export const capitalizeSignerTitleBlock = (value: string) => {
  const lines = String(value ?? "").split(/\r?\n/);
  let capped = false;
  return lines
    .map((line) => {
      const trimmed = line.trim();
      if (!trimmed || capped) return line;
      capped = true;
      const leading = line.match(/^\s*/)?.[0] ?? "";
      return `${leading}${capitalizeReportPosition(trimmed)}`;
    })
    .join("\n");
};

export const resolveUbdRestoreSignerTitle = (
  fields: Pick<UbdRestoreReportFields, "staffPosition" | "signerTitle">,
) => {
  const fromStaff = formatPositionTitleBlock(fields.staffPosition);
  if (fromStaff) return fromStaff;
  return capitalizeSignerTitleBlock(fields.signerTitle);
};

const UBD_RESTORE_RANK_TAIL =
  /(головний майстер-сержант|старший майстер-сержант|майстер-сержант|штаб-сержант|головний сержант|старший сержант|молодший сержант|старший лейтенант|молодший лейтенант|старший солдат|підполковник|полковник|лейтенант|сержант|капітан|майор|солдат|рекрут)$/iu;

const isUbdRestoreDateLine = (value: string) =>
  /^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value);

const normalizeCommanderUnitLine = (value: string) =>
  value
    .trim()
    .toLocaleLowerCase("uk-UA")
    .replace(/^командир(?:а)?\s+/iu, "")
    .replace(/\s+/g, " ");

const isDuplicateCommanderUnitLine = (first: string, second: string) => {
  const left = first.trim();
  const right = second.trim();
  if (!left || !right) return false;
  return normalizeCommanderUnitLine(left) === normalizeCommanderUnitLine(right);
};

export type UbdRestoreSignatoryParts = {
  titleLines: [string, string];
  rank: string;
  fullName: string;
  date: string;
  signatureData: string;
};

export const splitUbdRestoreSignatory = (
  signatory: {
    title: string;
    rank: string;
    fullName: string;
    signatureData?: string | null;
  } | null,
): UbdRestoreSignatoryParts => {
  if (!signatory) {
    return {
      titleLines: ["", ""],
      rank: "",
      fullName: "",
      date: "",
      signatureData: "",
    };
  }
  const lines = signatory.title
    .split(/\r?\n/)
    .map((line) => line.trim())
    .flatMap((line) => {
      const match = line.match(/^(.*?)(?:\s+)(\d{1,2}\.\d{1,2}\.\d{4})$/);
      if (match?.[1]?.trim()) return [match[1].trim(), match[2]];
      return [line];
    })
    .filter(Boolean);
  const date = lines.find(isUbdRestoreDateLine) ?? "";
  let rank = signatory.rank.trim();
  const titles: string[] = [];
  for (const line of lines) {
    if (isUbdRestoreDateLine(line)) continue;
    const tail = line.match(UBD_RESTORE_RANK_TAIL);
    if (tail && !rank) {
      rank = tail[1];
      const head = line.slice(0, -tail[1].length).trim();
      if (head) titles.push(head);
      continue;
    }
    if (
      rank &&
      line.toLocaleLowerCase("uk-UA") === rank.toLocaleLowerCase("uk-UA")
    ) {
      continue;
    }
    titles.push(line);
  }
  if (titles.length === 1) {
    const split = titles[0].match(/^(.*?)\s+(військової частини\s+\S+.*)$/i);
    if (split) {
      titles[0] = split[1].trim();
      titles[1] = split[2].trim();
    }
  }
  while (titles.length < 2) titles.push("");
  if (titles.length > 2) {
    titles.splice(1, titles.length, titles.slice(1).join(" "));
  }
  return {
    titleLines: [titles[0], titles[1]],
    rank,
    fullName:
      formatGivenSurname(signatory.fullName) || signatory.fullName.trim(),
    date,
    signatureData: signatory.signatureData?.trim() ?? "",
  };
};

export const resolveUbdRestoreSignatoryParts = (
  signatory: UbdRestoreSignatory | null | undefined,
  reportDate: string,
): UbdRestoreSignatoryParts => {
  const parts = splitUbdRestoreSignatory(signatory ?? null);
  if (signatory?.rank.trim()) parts.rank = signatory.rank.trim();
  if (signatory?.title.trim()) {
    const titleBlock = signatory.title.includes("\n")
      ? capitalizeSignerTitleBlock(signatory.title)
      : formatPositionTitleBlock(signatory.title) ||
        capitalizeSignerTitleBlock(signatory.title);
    const parsedTitleLines = titleBlock
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line && !isUbdRestoreDateLine(line));
    if (parsedTitleLines[0]) parts.titleLines[0] = parsedTitleLines[0];
    if (parsedTitleLines[1]) parts.titleLines[1] = parsedTitleLines[1];
  }
  const normalizedRank = parts.rank.trim().toLocaleLowerCase("uk-UA");
  const titleWithoutRank = [parts.titleLines[0], parts.titleLines[1]]
    .map((line) => line.trim())
    .filter(
      (line) =>
        line &&
        (!normalizedRank ||
          line.toLocaleLowerCase("uk-UA") !== normalizedRank),
    );
  parts.titleLines = [titleWithoutRank[0] ?? "", titleWithoutRank[1] ?? ""];
  if (isDuplicateCommanderUnitLine(parts.titleLines[0], parts.titleLines[1])) {
    parts.titleLines[1] = "";
  }
  if (signatory?.fullName.trim()) {
    parts.fullName =
      formatGivenSurname(signatory.fullName) || signatory.fullName.trim();
  }
  if (signatory?.signatureData?.trim()) {
    parts.signatureData = signatory.signatureData.trim();
  }
  parts.date = parts.date || reportDate.trim();
  return parts;
};

export const resolveCoveringSignerParts = (fields: UbdRestoreReportFields) =>
  resolveUbdRestoreSignatoryParts(
    fields.signatories.find((item) => item.blockType === "SIGNER"),
    fields.date,
  );

export const resolveApproverParts = (fields: UbdRestoreReportFields) =>
  resolveUbdRestoreSignatoryParts(
    fields.signatories.find((item) => item.blockType === "APPROVAL"),
    fields.date,
  );

export const buildUbdRestoreBody = (fields: UbdRestoreReportFields) => {
  const rank = fields.rank.trim() || "звання";
  const fullName = fields.fullName.trim() || "______";
  const position = capitalizeReportPosition(fields.staffPosition) || "посада";
  const circumstances =
    fields.circumstances.trim() || DEFAULT_UBD_RESTORE_CIRCUMSTANCES;
  const series = fields.certificateSeries.trim() || "серія ______";
  return `Дійсним доповідаю, що я, ${rank} ${fullName}, ${position}, ${circumstances}, ${series}. Пошкодження сталося ненавмисно, під час активного пересування та виконання поставленого завдання.`;
};

export const buildUbdRestorePetition = (fields: UbdRestoreReportFields) => {
  const rank = toUkrainianGenitiveRank(fields.rank) || "звання";
  const name =
    toUkrainianGenitiveGivenSurname(fields.fullName) ||
    formatGivenSurname(fields.fullName) ||
    "______";
  return `Клопочу по суті рапорту ${rank} ${name}`;
};

export const buildUbdRestoreFolderName = (fullName: string) => {
  const name = fullName.trim();
  return name
    ? `1ПБ Рапорт на відновлення УБД · ${name}`
    : "1ПБ Рапорт на відновлення УБД";
};

export const resolveUbdRestoreFolderName = (
  fullName: string,
  currentFolder?: string,
) => {
  const current = String(currentFolder ?? "").trim();
  if (!current) return buildUbdRestoreFolderName(fullName);
  if (/^1ПБ/i.test(current)) return current;
  if (/^Рапорт на відновлення УБД/i.test(current)) {
    return buildUbdRestoreFolderName(fullName);
  }
  return current;
};

export const createUbdRestoreFields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
  signatories: UbdRestoreSignatory[] = [],
): UbdRestoreReportFields => {
  const fullName = summary.name !== "Особа не вибрана" ? summary.name : "";
  const staffPosition = getPersonFullPositionTitle(row);

  return {
    commander: DEFAULT_UBD_RESTORE_COMMANDER,
    fullName,
    rank: summary.rank || "",
    staffPosition: capitalizeReportPosition(staffPosition),
    signerTitle: formatPositionTitleBlock(staffPosition),
    certificateSeries: "серія ______",
    circumstances: DEFAULT_UBD_RESTORE_CIRCUMSTANCES,
    requestText: DEFAULT_UBD_RESTORE_REQUEST,
    coveringCommander: DEFAULT_UBD_RESTORE_COVERING_COMMANDER,
    date: formatRestoreDate(new Date()),
    signatureData: "",
    signatureFileName: "",
    folderName: buildUbdRestoreFolderName(fullName),
    signatories,
    statusNote: "",
  };
};

export const mergeUbdRestoreFields = (
  defaults: UbdRestoreReportFields,
  value: unknown,
): UbdRestoreReportFields => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return defaults;
  }
  const next = {
    ...defaults,
    ...(value as Partial<UbdRestoreReportFields>),
    signatories: defaults.signatories,
  };
  const fullName = String(next.fullName ?? "").trim();
  const folderName = resolveUbdRestoreFolderName(
    fullName,
    String(next.folderName ?? ""),
  );
  return {
    ...next,
    folderName,
    staffPosition: capitalizeReportPosition(next.staffPosition),
    signerTitle: resolveUbdRestoreSignerTitle(next),
  };
};
