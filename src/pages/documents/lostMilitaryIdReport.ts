import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  getPersonFieldValue,
  getPersonFullPositionTitle,
} from "../personnel/personnelUtils";
import {
  formatNominativeGivenSurname,
  toUkrainianDativeFullName,
  toUkrainianDativeRank,
  toUkrainianGenitiveFullName,
  toUkrainianGenitiveGivenSurname,
  toUkrainianGenitiveRank,
  toUkrainianInstrumentalFullName,
  toUkrainianInstrumentalPosition,
  toUkrainianInstrumentalRank,
  toUkrainianNominativePosition,
} from "./lostMilitaryIdCases";
import { toUkrainianDativePosition, looksLikeUaDateToken } from "./form12Report";
import { capitalizeReportPosition } from "./reportPosition";
import { formatPositionTitleBlock } from "./ubdRestoreReport";

export const lostMilitaryIdWorkflowSteps = [
  { key: "document", title: "Заповнили рапорт" },
  { key: "sent", title: "Відправили рапорт" },
  { key: "order", title: "Наказ" },
  { key: "act", title: "Акт" },
  { key: "received", title: "Отримали" },
  { key: "handed", title: "Вручили" },
];

export type LostMilitaryIdSignatory = {
  blockType: "SIGNER" | "APPROVAL";
  title: string;
  rank: string;
  fullName: string;
  signatureData?: string | null;
};

export type LostMilitaryIdFields = {
  addressee: string;
  militaryUnit: string;
  fullName: string;
  rank: string;
  staffPosition: string;
  unitLabel: string;
  lossDate: string;
  isExactDate: boolean;
  circumstanceKind: "movement" | "custom";
  /** Місце події (с. Гришене тощо) — коли не переміщення. */
  lossLocation: string;
  fromLocation: string;
  toLocation: string;
  customCircumstances: string;
  searchConducted: boolean;
  searchResult: string;
  reporterTitle: string;
  reporterRank: string;
  reporterFullName: string;
  investigatorFullName: string;
  investigatorRank: string;
  investigatorPosition: string;
  investigatorPersonId: string;
  investigatorManual: boolean;
  reportDate: string;
  orderNumber: string;
  orderDate: string;
  approvalDate: string;
  reportNumber: string;
  personnelChiefName: string;
  birthDate: string;
  enlistedDate: string;
  enlistedOrder: string;
  education: string;
  maritalStatus: string;
  address: string;
  citizenship: string;
  folderName: string;
  statusNote: string;
  editManually: boolean;
  fullNameInstrumentalManual: string;
  investigatorDativeManual: string;
  reportTextOverride: string;
  orderTextOverride: string;
  actCircumstancesOverride: string;
  actConclusionsOverride: string;
  actProposalsOverride: string;
  signatories: LostMilitaryIdSignatory[];
};

const DEFAULT_UNIT = "А4862";
const DEFAULT_ADDRESSEE = "Командиру військової частини А4862";
const DEFAULT_REPORTER_TITLE =
  "Тимчасово виконуючий обов’язки командира 1 піхотного батальйону військової частини А4862";
const DEFAULT_REPORTER_RANK = "старший лейтенант";
const DEFAULT_REPORTER_NAME = "Андрій КІЯНЕНКО";
const DEFAULT_ORDER_COMMANDER_TITLE =
  "Тимчасово виконуючий обов’язки\nкомандира військової частини А4862";
const DEFAULT_ORDER_COMMANDER_RANK = "капітан";
const DEFAULT_ORDER_COMMANDER_NAME = "Олег АДАМОВ";
const DEFAULT_SEARCH_RESULT = "військовий квиток не знайдено";

const MONTHS_UK = [
  "січня",
  "лютого",
  "березня",
  "квітня",
  "травня",
  "червня",
  "липня",
  "серпня",
  "вересня",
  "жовтня",
  "листопада",
  "грудня",
];

export const formatUaDate = (value: Date) => {
  const pad = (part: number) => String(part).padStart(2, "0");
  return `${pad(value.getDate())}.${pad(value.getMonth() + 1)}.${value.getFullYear()}`;
};

export const formatUaLongDate = (value: string) => {
  const match = value.trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (!match) return value.trim();
  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = match[3];
  const monthName = MONTHS_UK[month - 1];
  if (!monthName) return value.trim();
  return `${day} ${monthName} ${year} року`;
};

/** Номер в/ч без префікса «військової частини» (напр. «А4862»). */
export const militaryUnitLabel = (
  unit: string,
  fallback = DEFAULT_UNIT,
) => {
  const text = unit.trim() || fallback;
  const match = text.match(/військової\s+частини\s+(.+)/iu);
  return (match?.[1] ?? text).replace(/\s+/g, " ").trim() || fallback;
};

/** «військової частини А4862» — без подвоєння, якщо префікс уже в полі. */
export const normalizeMilitaryUnitPhrase = (
  unit: string,
  fallback = DEFAULT_UNIT,
) => {
  const text = unit.trim() || fallback;
  if (/військової\s+частини/iu.test(text)) {
    return text.replace(/\s+/g, " ").trim();
  }
  return `військової частини ${text}`;
};

export const buildLostMilitaryIdFolderName = (fullName: string) => {
  const name = fullName.trim();
  return name
    ? `1ПБ Втрата військового квитка · ${name}`
    : "1ПБ Втрата військового квитка";
};

const joinSpaced = (...parts: Array<string | undefined>) =>
  parts
    .map((part) => String(part ?? "").replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .join(" ");

export const declinedPerson = (fields: LostMilitaryIdFields) => {
  const instrumentalName =
    fields.fullNameInstrumentalManual.trim() ||
    toUkrainianInstrumentalFullName(fields.fullName) ||
    fields.fullName;
  return {
    nominative: fields.fullName.trim(),
    instrumental: instrumentalName,
    genitive: toUkrainianGenitiveFullName(fields.fullName) || fields.fullName,
    rankInstrumental:
      toUkrainianInstrumentalRank(fields.rank) || fields.rank.trim(),
    rankGenitive: toUkrainianGenitiveRank(fields.rank) || fields.rank.trim(),
    positionInstrumental:
      toUkrainianInstrumentalPosition(fields.staffPosition) ||
      capitalizeReportPosition(fields.staffPosition),
  };
};

export const investigatorFromPersonnelRow = (
  row: EjournalPreviewRow | null,
) => {
  const summary = buildPersonSummary(row);
  const name = summary.name !== "Особа не вибрана" ? summary.name : "";
  const rawPosition = getPersonFullPositionTitle(row);
  return {
    investigatorFullName: name,
    investigatorRank: summary.rank || "",
    investigatorPosition:
      toUkrainianDativePosition(rawPosition) ||
      capitalizeReportPosition(rawPosition),
    investigatorPersonId: summary.externalId || "",
    investigatorManual: false,
    investigatorDativeManual: "",
  };
};

export const declinedInvestigator = (fields: LostMilitaryIdFields) => {
  const dativeName =
    fields.investigatorDativeManual.trim() ||
    toUkrainianDativeFullName(fields.investigatorFullName) ||
    fields.investigatorFullName;
  return {
    nominative: fields.investigatorFullName.trim(),
    dative: dativeName,
    rankDative:
      toUkrainianDativeRank(fields.investigatorRank) ||
      fields.investigatorRank.trim(),
    position: capitalizeReportPosition(fields.investigatorPosition),
  };
};

export const lossDateText = (fields: LostMilitaryIdFields) => {
  const date = fields.lossDate.trim() || "______";
  return fields.isExactDate ? date : `орієнтовно ${date}`;
};

/** Чи заповнено сценарій «переміщення з → до». */
export const usesMovementCircumstances = (fields: LostMilitaryIdFields) =>
  fields.circumstanceKind === "movement" &&
  Boolean(fields.fromLocation.trim() || fields.toLocation.trim());

export const formatLossLocationPhrase = (location: string) => {
  const text = location.trim();
  if (!text) return "";
  if (/^(с\.|село|м\.|місто|смт|с-ще|п\.|пгт|на\s|у\s|в\s)/iu.test(text)) {
    return text;
  }
  return `у ${text}`;
};

/** Дата + місце + обставини (коли немає переміщення). */
export const buildLossEventCircumstancesPhrase = (fields: LostMilitaryIdFields) => {
  const location = formatLossLocationPhrase(fields.lossLocation);
  const details = fields.customCircumstances.trim();
  if (location && details) {
    const detail =
      details.charAt(0).toLocaleLowerCase("uk-UA") + details.slice(1);
    return `${location}, ${detail}`;
  }
  if (location) return location;
  if (details) return details;
  return "за встановлених обставин";
};

export const buildMovementCircumstancesPhrase = (fields: LostMilitaryIdFields) => {
  const from = fields.fromLocation.trim() || "______";
  const to = fields.toLocation.trim() || "______";
  return `під час переміщення з ${from} до ${to}`;
};

export const circumstancesText = (fields: LostMilitaryIdFields) => {
  if (usesMovementCircumstances(fields)) {
    return buildMovementCircumstancesPhrase(fields);
  }
  return buildLossEventCircumstancesPhrase(fields);
};

const documentLikelyDestroyed = (fields: LostMilitaryIdFields) =>
  /(згор|спал|уничтож|знищ|fpv|фпв|дрон|каб|арт)/iu.test(
    `${fields.customCircumstances} ${fields.lossLocation}`,
  );

export const lossDocumentFateHint = (fields: LostMilitaryIdFields) => {
  if (documentLikelyDestroyed(fields)) {
    return "документ, ймовірно, знищено під час події";
  }
  if (usesMovementCircumstances(fields)) {
    return "документ, ймовірно, загублено під час переміщення або серед особистих речей";
  }
  return "документ, ймовірно, загублено або знищено за вказаних обставин";
};

const RANK_IN_TITLE =
  /(головний майстер-сержант|старший майстер-сержант|майстер-сержант|штаб-сержант|головний сержант|старший сержант|молодший сержант|старший лейтенант|молодший лейтенант|старший солдат|підполковник|полковник|лейтенант|сержант|капітан|майор|солдат)$/iu;

const escapeRegExp = (value: string) =>
  value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

export const reportSignerOf = (fields: LostMilitaryIdFields) =>
  fields.signatories.find((item) => item.blockType === "SIGNER") ??
  fields.signatories[0] ??
  null;

export const approvalSignatoryOf = (fields: LostMilitaryIdFields) =>
  fields.signatories.find((item) => item.blockType === "APPROVAL") ?? null;

const RANK_IN_SIGNATORY_TITLE =
  /(головний майстер-сержант|старший майстер-сержант|майстер-сержант|штаб-сержант|головний сержант|старший сержант|молодший сержант|старший лейтенант|молодший лейтенант|старший солдат|підполковник|полковник|лейтенант|сержант|капітан|майор|солдат)$/iu;

export const isLostMilitaryIdDateLine = (value: string) =>
  /^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value.trim());

export const splitLostMilitaryIdSignatory = (
  signatory: LostMilitaryIdSignatory | null,
) => {
  if (!signatory) {
    return {
      titleLines: [] as string[],
      rank: "",
      fullName: "",
      signatureData: "",
    };
  }
  const lines = signatory.title
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  let rank = signatory.rank.trim();
  const titleLines: string[] = [];
  for (const line of lines) {
    if (isLostMilitaryIdDateLine(line)) continue;
    const tail = line.match(RANK_IN_SIGNATORY_TITLE);
    if (tail && !rank) {
      rank = tail[1];
      const head = line.slice(0, -tail[1].length).trim();
      if (head) titleLines.push(head);
      continue;
    }
    if (
      rank &&
      line.toLocaleLowerCase("uk-UA") === rank.toLocaleLowerCase("uk-UA")
    ) {
      continue;
    }
    titleLines.push(line);
  }
  return {
    titleLines,
    rank,
    fullName: formatNominativeGivenSurname(signatory.fullName),
    signatureData: signatory.signatureData?.trim() ?? "",
  };
};

/** Дата під підписом: день — вручну, місяць і рік — поточні. */
export const buildManualSignatoryDateLine = (value: Date = new Date()) => {
  const monthName = MONTHS_UK[value.getMonth()];
  if (!monthName) return "«  »  ____________  20___ року";
  return `«  »  ${monthName}  ${value.getFullYear()} року`;
};

export const actApprovalDateLine = (_fields?: LostMilitaryIdFields) =>
  buildManualSignatoryDateLine();

const signatoryFooterParts = (signatory: LostMilitaryIdSignatory) => {
  const parts = splitLostMilitaryIdSignatory(signatory);
  const rank = parts.rank || signatory.rank.trim();
  const titleLines =
    parts.titleLines.length > 0
      ? parts.titleLines
      : formatReporterTitleLines(signatory.title, rank);
  return {
    titleLines,
    rank,
    name:
      parts.fullName ||
      formatNominativeGivenSurname(signatory.fullName),
    signatureData: signatory.signatureData?.trim() ?? "",
  };
};

const isUsableCommanderSignatory = (
  signatory: LostMilitaryIdSignatory | null,
) => {
  if (!signatory || isIncompleteLostMilitaryIdSignatory(signatory)) return false;
  const parts = splitLostMilitaryIdSignatory(signatory);
  const rank = (parts.rank || signatory.rank).trim();
  const name = (parts.fullName || signatory.fullName).trim();
  return Boolean(rank && name);
};

export const approvalFooterBlock = (fields: LostMilitaryIdFields) => {
  const approval = approvalSignatoryOf(fields);
  if (approval && !isIncompleteLostMilitaryIdSignatory(approval)) {
    return signatoryFooterParts(approval);
  }
  const unitTitle = `Командир ${normalizeMilitaryUnitPhrase(fields.militaryUnit)}`;
  return {
    titleLines: [unitTitle],
    rank: "",
    name: "",
    signatureData: approval?.signatureData?.trim() ?? "",
  };
};

/** Підпис командира в наказі — APPROVAL, інакше SIGNER, інакше типові дані. */
export const orderFooterBlock = (fields: LostMilitaryIdFields) => {
  const approval = approvalSignatoryOf(fields);
  const signer = reportSignerOf(fields);
  for (const candidate of [approval, signer]) {
    if (isUsableCommanderSignatory(candidate)) {
      return signatoryFooterParts(candidate!);
    }
  }
  return {
    titleLines: formatReporterTitleLines(
      DEFAULT_ORDER_COMMANDER_TITLE,
      DEFAULT_ORDER_COMMANDER_RANK,
    ),
    rank: DEFAULT_ORDER_COMMANDER_RANK,
    name: formatNominativeGivenSurname(DEFAULT_ORDER_COMMANDER_NAME),
    signatureData:
      approval?.signatureData?.trim() ||
      signer?.signatureData?.trim() ||
      "",
  };
};

export const instrumentalInvestigatorLine = (fields: LostMilitaryIdFields) => {
  const nominative = toUkrainianNominativePosition(fields.investigatorPosition);
  const position = toUkrainianInstrumentalPosition(
    nominative || fields.investigatorPosition,
  );
  const rank =
    toUkrainianInstrumentalRank(fields.investigatorRank) ||
    fields.investigatorRank.trim();
  const name =
    toUkrainianInstrumentalFullName(fields.investigatorFullName) ||
    fields.investigatorFullName.trim();
  return joinSpaced(position, rank, name);
};

export const investigatorFooterLines = (fields: LostMilitaryIdFields) => {
  const nominative = toUkrainianNominativePosition(fields.investigatorPosition);
  const text = capitalizeReportPosition(
    nominative || fields.investigatorPosition,
  );
  if (!text) return [] as string[];
  if (text.includes("\n")) {
    return text
      .split(/\n/)
      .map((line) => line.trim())
      .filter(Boolean);
  }
  const blocked = formatPositionTitleBlock(text);
  if (blocked.includes("\n")) {
    return blocked
      .split(/\n/)
      .map((line) => line.trim())
      .filter(Boolean);
  }
  const zSplit = text.match(/^(.*?)\s+(з\s+[а-яіїєґ].*)$/iu);
  if (zSplit) {
    const second = zSplit[2].trim();
    return [
      zSplit[1].trim(),
      second.endsWith(":") ? second : `${second}:`,
    ];
  }
  return [text];
};

export const investigatorFooterBlock = (fields: LostMilitaryIdFields) => ({
  titleLines: investigatorFooterLines(fields),
  rank: fields.investigatorRank.trim(),
  name: formatNominativeGivenSurname(fields.investigatorFullName),
});

export const formatReporterTitleLines = (title: string, rank = "") => {
  const explicit = title
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (explicit.length > 1) {
    return explicit.filter(
      (line) =>
        !isLostMilitaryIdDateLine(line) &&
        line.toLocaleLowerCase("uk-UA") !== rank.trim().toLocaleLowerCase("uk-UA"),
    );
  }
  let text = (explicit[0] || title).replace(/\s+/g, " ").trim();
  if (rank) {
    text = text.replace(new RegExp(`\\s+${escapeRegExp(rank)}$`, "iu"), "").trim();
  } else {
    text = text.replace(RANK_IN_TITLE, "").trim();
  }
  const lines: string[] = [];
  const acting = text.match(
    /^(Тимчасово виконуюч(?:ий|ого)\s+обов['’ʼ]язки)\s+(.*)$/iu,
  );
  let rest = text;
  if (acting) {
    lines.push(acting[1]);
    rest = acting[2].trim();
  }
  const unit = rest.match(/^(.*?)[, ]+(військової частини\s+\S+)$/iu);
  if (unit) {
    if (unit[1].trim()) lines.push(unit[1].trim());
    lines.push(unit[2].trim());
  } else if (rest) {
    lines.push(rest);
  }
  return lines.filter(Boolean);
};

export const extractRankFromTitle = (title: string, fallback = "") => {
  if (fallback.trim()) return fallback.trim();
  return title.match(RANK_IN_TITLE)?.[1] ?? "";
};

export const isLostMilitaryIdSignatoryPlaceholder = (value: string) => {
  const text = value.trim();
  if (!text) return true;
  if (/_{2,}/.test(text)) return true;
  if (/прізвище\s+та\s+ініціали/i.test(text)) return true;
  return false;
};

export const isIncompleteLostMilitaryIdSignatory = (
  signatory: LostMilitaryIdSignatory | null,
) => {
  if (!signatory) return true;
  const parts = splitLostMilitaryIdSignatory(signatory);
  const titleLines = signatory.title
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const titlePlaceholder =
    titleLines.some(isLostMilitaryIdSignatoryPlaceholder) ||
    parts.titleLines.some(isLostMilitaryIdSignatoryPlaceholder);
  const rank = (parts.rank || signatory.rank).trim();
  const name = (parts.fullName || signatory.fullName).trim();
  return (
    titlePlaceholder ||
    isLostMilitaryIdSignatoryPlaceholder(rank) ||
    isLostMilitaryIdSignatoryPlaceholder(name)
  );
};

const fallbackReporterField = (value: string, fallback: string) =>
  isLostMilitaryIdSignatoryPlaceholder(value) ? fallback : value.trim();

const resolveReporterFields = (fields: LostMilitaryIdFields) => {
  const signer = reportSignerOf(fields);
  const signatureData = signer?.signatureData?.trim() ?? "";

  if (signer && !isIncompleteLostMilitaryIdSignatory(signer)) {
    const parts = splitLostMilitaryIdSignatory(signer);
    const rank =
      parts.rank || signer.rank.trim() || fields.reporterRank.trim();
    const titleLines =
      parts.titleLines.length > 0
        ? parts.titleLines
        : formatReporterTitleLines(signer.title, rank);
    const reporterFullName =
      parts.fullName ||
      formatNominativeGivenSurname(signer.fullName) ||
      fields.reporterFullName;
    return {
      reporterTitle: titleLines.join("\n"),
      reporterRank: rank,
      reporterFullName,
      titleLines,
      signatureData,
    };
  }

  const reporterTitle = fallbackReporterField(
    fields.reporterTitle,
    DEFAULT_REPORTER_TITLE,
  );
  const reporterRank = fallbackReporterField(
    fields.reporterRank,
    DEFAULT_REPORTER_RANK,
  );
  const reporterFullName = fallbackReporterField(
    fields.reporterFullName,
    DEFAULT_REPORTER_NAME,
  );
  return {
    reporterTitle,
    reporterRank,
    reporterFullName,
    titleLines: formatReporterTitleLines(reporterTitle, reporterRank),
    signatureData,
  };
};

export const applyReporterFromSignatory = (
  fields: LostMilitaryIdFields,
): LostMilitaryIdFields => {
  const resolved = resolveReporterFields(fields);
  return {
    ...fields,
    reporterTitle: resolved.reporterTitle || fields.reporterTitle,
    reporterRank: resolved.reporterRank || fields.reporterRank,
    reporterFullName: resolved.reporterFullName || fields.reporterFullName,
  };
};

export const reporterHeaderBlock = (fields: LostMilitaryIdFields) => {
  const resolved = resolveReporterFields(fields);
  const genitiveRank =
    toUkrainianGenitiveRank(resolved.reporterRank) || resolved.reporterRank;
  const genitiveName = toUkrainianGenitiveGivenSurname(
    resolved.reporterFullName,
  );
  return [
    ...resolved.titleLines.map((line, index) =>
      index === 0
        ? line
            .replace(/^Тимчасово виконуючий/iu, "Тимчасово виконуючого")
            .replace(/^Командир(?!а)/iu, "Командира")
        : line,
    ),
    joinSpaced(genitiveRank, genitiveName),
  ].filter(Boolean);
};

export const reporterFooterBlock = (fields: LostMilitaryIdFields) => {
  const resolved = resolveReporterFields(fields);
  return {
    titleLines: resolved.titleLines,
    rank: resolved.reporterRank.trim(),
    name: formatNominativeGivenSurname(resolved.reporterFullName),
    signatureData: resolved.signatureData,
  };
};

export const buildLostMilitaryIdReportText = (fields: LostMilitaryIdFields) => {
  if (fields.reportTextOverride.trim()) return fields.reportTextOverride.trim();
  const person = declinedPerson(fields);
  const investigator = declinedInvestigator(fields);
  const unit = normalizeMilitaryUnitPhrase(fields.militaryUnit);
  const search = fields.searchConducted
    ? `Після виявлення факту втрати військового квитка були проведені пошукові заходи, однак ${
        fields.searchResult.trim() || DEFAULT_SEARCH_RESULT
      }.`
    : "Пошукові заходи за фактом втрати військового квитка не проводились.";
  const investigatorLine = investigator.nominative
    ? `Проведення службового розслідування пропоную доручити ${joinSpaced(
        investigator.position,
        investigator.rankDative,
        investigator.dative,
      )}.`
    : "Проведення службового розслідування пропоную доручити командиру підрозділу.";

  return [
    `Доповідаю, що військовослужбовцем ${unit} ${person.rankInstrumental} ${person.instrumental}, ${person.positionInstrumental}, ${lossDateText(fields)}, ${circumstancesText(fields)} було втрачено військовий квиток.`,
    search,
    `У зв’язку з викладеним прошу призначити службове розслідування за фактом втрати військового квитка ${person.rankInstrumental} ${person.instrumental}.`,
    investigatorLine,
  ].join("\n\n");
};

export const buildLostMilitaryIdOrderText = (fields: LostMilitaryIdFields) => {
  if (fields.orderTextOverride.trim()) return fields.orderTextOverride.trim();
  const person = declinedPerson(fields);
  const investigator = declinedInvestigator(fields);
  const unit = normalizeMilitaryUnitPhrase(fields.militaryUnit);
  const assign = investigator.nominative
    ? `Проведення службового розслідування доручити ${joinSpaced(
        investigator.position,
        investigator.rankDative,
        investigator.dative,
      )}.`
    : "Проведення службового розслідування доручити визначеній посадовій особі.";
  return [
    `Призначити службове розслідування за фактом втрати військового квитка військовослужбовцем ${unit} ${person.rankInstrumental} ${person.instrumental}.`,
    assign,
    "Службове розслідування провести відповідно до вимог статті 85 Статуту внутрішньої служби Збройних Сил України та Порядку проведення службового розслідування у Збройних Силах України, затвердженого наказом Міністерства оборони України від 21.11.2017 № 608 (зі змінами).",
    "Матеріали службового розслідування подати на затвердження у встановлений строк.",
  ].join("\n\n");
};

export const reportAuthorLine = (fields: LostMilitaryIdFields) => {
  const resolved = resolveReporterFields(fields);
  return joinSpaced(
    resolved.titleLines.join(" "),
    resolved.reporterRank,
    formatNominativeGivenSurname(resolved.reporterFullName),
  );
};

const formatBirthDateLabel = (value: string) => {
  const text = value.trim();
  if (!text) return "";
  if (/р\.?\s*н\.?/i.test(text)) return `Народився ${text.replace(/\s+/g, " ")}`;
  return `Народився ${text}р.н.`;
};

export const buildLostMilitaryIdPersonExplanation = (
  fields: LostMilitaryIdFields,
) => {
  const person = declinedPerson(fields);
  const date = fields.lossDate.trim() || "______";
  const circumstances = circumstancesText(fields);
  const searchNote = fields.searchConducted
    ? documentLikelyDestroyed(fields)
      ? " Пошуки квитка результату не дали."
      : " Пошуки квитка ні до чого не призвели."
    : "";
  const reportNote = reportAuthorLine(fields)
    ? " Про дану подію командир підрозділу доповів рапортом."
    : "";
  return `${person.rankInstrumental} ${person.nominative} пояснив, що втрата військового квитка сталася ${date} ${circumstances}. Куди міг зникнути документ, військовослужбовець пояснити не може.${searchNote}${reportNote}`;
};

export const buildLostMilitaryIdActAttachments = (
  fields: LostMilitaryIdFields,
) => {
  const person = declinedPerson(fields);
  const unit = normalizeMilitaryUnitPhrase(fields.militaryUnit);
  const unitNumber = militaryUnitLabel(fields.militaryUnit);
  const orderNumber = fields.orderNumber.trim() || "______";
  const orderDate = fields.orderDate.trim() || "____.____.______";
  const reportNumber = fields.reportNumber.trim();
  const reportDate = fields.reportDate.trim() || "____.____.______";
  const reporter = reportAuthorLine(fields);
  const enlistedOrder = fields.enlistedOrder.trim();
  const enlistedDate = fields.enlistedDate.trim() || "____.____.______";

  return [
    `1. Витяг з наказу командира ${unit} (з адміністративно-господарської діяльності) №${orderNumber} від ${orderDate} про призначення службового розслідування за фактом втрати військового квитка ${person.rankGenitive} ${person.genitive}.`,
    reporter
      ? `2. Копію рапорту ${reporter}${
          reportNumber ? ` №${reportNumber}` : ""
        } від ${reportDate} про втрату військового квитка ${person.rankInstrumental} ${person.instrumental}.`
      : `2. Копію рапорту про втрату військового квитка ${person.rankInstrumental} ${person.instrumental} від ${reportDate}.`,
    enlistedOrder
      ? `3. Витяг із наказу командира ${unit} (по стройовій частині) ${enlistedOrder} про зарахування ${person.rankGenitive} ${person.genitive} до ${unit}.`
      : `3. Витяг із наказу командира ${unit} (по стройовій частині) від ${enlistedDate} про зарахування ${person.rankGenitive} ${person.genitive} до ${unitNumber}.`,
    `4. Бланк з отриманими поясненнями ${person.rankGenitive} ${person.genitive}.`,
  ];
};

export const buildLostMilitaryIdActDutySection = (fields: LostMilitaryIdFields) => {
  const person = declinedPerson(fields);
  const date = fields.lossDate.trim() || "______";
  const eventPhrase = usesMovementCircumstances(fields)
    ? buildMovementCircumstancesPhrase(fields).replace(/^під час /iu, "")
    : buildLossEventCircumstancesPhrase(fields);
  const searchClause = fields.searchConducted
    ? documentLikelyDestroyed(fields)
      ? "У поясненні військовослужбовець зазначив, що документ знищено під час події."
      : usesMovementCircumstances(fields)
        ? "У поясненні військовослужбовець зазначив, що документ, ймовірно, загублено під час переміщення."
        : "У поясненні військовослужбовець зазначив, що документ загублено за вказаних обставин."
    : "";
  return [
    `5.1. Дії військовослужбовця:`,
    `Втрата військового квитка ${person.rankInstrumental} ${person.instrumental} відбулася ${eventPhrase}. ${searchClause} ${reportAuthorLine(fields) ? "Доповів рапортом." : ""}`.trim(),
    `Зв'язок правопорушення з виконанням військовослужбовцем обов'язків військової служби:`,
    `Військовослужбовець сумлінно виконує свої службові обов'язки.`,
    `${date} втрата військового квитка ${person.rankInstrumental} ${person.instrumental} відбулася внаслідок особистої недбалості.`,
    `Заперечення, заяви та клопотання особи, стосовно якої проведено службове розслідування, мотиви їх відхилення чи підстави для задоволення: не надходило.`,
  ];
};
export const buildLostMilitaryIdActCircumstances = (
  fields: LostMilitaryIdFields,
) => {
  if (fields.actCircumstancesOverride.trim()) {
    return fields.actCircumstancesOverride.trim();
  }
  const person = declinedPerson(fields);
  const longDate = formatUaLongDate(fields.lossDate) || lossDateText(fields);
  const reporter = reportAuthorLine(fields);
  return [
    `Службове розслідування проводиться за фактом втрати військового квитка військовослужбовцем ${person.instrumental}.`,
    `${longDate} ${person.rankInstrumental} ${person.nominative} ${circumstancesText(fields)} втратив військовий квиток. Зі слів військовослужбовця, ${lossDocumentFateHint(fields)}.`,
    fields.searchConducted
      ? `Пошук військового квитка в районі розташування підрозділу та серед особистих речей результату не дав: ${
          fields.searchResult.trim() || DEFAULT_SEARCH_RESULT
        }.`
      : "Пошукові заходи не проводились.",
    reporter
      ? `Про втрату документа ${reporter} доповів рапортом${
          fields.reportNumber.trim() ? ` №${fields.reportNumber.trim()}` : ""
        }${fields.reportDate.trim() ? ` від ${fields.reportDate.trim()}` : ""}.`
      : `Про втрату документа доповідено рапортом${
          fields.reportDate.trim() ? ` від ${fields.reportDate.trim()}` : ""
        }.`,
  ].join("\n\n");
};

export const buildLostMilitaryIdPersonCard = (fields: LostMilitaryIdFields) => {
  const person = declinedPerson(fields);
  const unit = normalizeMilitaryUnitPhrase(fields.militaryUnit);
  const unitNumber = militaryUnitLabel(fields.militaryUnit);
  const birth = formatBirthDateLabel(fields.birthDate);
  const personal = [
    birth,
    fields.citizenship.trim() || "громадянин України",
    fields.education.trim() ? `Освіта: ${fields.education.trim()}` : "",
    fields.maritalStatus.trim() || "",
    fields.address.trim() ? `адреса реєстрації: ${fields.address.trim()}` : "",
  ]
    .filter(Boolean)
    .join(", ");
  const enlisted = fields.enlistedDate.trim()
    ? `В ЗСУ з ${fields.enlistedDate.trim()}. До списків частини ${unitNumber} зарахований ${fields.enlistedDate.trim()}${
        fields.enlistedOrder.trim()
          ? ` відповідно до ${fields.enlistedOrder.trim()}`
          : ""
      }.`
    : "";
  return [
    `${person.nominative}, ${fields.rank.trim() || "______"}, ${
      capitalizeReportPosition(fields.staffPosition) || "______"
    } ${unit}.`,
    enlisted,
    personal,
  ]
    .filter(Boolean)
    .join(" ");
};

export const buildLostMilitaryIdActConclusions = (
  fields: LostMilitaryIdFields,
) => {
  if (fields.actConclusionsOverride.trim()) {
    return fields.actConclusionsOverride.trim();
  }
  const person = declinedPerson(fields);
  return `${person.rankInstrumental} ${person.nominative} діяв згідно своїх службових обов’язків та статуту ЗСУ. За допущену внаслідок недбалого ставлення військовослужбовцем до зберігання службових документів, що є порушенням вимог статті 11 Дисциплінарного статуту Збройних Сил України, ${person.rankInstrumental} ${person.nominative} заслуговує на притягнення до дисциплінарної відповідальності. Ознак умисних дій або корисливих мотивів не встановлено.`;
};

export const buildLostMilitaryIdActProposals = (fields: LostMilitaryIdFields) => {
  if (fields.actProposalsOverride.trim()) {
    return fields.actProposalsOverride.trim();
  }
  const person = declinedPerson(fields);
  const unit = normalizeMilitaryUnitPhrase(fields.militaryUnit);
  const order =
    fields.orderNumber.trim() && fields.orderDate.trim()
      ? `наказу командира ${unit} від ${fields.orderDate} №${fields.orderNumber}`
      : `наказу командира ${unit} про призначення службового розслідування`;
  const chief = fields.personnelChiefName.trim() || "начальнику відділення персоналу та стройового штабу";
  return [
    `1. Розслідування стосовно ${order} вважати завершеним.`,
    `2. З урахуванням матеріалів службового розслідування та частини 1 пункту 6 Положення про військовий квиток осіб рядового, сержантського і старшинського складу, затвердженого Указом Президента України від 30 грудня 2016 року № 582/2016, за допущену втрату військового квитка ${person.rankGenitive} ${person.genitive} притягнути до дисциплінарної відповідальності.`,
    `3. ${chief} ${unit} забезпечити направлення у встановленому порядку документів для відновлення військового квитка ${person.rankGenitive} ${person.genitive}.`,
  ].join("\n");
};

const pickPersonnel = (row: EjournalPreviewRow | null, parts: string[]) =>
  getPersonFieldValue(row, parts).trim();

export const createLostMilitaryIdFields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
  signatories: LostMilitaryIdSignatory[] = [],
): LostMilitaryIdFields => {
  const fullName = summary.name !== "Особа не вибрана" ? summary.name : "";
  const staffPosition = getPersonFullPositionTitle(row);
  return applyReporterFromSignatory({
    addressee: DEFAULT_ADDRESSEE,
    militaryUnit: DEFAULT_UNIT,
    fullName,
    rank: summary.rank || "",
    staffPosition: capitalizeReportPosition(staffPosition),
    unitLabel: pickPersonnel(row, ["підрозділ"]) || "",
    lossDate: "",
    isExactDate: true,
    circumstanceKind: "custom",
    lossLocation: "",
    fromLocation: "",
    toLocation: "",
    customCircumstances: "",
    searchConducted: false,
    searchResult: DEFAULT_SEARCH_RESULT,
    reporterTitle: DEFAULT_REPORTER_TITLE,
    reporterRank: DEFAULT_REPORTER_RANK,
    reporterFullName: DEFAULT_REPORTER_NAME,
    investigatorFullName: "",
    investigatorRank: "",
    investigatorPosition: "",
    investigatorPersonId: "",
    investigatorManual: false,
    reportDate: formatUaDate(new Date()),
    orderNumber: "",
    orderDate: "",
    approvalDate: "",
    reportNumber: "",
    personnelChiefName: "",
    birthDate: summary.birthDate || pickPersonnel(row, ["дата_народження"]),
    enlistedDate:
      pickPersonnel(row, ["дата_зарахування"]) ||
      pickPersonnel(row, ["зарахування", "списків"]),
    enlistedOrder: [
      pickPersonnel(row, ["наказ_про_зарахування", "номер"]),
      pickPersonnel(row, ["наказ_про_зарахування", "дата"]),
    ]
      .filter(Boolean)
      .join(" від "),
    education: pickPersonnel(row, ["освіта"]),
    maritalStatus: pickPersonnel(row, ["сімейн"]) || pickPersonnel(row, ["одруж"]),
    address:
      pickPersonnel(row, ["адреса", "реєстр"]) ||
      pickPersonnel(row, ["зареєстров"]) ||
      pickPersonnel(row, ["адреса"]),
    citizenship: "громадянин України",
    folderName: buildLostMilitaryIdFolderName(fullName),
    statusNote: "",
    editManually: false,
    fullNameInstrumentalManual: "",
    investigatorDativeManual: "",
    reportTextOverride: "",
    orderTextOverride: "",
    actCircumstancesOverride: "",
    actConclusionsOverride: "",
    actProposalsOverride: "",
    signatories,
  });
};

export const mergeLostMilitaryIdFields = (
  defaults: LostMilitaryIdFields,
  value: unknown,
): LostMilitaryIdFields => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return defaults;
  }
  const next = value as Partial<LostMilitaryIdFields>;
  const merged: LostMilitaryIdFields = {
    ...defaults,
    ...next,
    isExactDate: Boolean(next.isExactDate ?? defaults.isExactDate),
    searchConducted: Boolean(next.searchConducted ?? defaults.searchConducted),
    editManually: Boolean(next.editManually ?? defaults.editManually),
    investigatorManual:
      typeof next.investigatorManual === "boolean"
        ? next.investigatorManual
        : Boolean(
            String(next.investigatorFullName ?? "").trim() &&
              !String(next.investigatorPersonId ?? "").trim(),
          ),
    investigatorPersonId: String(
      next.investigatorPersonId ?? defaults.investigatorPersonId ?? "",
    ),
    circumstanceKind:
      next.circumstanceKind === "custom" || next.circumstanceKind === "movement"
        ? next.circumstanceKind
        : defaults.circumstanceKind,
  };
  const fullName =
    String(merged.fullName ?? "").trim() || String(defaults.fullName ?? "").trim();
  const currentFolder = String(merged.folderName ?? "").trim();
  return applyReporterFromSignatory({
    ...merged,
    signatories: defaults.signatories.length
      ? defaults.signatories
      : Array.isArray(merged.signatories)
        ? merged.signatories
        : [],
    fullName,
    staffPosition: capitalizeReportPosition(merged.staffPosition),
    investigatorPosition: (() => {
      const text = String(merged.investigatorPosition ?? "").trim();
      if (!text || looksLikeUaDateToken(text.replace(/у$/iu, ""))) return "";
      return capitalizeReportPosition(text);
    })(),
    folderName:
      !currentFolder ||
      (/втрата військового квитка/i.test(currentFolder) &&
        fullName &&
        !currentFolder.includes(fullName))
        ? buildLostMilitaryIdFolderName(fullName)
        : currentFolder,
  });
};
