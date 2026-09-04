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
} from "./lostMilitaryIdCases";
import { toUkrainianDativePosition } from "./form12Report";
import { capitalizeReportPosition } from "./reportPosition";

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
  reportNumber: string;
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
  const position = getPersonFullPositionTitle(row);
  return {
    investigatorFullName: name,
    investigatorRank: summary.rank || "",
    investigatorPosition:
      toUkrainianDativePosition(position) || capitalizeReportPosition(position),
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

export const circumstancesText = (fields: LostMilitaryIdFields) => {
  if (fields.circumstanceKind === "custom") {
    return fields.customCircumstances.trim() || "______";
  }
  const from = fields.fromLocation.trim() || "______";
  const to = fields.toLocation.trim() || "______";
  return `під час переміщення з ${from} до ${to}`;
};

const RANK_IN_TITLE =
  /(головний майстер-сержант|старший майстер-сержант|майстер-сержант|штаб-сержант|головний сержант|старший сержант|молодший сержант|старший лейтенант|молодший лейтенант|старший солдат|підполковник|полковник|лейтенант|сержант|капітан|майор|солдат)$/iu;

const escapeRegExp = (value: string) =>
  value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

export const reportSignerOf = (fields: LostMilitaryIdFields) =>
  fields.signatories.find((item) => item.blockType === "SIGNER") ??
  fields.signatories[0] ??
  null;

export const formatReporterTitleLines = (title: string, rank = "") => {
  const explicit = title
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (explicit.length > 1) {
    return explicit.filter(
      (line) =>
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

export const applyReporterFromSignatory = (
  fields: LostMilitaryIdFields,
): LostMilitaryIdFields => {
  const signer = reportSignerOf(fields);
  if (!signer) return fields;
  const rank = extractRankFromTitle(signer.title, signer.rank);
  const titleLines = formatReporterTitleLines(signer.title, rank);
  return {
    ...fields,
    reporterTitle: titleLines.join("\n") || fields.reporterTitle,
    reporterRank: rank || fields.reporterRank,
    reporterFullName: signer.fullName.trim() || fields.reporterFullName,
  };
};

export const reporterHeaderBlock = (fields: LostMilitaryIdFields) => {
  const applied = applyReporterFromSignatory(fields);
  const lines = formatReporterTitleLines(
    applied.reporterTitle,
    applied.reporterRank,
  );
  const genitiveRank =
    toUkrainianGenitiveRank(applied.reporterRank) || applied.reporterRank;
  const genitiveName = toUkrainianGenitiveGivenSurname(applied.reporterFullName);
  return [
    ...lines.map((line, index) =>
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
  const applied = applyReporterFromSignatory(fields);
  const signer = reportSignerOf(applied);
  return {
    titleLines: formatReporterTitleLines(
      applied.reporterTitle,
      applied.reporterRank,
    ),
    rank: applied.reporterRank.trim(),
    name: formatNominativeGivenSurname(applied.reporterFullName),
    signatureData: signer?.signatureData || "",
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

export const buildLostMilitaryIdActCircumstances = (
  fields: LostMilitaryIdFields,
) => {
  if (fields.actCircumstancesOverride.trim()) {
    return fields.actCircumstancesOverride.trim();
  }
  const person = declinedPerson(fields);
  const unit = normalizeMilitaryUnitPhrase(fields.militaryUnit);
  const longDate = formatUaLongDate(fields.lossDate) || lossDateText(fields);
  return [
    `Службове розслідування проводиться за фактом втрати військового квитка військовослужбовцем ${person.instrumental}.`,
    `${longDate} ${person.rankInstrumental} ${person.nominative} ${circumstancesText(fields)} втратив військовий квиток.`,
    fields.searchConducted
      ? `Пошук військового квитка в районі розташування підрозділу та серед особистих речей результату не дав: ${
          fields.searchResult.trim() || DEFAULT_SEARCH_RESULT
        }.`
      : "Пошукові заходи не проводились.",
    `Про втрату документа командир 1 піхотного батальйону ${unit} доповів рапортом${
      fields.reportDate.trim() ? ` від ${fields.reportDate.trim()}` : ""
    }.`,
  ].join("\n\n");
};

export const buildLostMilitaryIdPersonCard = (fields: LostMilitaryIdFields) => {
  const person = declinedPerson(fields);
  const unit = normalizeMilitaryUnitPhrase(fields.militaryUnit);
  const unitNumber = militaryUnitLabel(fields.militaryUnit);
  return [
    `${person.nominative}, ${fields.rank.trim() || "______"}, ${
      capitalizeReportPosition(fields.staffPosition) || "______"
    } ${unit}.`,
    fields.enlistedDate.trim()
      ? `В ЗСУ з ${fields.enlistedDate}. До списків частини ${unitNumber} зарахований ${fields.enlistedDate}${
          fields.enlistedOrder.trim()
            ? ` відповідно до ${fields.enlistedOrder}`
            : ""
        }.`
      : "",
    [
      fields.birthDate.trim() ? `Народився ${fields.birthDate}` : "",
      fields.citizenship.trim() || "громадянин України",
      fields.education.trim() ? `освіта: ${fields.education}` : "",
      fields.maritalStatus.trim() || "",
      fields.address.trim() ? `адреса реєстрації: ${fields.address}` : "",
    ]
      .filter(Boolean)
      .join(", "),
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
  return [
    `1. Розслідування стосовно ${order} вважати завершеним.`,
    `2. З урахуванням матеріалів службового розслідування та частини 1 пункту 6 Положення про військовий квиток осіб рядового, сержантського і старшинського складу, затвердженого Указом Президента України від 30 грудня 2016 року № 582/2016, за допущену втрату військового квитка ${person.rankGenitive} ${person.genitive} притягнути до дисциплінарної відповідальності.`,
    `3. Начальнику відділення персоналу та стройового штабу ${unit} забезпечити направлення у встановленому порядку документів для відновлення військового квитка ${person.rankGenitive} ${person.genitive}.`,
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
    isExactDate: false,
    circumstanceKind: "movement",
    fromLocation: "",
    toLocation: "",
    customCircumstances: "",
    searchConducted: true,
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
    reportNumber: "",
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
    investigatorPosition: capitalizeReportPosition(merged.investigatorPosition),
    folderName:
      !currentFolder ||
      (/втрата військового квитка/i.test(currentFolder) &&
        fullName &&
        !currentFolder.includes(fullName))
        ? buildLostMilitaryIdFolderName(fullName)
        : currentFolder,
  });
};
