import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  getPersonFullPositionTitle,
} from "../personnel/personnelUtils";
import { toUkrainianDativeRank } from "./form12Report";
import { capitalizeReportPosition } from "./reportPosition";
import { splitFullNameParts } from "./serviceCharacteristicReport";

const SELECTED_PERSON_FULL_POSITION_KEY =
  "army-grid:selected-person-full-position";

export const readStoredSelectedPersonFullPosition = () => {
  if (typeof window === "undefined") return "";
  return (window.localStorage.getItem(SELECTED_PERSON_FULL_POSITION_KEY) || "")
    .trim();
};

export const storeSelectedPersonFullPosition = (value: string) => {
  if (typeof window === "undefined") return;
  const text = value.trim();
  if (text) {
    window.localStorage.setItem(SELECTED_PERSON_FULL_POSITION_KEY, text);
  } else {
    window.localStorage.removeItem(SELECTED_PERSON_FULL_POSITION_KEY);
  }
};

/** Витягнути повну посаду з `profile.roster.row.values` або з самого row. */
export const getFullPositionFromPersonnelProfileRoster = (
  roster: { row?: Record<string, unknown> | null } | null | undefined,
) => {
  if (!roster?.row || typeof roster.row !== "object") return "";
  const values = roster.row.values;
  if (values && typeof values === "object" && !Array.isArray(values)) {
    return getPersonFullPositionTitle(values as EjournalPreviewRow);
  }
  return getPersonFullPositionTitle(roster.row as EjournalPreviewRow);
};

export type ZhbdCertificateSignatory = {
  blockType: "SIGNER" | "APPROVAL";
  title: string;
  rank: string;
  fullName: string;
  signatureData?: string | null;
};

export type ZhbdCertificateFields = {
  rank: string;
  fullName: string;
  staffPosition: string;
  /** Реальна повна посада з ранкового (для підказки в UI, не в текст довідки). */
  actualFullPosition: string;
  periodFrom: string;
  periodTo: string;
  periodNote: string;
  bodyParagraph: string;
  basisOrders: string;
  basisJournal: string;
  headerDate: string;
  documentNumber: string;
  date: string;
  folderName: string;
  signatories: ZhbdCertificateSignatory[];
  statusNote: string;
};

/** Same progress as службова характеристика. */
export const zhbdCertificateWorkflowSteps = [
  { key: "created", title: "Створили" },
  { key: "document", title: "Заповнили" },
  { key: "sent", title: "Відправлено" },
  { key: "confirmed", title: "Підтверджено" },
  { key: "forCharacteristic", title: "Для характеристики" },
];

const DEFAULT_BASIS_ORDERS =
  "Бойові розпорядження командира військової частини А4862 від 01.01.2026. № 4862/6/01/001/дск, від 01.04.2026 № 4862/829/04/089/дск.";

const DEFAULT_BASIS_JOURNAL =
  "Журнал ведення бойових дій військової частини А4862 (інв. № 171/дск. від 01.01.2026 року)";

const formatUaDate = (value: Date) => {
  const pad = (part: number) => String(part).padStart(2, "0");
  return `${pad(value.getDate())}.${pad(value.getMonth() + 1)}.${value.getFullYear()}`;
};

const formatHeaderDate = (value: Date) => {
  const pad = (part: number) => String(part).padStart(2, "0");
  return `“${pad(value.getDate())}” ${pad(value.getMonth() + 1)}.${value.getFullYear()}р.`;
};

export const formatZhbdFullName = (fullName: string) => {
  const { lastName, firstName, patronymic } = splitFullNameParts(fullName);
  return [lastName, firstName, patronymic].filter(Boolean).join(" ");
};

/** Звання сержантського складу (для довідок ЖБД після тестування). */
export const isZhbdSergeantRank = (rank: string) => {
  const text = rank.trim().toLocaleLowerCase("uk-UA");
  if (!text) return false;
  if (/(лейтенант|капітан|майор|підполковник|полковник|генерал)/i.test(text)) {
    return false;
  }
  return /(сержант|старшина)/i.test(text);
};

/**
 * Посада на момент БЗ для сержантів:
 * — якщо в штаті вже «командир … відділення/взводу» — лишаємо повний текст з загального списку;
 * — інакше (медик, оператор тощо) — «командир відділення» або «командир взводу».
 */
export const resolveZhbdCombatStaffPosition = (
  rank: string,
  staffPosition: string,
) => {
  const raw = staffPosition.trim();
  if (!raw || !isZhbdSergeantRank(rank)) return raw;
  const pos = raw.toLocaleLowerCase("uk-UA").replace(/\s+/g, " ");
  // Уже командирська посада — не скорочуємо (потрібен повний текст у довідці).
  if (/командир/.test(pos) && /(відділен|взвод)/.test(pos)) {
    return raw;
  }
  const isPlatoonCommander =
    /командир(?:а)?(?:\s+\d+)?(?:\s+[а-яіїєґ\-]+)*\s+взвод/.test(pos) &&
    !/відділен/.test(pos);
  return isPlatoonCommander ? "командир взводу" : "командир відділення";
};

export const buildZhbdBodyParagraph = (fields: {
  rank: string;
  fullName: string;
  staffPosition: string;
  periodFrom: string;
  periodTo: string;
  periodNote: string;
}) => {
  const dativeRank =
    toUkrainianDativeRank(fields.rank) || fields.rank.trim() || "______";
  const name = formatZhbdFullName(fields.fullName) || "______";
  // Беремо текст поля як є — без повторного скорочення.
  const position = fields.staffPosition.trim() || "______";
  const from = fields.periodFrom.trim() || "______";
  const to = fields.periodTo.trim() || "теперішній час";
  const note = fields.periodNote.trim();
  const period = note
    ? `в період з ${from} по ${to} (${note})`
    : `в період з ${from} по ${to}`;
  return `Надана ${dativeRank} ${name}, в тому що виконував (виконує) бойові (спеціальні) завдання на посаді ${position} ${period}.`;
};

export const createZhbdCertificateFields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
  signatories: ZhbdCertificateSignatory[] = [],
): ZhbdCertificateFields => {
  const fullName = summary.name !== "Особа не вибрана" ? summary.name : "";
  const rank = summary.rank || "";
  const rawStaffPosition =
    getPersonFullPositionTitle(row) || readStoredSelectedPersonFullPosition();
  const staffPosition = capitalizeReportPosition(
    resolveZhbdCombatStaffPosition(rank, rawStaffPosition),
  );
  const now = new Date();
  const yearStart = `01.01.${now.getFullYear()}`;
  const base = {
    rank,
    fullName: formatZhbdFullName(fullName),
    staffPosition,
    actualFullPosition: rawStaffPosition,
    periodFrom: yearStart,
    periodTo: "теперішній час",
    periodNote: "",
    basisOrders: DEFAULT_BASIS_ORDERS,
    basisJournal: DEFAULT_BASIS_JOURNAL,
    headerDate: formatHeaderDate(now),
    documentNumber: "№ ______",
    date: formatUaDate(now),
    folderName: fullName ? `Довідка ЖБД · ${fullName}` : "Довідка ЖБД",
    signatories,
    statusNote: "",
  };
  return {
    ...base,
    bodyParagraph: buildZhbdBodyParagraph(base),
  };
};

export const mergeZhbdCertificateFields = (
  defaults: ZhbdCertificateFields,
  value: unknown,
): ZhbdCertificateFields => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return defaults;
  }
  const next = value as Partial<ZhbdCertificateFields>;
  const merged = {
    ...defaults,
    ...next,
    signatories: Array.isArray(next.signatories)
      ? next.signatories
      : defaults.signatories,
  };
  const pick = (personnel: string, document: string) =>
    String(personnel ?? "").trim() || String(document ?? "").trim();

  const rank = pick(defaults.rank, merged.rank);
  const fullName = pick(defaults.fullName, merged.fullName);
  const actualFullPosition = pick(
    defaults.actualFullPosition,
    String(merged.actualFullPosition ?? ""),
  );
  const staffPosition = capitalizeReportPosition(
    resolveZhbdCombatStaffPosition(
      rank,
      actualFullPosition || pick(defaults.staffPosition, merged.staffPosition),
    ),
  );
  const periodFrom = pick(defaults.periodFrom, merged.periodFrom);
  const periodTo = pick(defaults.periodTo, merged.periodTo);
  const periodNote = String(merged.periodNote ?? defaults.periodNote ?? "");

  const savedBody = String(merged.bodyParagraph ?? "").trim();
  const expectedBody = buildZhbdBodyParagraph({
    rank,
    fullName,
    staffPosition,
    periodFrom,
    periodTo,
    periodNote,
  });
  const shouldRefreshBody =
    !savedBody ||
    /АЛЄКСЄЄВ\s+Дмитро\s+Юрійович/i.test(savedBody) ||
    savedBody === defaults.bodyParagraph ||
    // Старе скорочення «командир відділення/взводу», коли вже є повна посада
    (Boolean(staffPosition) &&
      /на посаді\s+командир\s+(відділення|взводу)\b/i.test(savedBody) &&
      !new RegExp(
        `на посаді\\s+${staffPosition.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`,
        "i",
      ).test(savedBody));

  const refreshed = {
    ...merged,
    rank,
    fullName,
    staffPosition,
    actualFullPosition,
    periodFrom,
    periodTo,
    periodNote,
  };

  return {
    ...refreshed,
    bodyParagraph: shouldRefreshBody ? expectedBody : savedBody,
  };
};
