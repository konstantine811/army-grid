import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  getPersonFieldValue,
} from "../personnel/personnelUtils";

export type ServiceCharacteristicSignatory = {
  blockType: "SIGNER" | "APPROVAL";
  title: string;
  rank: string;
  fullName: string;
  signatureData?: string | null;
};

export type ServiceCharacteristicFields = {
  rank: string;
  lastName: string;
  firstName: string;
  patronymic: string;
  staffPosition: string;
  introParagraph: string;
  professionalParagraph: string;
  combatParagraph: string;
  moralParagraph: string;
  drillParagraph: string;
  conclusion: string;
  date: string;
  folderName: string;
  signatories: ServiceCharacteristicSignatory[];
  statusNote: string;
};

export const serviceCharacteristicWorkflowSteps = [
  { key: "created", title: "Створили" },
  { key: "document", title: "Заповнили" },
  { key: "sent", title: "Відправлено" },
  { key: "confirmed", title: "Підтверджено" },
  { key: "forCharacteristic", title: "Для характеристики" },
];

const DEFAULT_CONCLUSION = "Займаній посаді відповідає.";

const DEFAULT_PROFESSIONAL =
  "У професійному відношенні підготовлений добре, постійно працює над підвищенням свого професійного рівня. До виконання службових обов’язків ставиться сумлінно та відповідально. Поставлені завдання виконує якісно, точно й у встановлені терміни, проявляючи при цьому розумну ініціативу.";

const DEFAULT_COMBAT =
  "У складних умовах діє впевнено, здатний швидко орієнтуватися в обстановці та приймати обґрунтовані й виважені рішення. Неухильно дотримується вимог військової дисципліни, правил військової ввічливості та носіння військової форми одягу.";

const DEFAULT_MORAL =
  "Морально стійкий, урівноважений та вимогливий до себе. Власну діяльність оцінює критично, на зауваження реагує адекватно, своєчасно усуває виявлені недоліки. У колективі користується заслуженим авторитетом, підтримує доброзичливі та ділові взаємовідносини.";

const DEFAULT_DRILL =
  "У стройовому відношенні підтягнутий, фізично добре розвинений, має високу працездатність. Свої знання, сили та старання спрямовує на сумлінне виконання військового обов’язку, зміцнення обороноздатності України та розбудову Збройних Сил України. До адміністративної відповідальності за вчинення корупційного або військового адміністративного правопорушення не притягувався.";

const formatUaDate = (value: Date) => {
  const pad = (part: number) => String(part).padStart(2, "0");
  return `${pad(value.getDate())}.${pad(value.getMonth() + 1)}.${value.getFullYear()}`;
};

export const splitFullNameParts = (fullName: string) => {
  const parts = fullName
    .trim()
    .replace(/\s*\([^)]+\)\s*$/, "")
    .split(/\s+/)
    .filter(Boolean);
  return {
    lastName: (parts[0] ?? "").toLocaleUpperCase("uk-UA"),
    firstName: parts[1] ?? "",
    patronymic: parts.slice(2).join(" "),
  };
};

export const buildIntroParagraph = (
  rank: string,
  firstName: string,
  lastName: string,
) => {
  const who = [rank.trim() || "військовослужбовець", firstName.trim(), lastName.trim()]
    .filter(Boolean)
    .join(" ");
  return `За час проходження служби на займаній посаді ${who} зарекомендував себе як дисциплінований, відповідальний та грамотний військовослужбовець.`;
};

export const createServiceCharacteristicFields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
  signatories: ServiceCharacteristicSignatory[] = [],
): ServiceCharacteristicFields => {
  const fullName = summary.name !== "Особа не вибрана" ? summary.name : "";
  const { lastName, firstName, patronymic } = splitFullNameParts(fullName);
  const rank = summary.rank || "";
  const staffPosition =
    getPersonFieldValue(row, ["посада"]) ||
    getPersonFieldValue(row, ["штатна", "посада"]) ||
    getPersonFieldValue(row, ["чим", "займається"]) ||
    "";

  return {
    rank,
    lastName,
    firstName,
    patronymic,
    staffPosition,
    introParagraph: buildIntroParagraph(rank, firstName, lastName),
    professionalParagraph: DEFAULT_PROFESSIONAL,
    combatParagraph: DEFAULT_COMBAT,
    moralParagraph: DEFAULT_MORAL,
    drillParagraph: DEFAULT_DRILL,
    conclusion: DEFAULT_CONCLUSION,
    date: formatUaDate(new Date()),
    folderName: fullName
      ? `Службова характеристика · ${fullName}`
      : "Службова характеристика",
    signatories,
    statusNote: "",
  };
};

export const mergeServiceCharacteristicFields = (
  defaults: ServiceCharacteristicFields,
  value: unknown,
): ServiceCharacteristicFields => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return defaults;
  }
  const next = value as Partial<ServiceCharacteristicFields>;
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
  const lastName = pick(defaults.lastName, merged.lastName);
  const firstName = pick(defaults.firstName, merged.firstName);
  const patronymic = pick(defaults.patronymic, merged.patronymic);

  const savedIntro = String(merged.introParagraph ?? "").trim();
  const shouldRefreshIntro =
    !savedIntro ||
    /солдат\s+Євген\s+БРИЛЬ/i.test(savedIntro) ||
    savedIntro === defaults.introParagraph;

  return {
    ...merged,
    rank,
    lastName,
    firstName,
    patronymic,
    staffPosition: pick(defaults.staffPosition, merged.staffPosition),
    introParagraph: shouldRefreshIntro
      ? buildIntroParagraph(rank, firstName, lastName)
      : savedIntro,
  };
};
