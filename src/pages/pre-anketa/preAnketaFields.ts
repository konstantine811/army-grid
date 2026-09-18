export type PreAnketaFieldKey =
  | "callsign"
  | "fullName"
  | "rank"
  | "rnokpp"
  | "iban"
  | "passport"
  | "driverLicense"
  | "birthDate"
  | "birthPlace"
  | "registrationAddress"
  | "actualAddress"
  | "phone1"
  | "phone2"
  | "education"
  | "work"
  | "serviceContract"
  | "serviceMobilized"
  | "servedBefore2022"
  | "arrivedFromUnit"
  | "conscription"
  | "ubdCertificate"
  | "criminalCharged"
  | "criminalNotCharged"
  | "familyStatus"
  | "mother"
  | "father"
  | "trustedPerson"
  | "children"
  | "sports"
  | "sportsRank"
  | "signatureDate";

export type PreAnketaFormFields = Record<PreAnketaFieldKey, string>;

export type PreAnketaFieldDef = {
  key: PreAnketaFieldKey;
  label: string;
  group: string;
  multiline?: boolean;
  checkbox?: boolean;
};

export const PRE_ANKETA_FIELD_DEFS: PreAnketaFieldDef[] = [
  { key: "callsign", label: "Позивний", group: "Основне" },
  { key: "fullName", label: "П.І.Б.", group: "Основне" },
  { key: "rank", label: "Військове звання", group: "Основне" },
  { key: "rnokpp", label: "Ідентифікаційний (код) номер", group: "Основне" },
  { key: "iban", label: "IBAN", group: "Основне" },
  { key: "passport", label: "Паспорт (серія, номер, коли і ким виданий)", group: "Документи", multiline: true },
  { key: "driverLicense", label: "Посвідчення водія (категорія)", group: "Документи" },
  { key: "birthDate", label: "Дата народження", group: "Особисті дані" },
  { key: "birthPlace", label: "Місце народження", group: "Особисті дані", multiline: true },
  { key: "registrationAddress", label: "Адреса (прописка)", group: "Адреси", multiline: true },
  { key: "actualAddress", label: "Адреса фактичного проживання", group: "Адреси", multiline: true },
  { key: "phone1", label: "Телефон 1", group: "Контакти" },
  { key: "phone2", label: "Телефон 2", group: "Контакти" },
  { key: "education", label: "Освіта", group: "Особисті дані", multiline: true },
  { key: "work", label: "Основне місце роботи", group: "Особисті дані", multiline: true },
  { key: "serviceContract", label: "Контрактник", group: "Служба", checkbox: true },
  { key: "serviceMobilized", label: "Мобілізований", group: "Служба", checkbox: true },
  { key: "servedBefore2022", label: "Чи служив у ЗС до 24.02.2022", group: "Служба" },
  { key: "arrivedFromUnit", label: "З якої військової частини прибув", group: "Служба", multiline: true },
  { key: "conscription", label: "Яким РТЦК та СП мобілізований та коли", group: "Служба", multiline: true },
  { key: "ubdCertificate", label: "Посвідчення УБД (серія, номер)", group: "Документи" },
  { key: "criminalCharged", label: "Притягався до кримінальної відповідальності", group: "Інше", checkbox: true },
  { key: "criminalNotCharged", label: "Не притягався до кримінальної відповідальності", group: "Інше", checkbox: true },
  { key: "familyStatus", label: "Сімейний стан", group: "Родичі", multiline: true },
  { key: "mother", label: "Мати", group: "Родичі", multiline: true },
  { key: "father", label: "Батько", group: "Родичі", multiline: true },
  { key: "trustedPerson", label: "Довірена особа", group: "Родичі", multiline: true },
  { key: "children", label: "Діти", group: "Родичі", multiline: true },
  { key: "sports", label: "Спорт", group: "Інше", multiline: true },
  { key: "sportsRank", label: "Спортивний розряд", group: "Інше" },
  { key: "signatureDate", label: "Дата підпису", group: "Підпис" },
];

export const createEmptyPreAnketaForm = (): PreAnketaFormFields =>
  Object.fromEntries(
    PRE_ANKETA_FIELD_DEFS.map((field) => [field.key, ""]),
  ) as PreAnketaFormFields;

export const preAnketaFieldLabel = (key: PreAnketaFieldKey) =>
  PRE_ANKETA_FIELD_DEFS.find((field) => field.key === key)?.label ?? key;

export const isPreAnketaCheckboxField = (key: PreAnketaFieldKey) =>
  PRE_ANKETA_FIELD_DEFS.some((field) => field.key === key && field.checkbox);

const sanitizePreAnketaFileNamePart = (value: string) =>
  value.trim().replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim();

export const formatPreAnketaDownloadName = (fields: PreAnketaFormFields) => {
  const parts = [
    sanitizePreAnketaFileNamePart(fields.fullName),
    sanitizePreAnketaFileNamePart(fields.callsign),
  ].filter(Boolean);
  const base = parts.length > 0 ? parts.join(" ") : "Анкета";
  return `${base}.docx`;
};
