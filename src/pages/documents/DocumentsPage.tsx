import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Alert,
  Box,
  Button,
  IconButton,
  LinearProgress,
  MenuItem,
  Stack,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  ArticleOutlinedIcon,
  ContentCopyOutlinedIcon,
  DeleteOutlineOutlinedIcon,
  FileDownloadOutlinedIcon,
  FileUploadOutlinedIcon,
  PictureAsPdfOutlinedIcon,
} from "@/components/sci/icons";
import dayjs from "dayjs";
import "dayjs/locale/uk";
import {
  api,
  type BackendDocumentSignatoryPreset,
  type BackendPersonDocument,
  type BackendPersonQuestionnaireMeta,
  type BackendPersonnelOverview,
  type BackendPersonnelOverviewRow,
} from "../../api";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
  readDataCache,
} from "../../data/idbDataCache";
import { useAuth } from "../../auth/AuthProvider";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  buildQuestionnaireExportFileName,
  collectPersonExternalIdCandidates,
  dataUrlToObjectUrl,
  downloadQuestionnairePdf,
  formatUaPhoneDisplay,
  getPersonFieldValue,
  getPersonFullPositionTitle,
  getPersonDisplayName,
  isUnstablePersonExternalId,
  migratePersonAttachmentsBetweenIds,
  pickPreferredPersonRank,
  revokeQuestionnairePreviewUrl,
} from "../personnel/personnelUtils";
import { FloatingQuestionnairePreview } from "../personnel/FloatingQuestionnairePreview";
import {
  PERSON_PHONES_DOCUMENT_TYPE,
  readStoredPersonPhones,
} from "../personnel/personPhonesStore";
import {
  applyEnrichmentToPreviewRow,
  syncEnrichmentToPerson,
} from "../personnel/personEnrichment";
import {
  extractPersonSignature,
  migrateStoredPersonSignatures,
  persistPersonSignature,
  PERSON_SIGNATURE_DOCUMENT_TYPE,
  withPersonSignature,
  type PersonSignatureRecord,
} from "../personnel/personSignatureStore";
import { exportDocumentsJournalExcel } from "./documentsJournalExcel";
import { loadAnketaEdits, applyAnketaEditsToRows } from "../anketa-data/anketaEdits";
import { normalizeAnketaNameKey } from "../anketa-data/anketaPersonMatch";
import {
  loadAnketaSheetPreferCache,
  type AnketaRow,
} from "../anketa-data/anketaSheet";
import {
  buildUbdNotSubmittedRowFromDocument,
  exportUbdNotSubmittedExcel,
  exportUbdStatusExcel,
} from "./ubdNotSubmittedExcel";
import {
  buildFighterTaskPeriodText,
  formatUbdTaskPeriodEndDate,
  getFighterTaskPlace,
  parseUbdTaskPeriodEndDate,
} from "../personnel/fighterStatusImport";
import { mapRosterLatestToPreviewRows, readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import { Spinner } from "@/components/ui/spinner/spinner";
import {
  copyPngDataUrlToClipboard,
  getCommanderSignatureTransparent,
  processSignatureTransparentBackground,
} from "./ubdSignatureImage";
import { createUbdWordBlob } from "./ubdWordExport";
import {
  createForm6Fields,
  form6WorkflowSteps,
  mergeForm6Fields,
  type Form6ReportFields,
  type Form6Signatory,
} from "./form6Report";
import { createForm6WordBlob } from "./form6WordExport";
import {
  createForm12Fields,
  form12WorkflowSteps,
  mergeForm12Fields,
  type Form12ReportFields,
  type Form12Signatory,
} from "./form12Report";
import { createForm12WordBlob } from "./form12WordExport";
import {
  createServiceCharacteristicFields,
  mergeServiceCharacteristicFields,
  serviceCharacteristicWorkflowSteps,
  type ServiceCharacteristicFields,
  type ServiceCharacteristicSignatory,
} from "./serviceCharacteristicReport";
import { createServiceCharacteristicWordBlob } from "./serviceCharacteristicWordExport";
import {
  buildZhbdBodyParagraph,
  createZhbdCertificateFields,
  getFullPositionFromPersonnelProfileRoster,
  mergeZhbdCertificateFields,
  readStoredSelectedPersonFullPosition,
  resolveZhbdCombatStaffPosition,
  storeSelectedPersonFullPosition,
  zhbdCertificateWorkflowSteps,
  type ZhbdCertificateFields,
  type ZhbdCertificateSignatory,
} from "./zhbdCertificateReport";
import { createZhbdCertificateWordBlob } from "./zhbdCertificateWordExport";
import {
  WordDocumentPreview,
  printRenderedWordPreview,
  useWordPreviewBlob,
} from "./WordDocumentPreview";
import {
  createUbdRestoreFields,
  mergeUbdRestoreFields,
  ubdRestoreWorkflowSteps,
  type UbdRestoreReportFields,
  type UbdRestoreSignatory,
} from "./ubdRestoreReport";
import { createUbdRestoreWordBlob } from "./ubdRestoreWordExport";

dayjs.locale("uk");

type DocumentMode =
  | "default"
  | "salaryPowerAttorney"
  | "ubdReport"
  | "ubdRestoreReport"
  | "form6Report"
  | "form12Report"
  | "serviceCharacteristic"
  | "zhbdCertificate"
  | "temporaryMilitaryId";

type DefaultDocumentFieldKey = "pib" | "rank" | "unit" | "date";
type DefaultDocumentFields = Record<DefaultDocumentFieldKey, string>;

type SalaryDocumentFields = {
  fullName: string;
  rnokpp: string;
  iban: string;
  bankMfo: string;
  bankName: string;
  rank: string;
  date: string;
  commander: string;
  personnelChief: string;
  folderName: string;
  statusNote: string;
};

type SalaryWorkflowState = {
  completed: Record<string, boolean>;
  currentStatus: string;
  accountFileName: string;
  signedScanFileName: string;
  mergedPdfFileName: string;
};

type UbdReportFields = {
  commander: string;
  fullName: string;
  rank: string;
  staffPosition: string;
  birthDate: string;
  rnokpp: string;
  taskPeriod: string;
  taskPlace: string;
  basis: string;
  basisNumber: string;
  basisDate: string;
  battalionCommander: string;
  approvalOfficer: string;
  signatories: DocumentSignatorySnapshot[];
  date: string;
  folderName: string;
  statusNote: string;
};

type TemporaryMilitaryIdFields = {
  fullName: string;
  rank: string;
  callSign: string;
  birthDate: string;
  photoData: string;
  statusNote: string;
};

type DocumentSignatorySnapshot = Pick<
  BackendDocumentSignatoryPreset,
  | "label"
  | "blockType"
  | "title"
  | "rank"
  | "fullName"
  | "signatureData"
  | "showDate"
  | "sortOrder"
> & { sourceId?: string };

type UbdScanFile = {
  id: string;
  name: string;
  type: string;
  dataUrl: string;
  uploadedAt: string;
};

type DocumentFiles = {
  ubdScans?: UbdScanFile[];
  ticketPhoto?: UbdScanFile | null;
};

type PersonStatusSnapshot = {
  label: string;
  detail: string;
};

type DocumentRouteState = {
  requestedPersonId: string;
  requestedDocumentType: string;
  requestedDocumentId: string;
  isPersonDocumentMode: boolean;
};

const SALARY_WORKFLOW_STORAGE_KEY = "army-grid:salary-power-attorney-workflow";

const salaryWorkflowSteps = [
  {
    key: "account",
    title: "Скрін реквізитів рахунку та ІНН",
  },
  {
    key: "document",
    title: "Заповнили довіреність",
  },
  {
    key: "print",
    title: "Роздрукували у 2 екземплярах",
  },
  {
    key: "sign",
    title: "Надали підпис службовцю та командиру",
  },
  {
    key: "scan",
    title: "Скан копії та загальний PDF",
  },
  {
    key: "ready",
    title: "Готово до відправки",
  },
  {
    key: "sent",
    title: "Відправили",
  },
  {
    key: "received",
    title: "Отримали",
  },
  {
    key: "handed",
    title: "Вручили",
  },
];

const ubdWorkflowSteps = [
  {
    key: "document",
    title: "Заповнили рапорт",
  },
  {
    key: "account",
    title: "Фото 3x4 та ІНН",
  },
  {
    key: "scan",
    title: "Скани паспорта, ІНН та статусу",
  },
  {
    key: "ready",
    title: "Готово до відправки",
  },
  {
    key: "sent",
    title: "Відправили",
  },
  {
    key: "received",
    title: "Отримали",
  },
  {
    key: "handed",
    title: "Вручили",
  },
];

const temporaryMilitaryIdWorkflowSteps = [
  {
    key: "photo",
    title: "Фото",
  },
  {
    key: "order",
    title: "Замовити",
  },
  {
    key: "received",
    title: "Отримали",
  },
  {
    key: "handed",
    title: "Вручили",
  },
];

const ukrainianBanks = [
  { mfo: "305299", name: "АТ КБ «ПРИВАТБАНК»" },
  { mfo: "300465", name: "АТ «ОЩАДБАНК»" },
  { mfo: "322001", name: "АТ «УНІВЕРСАЛ БАНК» / MONOBANK" },
  { mfo: "380805", name: "АТ «РАЙФФАЙЗЕН БАНК»" },
  { mfo: "351005", name: "АТ «УКРСИББАНК»" },
  { mfo: "334851", name: "АТ «ПУМБ»" },
  { mfo: "300346", name: "АТ «СЕНС БАНК»" },
  { mfo: "307770", name: "АТ «АКЦЕНТ-БАНК»" },
  { mfo: "320478", name: "АБ «УКРГАЗБАНК»" },
  { mfo: "300528", name: "АТ «ОТП БАНК»" },
  { mfo: "300614", name: "АТ «КРЕДІ АГРІКОЛЬ БАНК»" },
  { mfo: "325365", name: "АТ «КРЕДОБАНК»" },
  { mfo: "328209", name: "АБ «ПІВДЕННИЙ»" },
  { mfo: "339500", name: "АТ «ТАСКОМБАНК»" },
  { mfo: "328168", name: "АТ «МТБ БАНК»" },
  { mfo: "380526", name: "АТ «КБ «ГЛОБУС»" },
  { mfo: "custom", name: "Інший банк" },
];

const createDefaultFields = (
  summary: ReturnType<typeof buildPersonSummary>,
): DefaultDocumentFields => ({
  pib: summary.name !== "Особа не вибрана" ? summary.name : "",
  rank: summary.rank || "",
  unit: summary.location || "",
  date: dayjs().format("DD.MM.YYYY"),
});

const normalizeIban = (value: string) =>
  String(value ?? "")
    .replace(/\s+/g, "")
    .toUpperCase();

const formatIban = (value: string) =>
  normalizeIban(value)
    .replace(/(.{4})/g, "$1 ")
    .trim();

const extractIban = (row: EjournalPreviewRow | null) => {
  const direct =
    getPersonFieldValue(row, ["iban"]) ||
    getPersonFieldValue(row, ["iban", "рахунок"]) ||
    getPersonFieldValue(row, ["iban", "банк"]) ||
    getPersonFieldValue(row, ["рахунок"]) ||
    getPersonFieldValue(row, ["банківський", "рахунок"]) ||
    "";
  const directMatch = direct.match(/UA\s*\d{2}(?:\s*\d){25}/i);
  if (directMatch) return normalizeIban(directMatch[0]);

  const allText = Object.values(row ?? {})
    .map((value) => String(value ?? ""))
    .join(" ");
  const anyMatch = allText.match(/UA\s*\d{2}(?:\s*\d){25}/i);
  return anyMatch ? normalizeIban(anyMatch[0]) : "";
};

const getBankMfoFromIban = (iban: string) => {
  const normalized = normalizeIban(iban);
  return /^UA\d{27}$/.test(normalized) ? normalized.slice(4, 10) : "";
};

const findBankByMfo = (mfo: string) =>
  ukrainianBanks.find((bank) => bank.mfo === mfo);

const surnameUpperName = (fullName: string) => {
  const parts = fullName.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "";
  return [parts[0].toUpperCase(), ...parts.slice(1)].join(" ");
};

const safeFilePart = (value: string) =>
  String(value || "document")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const UBD_BATTALION_COMMANDER_NAME = "Єгор СИДОРЕНКО";
const UBD_APPROVAL_OFFICER_NAME = "Олег АДАМОВ";

const legacyUbdSignatories = (): DocumentSignatorySnapshot[] => [
  {
    label: "Командир батальйону",
    blockType: "SIGNER",
    title: "Командир 1 піхотного батальйону\nвійськової частини A4862",
    rank: "старший лейтенант",
    fullName: UBD_BATTALION_COMMANDER_NAME,
    signatureData: UBD_COMMANDER_SIGNATURE_SRC,
    showDate: false,
    sortOrder: 0,
  },
  {
    label: "Затверджує",
    blockType: "APPROVAL",
    title:
      "Тимчасово виконуючий обов’язки\nкомандира військової частини A4862",
    rank: "капітан",
    fullName: UBD_APPROVAL_OFFICER_NAME,
    signatureData: null,
    showDate: false,
    sortOrder: 1,
  },
];

const toForm12Signatories = (
  records: DocumentSignatorySnapshot[],
): Form12Signatory[] =>
  records.map((record) => ({
    blockType: record.blockType === "APPROVAL" ? "APPROVAL" : "SIGNER",
    title: record.title,
    rank: record.rank,
    fullName: record.fullName,
    signatureData: record.signatureData ?? null,
  }));

const loadForm6SignatoryRecords = async () => {
  const own = await api.listDocumentSignatories("form6Report");
  if (own.length) return own;
  return api.listDocumentSignatories("ubdReport");
};

const toForm6Signatories = (
  records: DocumentSignatorySnapshot[],
): Form6Signatory[] =>
  records.map((record) => ({
    blockType: record.blockType === "APPROVAL" ? "APPROVAL" : "SIGNER",
    title: record.title,
    rank: record.rank,
    fullName: record.fullName,
    signatureData: record.signatureData ?? null,
  }));

const loadServiceCharacteristicSignatoryRecords = async () => {
  const own = await api.listDocumentSignatories("serviceCharacteristic");
  if (own.length) return own;
  const form6 = await api.listDocumentSignatories("form6Report");
  if (form6.length) return form6;
  return api.listDocumentSignatories("ubdReport");
};

const toServiceCharacteristicSignatories = (
  records: DocumentSignatorySnapshot[],
): ServiceCharacteristicSignatory[] =>
  records.map((record) => ({
    blockType: record.blockType === "APPROVAL" ? "APPROVAL" : "SIGNER",
    title: record.title,
    rank: record.rank,
    fullName: record.fullName,
    signatureData: record.signatureData ?? null,
  }));

const loadZhbdCertificateSignatoryRecords = async () => {
  const own = await api.listDocumentSignatories("zhbdCertificate");
  if (own.length) return own;
  const form6 = await api.listDocumentSignatories("form6Report");
  if (form6.length) return form6;
  return api.listDocumentSignatories("ubdReport");
};

const toZhbdCertificateSignatories = (
  records: DocumentSignatorySnapshot[],
): ZhbdCertificateSignatory[] =>
  records.map((record) => ({
    blockType: record.blockType === "APPROVAL" ? "APPROVAL" : "SIGNER",
    title: record.title,
    rank: record.rank,
    fullName: record.fullName,
    signatureData: record.signatureData ?? null,
  }));

const loadForm12SignatoryRecords = async () => {
  const own = await api.listDocumentSignatories("form12Report");
  if (own.length) return own;
  return api.listDocumentSignatories("ubdReport");
};

const loadUbdRestoreSignatoryRecords = async () => {
  const own = await api.listDocumentSignatories("ubdRestoreReport");
  if (own.length) return own;
  return api.listDocumentSignatories("ubdReport");
};

const toUbdRestoreSignatories = (
  records: DocumentSignatorySnapshot[],
): UbdRestoreSignatory[] =>
  records.map((record) => ({
    blockType: record.blockType === "APPROVAL" ? "APPROVAL" : "SIGNER",
    title: record.title,
    rank: record.rank,
    fullName: record.fullName,
    signatureData: record.signatureData ?? null,
  }));

const snapshotSignatories = (
  records: BackendDocumentSignatoryPreset[],
): DocumentSignatorySnapshot[] =>
  records.map((record) => ({
    sourceId: record.id,
    label: record.label,
    blockType: record.blockType,
    title: record.title,
    rank: record.rank,
    fullName: record.fullName,
    signatureData: record.signatureData,
    showDate: record.showDate,
    sortOrder: record.sortOrder,
  }));

const UBD_COMMANDER_SIGNATURE_SRC =
  "data:image/jpeg;base64,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";

const createSalaryFields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
): SalaryDocumentFields => {
  const iban = extractIban(row);
  const bankMfo = getBankMfoFromIban(iban);
  const bank = findBankByMfo(bankMfo);
  const fullName =
    summary.name !== "Особа не вибрана" ? surnameUpperName(summary.name) : "";

  return {
    fullName,
    rnokpp: summary.rnokpp || "",
    iban,
    bankMfo: bankMfo || "custom",
    bankName: bank?.name || "",
    rank: summary.rank || "",
    date: dayjs().format("DD.MM.YYYY"),
    commander: "Командиру 1 піхотного батальйону військової частини A4862",
    personnelChief: "молодший сержант________________      Віталій БОНДАР",
    folderName: fullName ? `ГЗ довіреність - ${fullName}` : "ГЗ довіреність",
    statusNote: "",
  };
};

const UBD_BASIS_KIND = "бойове розпорядження";
/** Поки без номера/дати — підставлятимуть вручну, коли будуть вірні дані. */
const DEFAULT_UBD_BASIS_NUMBER = "";
const DEFAULT_UBD_BASIS_DATE = "";
/** Старі автопідстановки — прибираємо з документів і експорту. */
const LEGACY_UBD_BASIS_NUMBERS = new Set([
  "4862/ОКП/1162/дск",
  "4862/ОКП/1162/дск".toLocaleLowerCase("uk-UA"),
]);
const LEGACY_UBD_BASIS_DATES = new Set(["09.05.2026", "09.05.26"]);

const stripUbdBasisNumber = (value: string) =>
  String(value ?? "")
    .trim()
    .replace(/^№\s*/, "");

const sanitizeUbdBasisNumber = (value: string) => {
  const cleaned = stripUbdBasisNumber(value);
  if (!cleaned) return "";
  if (LEGACY_UBD_BASIS_NUMBERS.has(cleaned)) return "";
  if (LEGACY_UBD_BASIS_NUMBERS.has(cleaned.toLocaleLowerCase("uk-UA"))) {
    return "";
  }
  return cleaned;
};

const sanitizeUbdBasisDate = (value: string) => {
  const cleaned = String(value ?? "").trim();
  if (!cleaned) return "";
  if (LEGACY_UBD_BASIS_DATES.has(cleaned)) return "";
  return cleaned;
};

const formatUbdBasisText = (number: string, date: string) => {
  const orderNumber = sanitizeUbdBasisNumber(number);
  const orderDate = sanitizeUbdBasisDate(date);
  const reference = [
    orderNumber ? `№${orderNumber}` : "",
    orderDate ? `від ${orderDate}` : "",
  ]
    .filter(Boolean)
    .join(" ");
  return [UBD_BASIS_KIND, reference].filter(Boolean).join(" ");
};

const parseUbdBasisParts = (value: string) => {
  const text = String(value ?? "").trim();
  const match = text.match(
    /№\s*(\S+)\s+від\s+(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/i,
  );
  if (!match) {
    return {
      basisNumber: DEFAULT_UBD_BASIS_NUMBER,
      basisDate: DEFAULT_UBD_BASIS_DATE,
    };
  }
  return {
    basisNumber: sanitizeUbdBasisNumber(match[1]),
    basisDate: sanitizeUbdBasisDate(
      match[2].replaceAll("/", ".").replaceAll("-", "."),
    ),
  };
};

const createUbdFields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
  signatories: DocumentSignatorySnapshot[] = legacyUbdSignatories(),
): UbdReportFields => {
  const fullName = summary.name !== "Особа не вибрана" ? summary.name : "";
  const staffPosition =
    getPersonFieldValue(row, ["посада"]) ||
    getPersonFieldValue(row, ["штатна", "посада"]) ||
    getPersonFieldValue(row, ["чим", "займається"]) ||
    "";

  return {
    commander: "Командиру військової частини A4862",
    fullName,
    rank: summary.rank || "",
    staffPosition,
    birthDate: summary.birthDate || "",
    rnokpp: summary.rnokpp || "",
    taskPeriod: buildFighterTaskPeriodText(row),
    taskPlace: getFighterTaskPlace(row),
    basisNumber: DEFAULT_UBD_BASIS_NUMBER,
    basisDate: DEFAULT_UBD_BASIS_DATE,
    basis: formatUbdBasisText(DEFAULT_UBD_BASIS_NUMBER, DEFAULT_UBD_BASIS_DATE),
    battalionCommander:
      "Командир 1 піхотного батальйону\nвійськової частини A4862\nстарший лейтенант",
    approvalOfficer:
      "Тимчасово виконуючий обов’язки\nкомандира військової частини A4862\nкапітан",
    signatories,
    date: dayjs().format("DD.MM.YYYY"),
    folderName: fullName ? `УБД рапорт - ${fullName}` : "УБД рапорт",
    statusNote: "",
  };
};

const padBirthDatePart = (value: number) => String(value).padStart(2, "0");

const formatTemporaryIdBirthDate = (value: string) => {
  const text = String(value ?? "").trim();
  if (!text) return "";

  const dotted = text.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/);
  if (dotted) {
    const day = Number(dotted[1]);
    const month = Number(dotted[2]);
    let year = Number(dotted[3]);
    if (year < 100) year += year >= 50 ? 1900 : 2000;
    if (
      day >= 1 &&
      day <= 31 &&
      month >= 1 &&
      month <= 12 &&
      year >= 1900 &&
      year <= 2100
    ) {
      return `${padBirthDatePart(day)}.${padBirthDatePart(month)}.${year}`;
    }
  }

  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[3]}.${iso[2]}.${iso[1]}`;

  return text;
};

const isFullBirthDate = (value: string) =>
  /^\d{2}\.\d{2}\.\d{4}$/.test(value.trim());

const createTemporaryMilitaryIdFields = (
  summary: ReturnType<typeof buildPersonSummary>,
  photoData = "",
): TemporaryMilitaryIdFields => {
  const fullName = summary.name !== "Особа не вибрана" ? summary.name : "";
  return {
    fullName,
    rank: summary.rank || "",
    callSign: summary.callSign || "",
    birthDate: formatTemporaryIdBirthDate(summary.birthDate),
    photoData,
    statusNote: "",
  };
};

const mergeTemporaryMilitaryIdFields = (
  defaults: TemporaryMilitaryIdFields,
  value: unknown,
): TemporaryMilitaryIdFields => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return defaults;
  }

  const saved = value as Partial<TemporaryMilitaryIdFields> & {
    birthYear?: string;
  };
  const { birthYear: _legacyBirthYear, ...savedFields } = saved;
  const savedDate = formatTemporaryIdBirthDate(
    saved.birthDate || saved.birthYear || "",
  );
  const birthDate = isFullBirthDate(savedDate)
    ? savedDate
    : isFullBirthDate(defaults.birthDate)
      ? defaults.birthDate
      : savedDate || defaults.birthDate;

  return {
    ...defaults,
    ...savedFields,
    rank: pickPreferredPersonRank(defaults.rank, saved.rank),
    birthDate,
    photoData: saved.photoData || defaults.photoData,
    statusNote: saved.statusNote ?? defaults.statusNote,
  };
};

const formatTemporaryIdDispatchText = (fields: TemporaryMilitaryIdFields) =>
  [
    `Звання: ${fields.rank.trim() || "—"}`,
    `ПІБ: ${fields.fullName.trim() || "—"}`,
    `Позивний: ${fields.callSign.trim() || "—"}`,
    `Дата народження: ${fields.birthDate.trim() || "—"}`,
    fields.statusNote.trim()
      ? `Додатково по статусу: ${fields.statusNote.trim()}`
      : "",
  ]
    .filter(Boolean)
    .join("\n");

const createEmptyWorkflow = (): SalaryWorkflowState => ({
  completed: {},
  currentStatus: salaryWorkflowSteps[0]?.key ?? "account",
  accountFileName: "",
  signedScanFileName: "",
  mergedPdfFileName: "",
});

const mergeSalaryFields = (
  defaults: SalaryDocumentFields,
  value: unknown,
): SalaryDocumentFields => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return defaults;
  }
  const merged = {
    ...defaults,
    ...(value as Partial<SalaryDocumentFields>),
  };
  const pick = (personnel: string, document: string) =>
    String(personnel ?? "").trim() || String(document ?? "").trim();
  return {
    ...merged,
    fullName: pick(defaults.fullName, merged.fullName),
    rnokpp: pick(defaults.rnokpp, merged.rnokpp),
    rank: pick(defaults.rank, merged.rank),
  };
};

const mergeSalaryWorkflow = (value: unknown): SalaryWorkflowState =>
  value && typeof value === "object" && !Array.isArray(value)
    ? { ...createEmptyWorkflow(), ...(value as Partial<SalaryWorkflowState>) }
    : createEmptyWorkflow();

const mergeUbdFields = (
  defaults: UbdReportFields,
  value: unknown,
): UbdReportFields => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return defaults;
  }

  const saved = value as Partial<UbdReportFields>;
  const parsed = parseUbdBasisParts(
    String(saved.basisNumber ? "" : saved.basis ?? defaults.basis),
  );
  const basisNumber =
    sanitizeUbdBasisNumber(String(saved.basisNumber ?? "")) ||
    sanitizeUbdBasisNumber(parsed.basisNumber);
  const basisDate =
    sanitizeUbdBasisDate(String(saved.basisDate ?? "")) ||
    sanitizeUbdBasisDate(parsed.basisDate);
  const merged = { ...defaults, ...saved };
  const pick = (personnel: string, document: string) =>
    String(personnel ?? "").trim() || String(document ?? "").trim();
  return {
    ...merged,
    fullName: pick(defaults.fullName, merged.fullName),
    rank: pick(defaults.rank, merged.rank),
    staffPosition: pick(defaults.staffPosition, merged.staffPosition),
    birthDate: pick(defaults.birthDate, merged.birthDate),
    rnokpp: pick(defaults.rnokpp, merged.rnokpp),
    basisNumber,
    basisDate,
    basis: formatUbdBasisText(basisNumber, basisDate),
    taskPeriod: merged.taskPeriod.trim() || defaults.taskPeriod,
    taskPlace: merged.taskPlace.trim() || defaults.taskPlace,
    signatories:
      Array.isArray(saved.signatories) && saved.signatories.length
        ? saved.signatories
        : defaults.signatories,
  };
};

const mergeDocumentFiles = (value: unknown): DocumentFiles =>
  value && typeof value === "object" && !Array.isArray(value)
    ? { ...(value as DocumentFiles) }
    : {};

/** Seed Form 6 phone from saved person phones if the form field is empty. */
const withSavedPersonPhone = (
  fields: Form6ReportFields,
  personExternalId: string,
): Form6ReportFields => {
  if (fields.phone.trim()) return fields;
  const stored = readStoredPersonPhones()[personExternalId.trim()] ?? [];
  if (!stored.length) return fields;
  return { ...fields, phone: formatUaPhoneDisplay(stored[0]) };
};

const compactText = (value: unknown) =>
  String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();

const createPersonStatusSnapshot = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
): PersonStatusSnapshot => {
  const rawStatus =
    getPersonFieldValue(row, ["статус"]) ||
    getPersonFieldValue(row, ["причина", "статус"]) ||
    getPersonFieldValue(row, ["в", "якому", "статусі"]) ||
    "";
  const detail =
    summary.location ||
    getPersonFieldValue(row, ["примітка"]) ||
    getPersonFieldValue(row, ["коментар"]) ||
    "";

  return {
    label: compactText(rawStatus || detail || "—"),
    detail: compactText(detail),
  };
};

const overviewStatusSnapshot = (
  row: BackendPersonnelOverviewRow,
): PersonStatusSnapshot => ({
  label: compactText(row.place || row.statusLabel || row.unit || "—"),
  detail: compactText(
    [
      row.statusLabel,
      row.place,
      row.plannedReturn ? `до ${row.plannedReturn}` : "",
    ]
      .filter(Boolean)
      .join(" · "),
  ),
});

const readDocumentPersonStatus = (
  document: BackendPersonDocument,
  overviewById: Record<string, PersonStatusSnapshot>,
) => {
  const fields = document.fields || {};
  const fromFields =
    typeof fields.personStatus === "string"
      ? fields.personStatus
      : typeof fields.serviceStatus === "string"
        ? fields.serviceStatus
        : "";
  const fromOverview = overviewById[document.personExternalId];
  return compactText(fromOverview?.label || fromFields || "—");
};

const CREATE_PERSON_DOCUMENT_TYPES = [
  { value: "salaryPowerAttorney", label: "Довіреність зарплати" },
  { value: "ubdReport", label: "Рапорт на УБД" },
  { value: "ubdRestoreReport", label: "Рапорт на відновлення УБД" },
  { value: "form6Report", label: "Форма 6" },
  { value: "form12Report", label: "Форма 12" },
  { value: "serviceCharacteristic", label: "Службова характеристика" },
  { value: "zhbdCertificate", label: "Довідка ЖБД" },
  { value: "temporaryMilitaryId", label: "Тимчасовий військовий квиток" },
] as const;

const JOURNAL_DOCUMENT_TYPE_FILTERS = [
  { value: "ALL", label: "Усі документи" },
  { value: "ubdReport", label: "Рапорт на УБД" },
  { value: "ubdRestoreReport", label: "Рапорт на відновлення УБД" },
  { value: "form6Report", label: "Форма 6" },
  { value: "form12Report", label: "Форма 12" },
  { value: "serviceCharacteristic", label: "Службова характеристика" },
  { value: "zhbdCertificate", label: "Довідка ЖБД" },
  { value: "temporaryMilitaryId", label: "Тимчасовий військовий квиток" },
  { value: "salaryPowerAttorney", label: "Довіреність зарплати" },
] as const;

const salaryDocumentTypeLabel = (type: string) =>
  type === "salaryPowerAttorney"
    ? "Довіреність зарплати"
    : type === "ubdReport"
      ? "Рапорт на УБД"
      : type === "ubdRestoreReport"
        ? "Рапорт на відновлення УБД"
        : type === "form6Report"
          ? "Форма 6"
          : type === "form12Report"
            ? "Форма 12"
            : type === "serviceCharacteristic"
              ? "Службова характеристика"
              : type === "zhbdCertificate"
                ? "Довідка ЖБД"
          : type === "temporaryMilitaryId"
            ? "Тимчасовий військовий квиток"
            : type;

const buildTicketPhotoFile = (photoData: string) =>
  photoData
    ? {
        id: `photo-${Date.now()}`,
        name: "фото.jpg",
        type: "image/jpeg",
        dataUrl: photoData,
        uploadedAt: new Date().toISOString(),
      }
    : null;

const buildPersonDocumentDraft = (
  type: DocumentMode,
  input: {
    personStatus: string;
    salaryFields: SalaryDocumentFields;
    ubdFields: UbdReportFields;
    form6Fields: Form6ReportFields;
    form12Fields: Form12ReportFields;
    serviceCharacteristicFields: ServiceCharacteristicFields;
    zhbdCertificateFields: ZhbdCertificateFields;
    restoreFields: UbdRestoreReportFields;
    ticketFields: TemporaryMilitaryIdFields;
  },
) => {
  const now = dayjs().format("DD.MM.YYYY HH:mm");
  if (type === "salaryPowerAttorney") {
    const workflow = createEmptyWorkflow();
    return {
      type,
      title: `Довіреність зарплати · ${now}`,
      status: workflow.currentStatus,
      fields: { ...input.salaryFields, personStatus: input.personStatus },
      workflow: workflow as unknown as Record<string, unknown>,
    };
  }
  if (type === "ubdReport") {
    const workflow = { ...createEmptyWorkflow(), currentStatus: "document" };
    return {
      type,
      title: `Рапорт на УБД · ${now}`,
      status: "document",
      fields: { ...input.ubdFields, personStatus: input.personStatus },
      workflow: workflow as unknown as Record<string, unknown>,
      files: { ubdScans: [] },
    };
  }
  if (type === "form6Report") {
    const workflow = { ...createEmptyWorkflow(), currentStatus: "document" };
    return {
      type,
      title: `Форма 6 · ${now}`,
      status: "document",
      fields: { ...input.form6Fields, personStatus: input.personStatus },
      workflow: workflow as unknown as Record<string, unknown>,
      files: { ubdScans: [] },
    };
  }
  if (type === "form12Report") {
    const workflow = { ...createEmptyWorkflow(), currentStatus: "document" };
    return {
      type,
      title: `Форма 12 · ${now}`,
      status: "document",
      fields: { ...input.form12Fields, personStatus: input.personStatus },
      workflow: workflow as unknown as Record<string, unknown>,
      files: { ubdScans: [] },
    };
  }
  if (type === "serviceCharacteristic") {
    const workflow = { ...createEmptyWorkflow(), currentStatus: "created" };
    return {
      type,
      title: `Службова характеристика · ${now}`,
      status: "created",
      fields: {
        ...input.serviceCharacteristicFields,
        personStatus: input.personStatus,
      },
      workflow: workflow as unknown as Record<string, unknown>,
      files: { ubdScans: [] },
    };
  }
  if (type === "zhbdCertificate") {
    const workflow = { ...createEmptyWorkflow(), currentStatus: "created" };
    return {
      type,
      title: `Довідка ЖБД · ${now}`,
      status: "created",
      fields: {
        ...input.zhbdCertificateFields,
        personStatus: input.personStatus,
      },
      workflow: workflow as unknown as Record<string, unknown>,
      files: { ubdScans: [] },
    };
  }
  if (type === "ubdRestoreReport") {
    const workflow = { ...createEmptyWorkflow(), currentStatus: "document" };
    return {
      type,
      title: `Рапорт на відновлення УБД · ${now}`,
      status: "document",
      fields: { ...input.restoreFields, personStatus: input.personStatus },
      workflow: workflow as unknown as Record<string, unknown>,
      files: { ubdScans: [] },
    };
  }
  if (type === "temporaryMilitaryId") {
    const workflow = { ...createEmptyWorkflow(), currentStatus: "photo" };
    const ticketPhoto = buildTicketPhotoFile(input.ticketFields.photoData);
    return {
      type,
      title: `Тимчасовий військовий квиток · ${now}`,
      status: "photo",
      fields: { ...input.ticketFields, personStatus: input.personStatus },
      workflow: workflow as unknown as Record<string, unknown>,
      files: { ticketPhoto, ubdScans: [] },
    };
  }
  return null;
};

const autoCreateInFlight = new Set<string>();

const workflowStatusLabel = (status?: string | null) =>
  [
    ...salaryWorkflowSteps,
    ...ubdWorkflowSteps,
    ...ubdRestoreWorkflowSteps,
    ...form6WorkflowSteps,
    ...form12WorkflowSteps,
    ...serviceCharacteristicWorkflowSteps,
    ...zhbdCertificateWorkflowSteps,
    ...temporaryMilitaryIdWorkflowSteps,
  ].find((step) => step.key === status)?.title ||
  status ||
  "статус не заданий";

const documentWorkflowSteps = (type?: string | null) =>
  type === "ubdRestoreReport"
    ? ubdRestoreWorkflowSteps
    : type === "form12Report"
      ? form12WorkflowSteps
      : type === "serviceCharacteristic"
        ? serviceCharacteristicWorkflowSteps
      : type === "zhbdCertificate"
        ? zhbdCertificateWorkflowSteps
      : type === "ubdReport" || type === "form6Report"
        ? ubdWorkflowSteps
        : type === "temporaryMilitaryId"
          ? temporaryMilitaryIdWorkflowSteps
          : salaryWorkflowSteps;

const resolveDocumentWorkflowStatus = (
  type: string | null | undefined,
  status?: string | null,
) => {
  if (type === "ubdRestoreReport") {
    if (status === "account" || status === "print" || status === "sign") {
      return "photo";
    }
    if (status === "ready") return "sent";
    return status ?? "";
  }
  if (type === "form12Report") {
    if (
      status === "account" ||
      status === "scan" ||
      status === "ready" ||
      status === "print" ||
      status === "sign"
    ) {
      return "sent";
    }
    return status ?? "";
  }
  if (type === "serviceCharacteristic" || type === "zhbdCertificate") {
    if (status === "ready") return "sent";
    if (status === "signed") return "confirmed";
    if (status === "handed") return "forCharacteristic";
    return status ?? "";
  }
  if (
    (type === "ubdReport" || type === "form6Report") &&
    (status === "print" || status === "sign")
  ) {
    return "scan";
  }
  return status ?? "";
};

const documentWorkflowStatusLabel = (document: BackendPersonDocument) =>
  documentWorkflowSteps(document.type).find(
    (step) =>
      step.key ===
      resolveDocumentWorkflowStatus(document.type, document.status),
  )?.title || workflowStatusLabel(document.status);

/** УБД-експорт: не включати від кроку «Відправили» і далі. */
const isUbdDocumentSentOrLater = (document: BackendPersonDocument) => {
  const workflow =
    document.workflow && typeof document.workflow === "object"
      ? (document.workflow as { currentStatus?: unknown })
      : {};
  const status = resolveDocumentWorkflowStatus(
    document.type,
    document.status ||
      (typeof workflow.currentStatus === "string"
        ? workflow.currentStatus
        : null),
  );
  return status === "sent" || status === "received" || status === "handed";
};

const getDocumentProgressPercent = (document: BackendPersonDocument) => {
  const workflow = mergeSalaryWorkflow(document.workflow);
  const steps = documentWorkflowSteps(document.type);
  const index = steps.findIndex(
    (step) =>
      step.key ===
      resolveDocumentWorkflowStatus(
        document.type,
        document.status || workflow.currentStatus,
      ),
  );
  const resolvedIndex = Math.max(0, index);
  return Math.round(((resolvedIndex + 1) / steps.length) * 100);
};

const getDocumentPersonName = (document: BackendPersonDocument) => {
  const fields = document.fields || {};
  const value =
    fields.fullName ||
    fields.pib ||
    fields.name ||
    fields["ПІБ"] ||
    fields["ФИО"];
  return typeof value === "string" && value.trim()
    ? value.trim()
    : `ID ${document.personExternalId}`;
};

const readDocumentUbdTaskPeriod = (document: BackendPersonDocument) => {
  if (document.type !== "ubdReport") return "";
  const value = document.fields?.taskPeriod;
  return typeof value === "string" ? value.trim() : "";
};

const readDocumentUbdTaskPeriodEndLabel = (document: BackendPersonDocument) =>
  formatUbdTaskPeriodEndDate(readDocumentUbdTaskPeriod(document));

const readDocumentUbdTaskPeriodEndTimestamp = (
  document: BackendPersonDocument,
) => {
  const date = parseUbdTaskPeriodEndDate(readDocumentUbdTaskPeriod(document));
  return date ? date.getTime() : 0;
};

type JournalSortField = "createdAt" | "updatedAt" | "progress" | "taskPeriodEnd";

const normalizeJournalSearchText = (value: string) =>
  value
    .trim()
    .toLocaleLowerCase("uk-UA")
    .replace(/['’`´]/g, "")
    .replace(/\s+/g, " ");

const documentMatchesJournalNameQuery = (
  document: BackendPersonDocument,
  query: string,
) => {
  const normalizedQuery = normalizeJournalSearchText(query);
  if (!normalizedQuery) return true;
  const name = normalizeJournalSearchText(getDocumentPersonName(document));
  const personId = normalizeJournalSearchText(document.personExternalId || "");
  return name.includes(normalizedQuery) || personId.includes(normalizedQuery);
};

const readDocumentStatusNote = (document: BackendPersonDocument) => {
  const value = document.fields?.statusNote;
  return typeof value === "string" ? value.trim() : "";
};

const documentCreatedMonthKey = (value?: string | null) => {
  const date = dayjs(value);
  return date.isValid() ? date.format("YYYY-MM") : "";
};

const formatJournalMonthLabel = (monthKey: string) => {
  const date = dayjs(`${monthKey}-01`);
  if (!date.isValid()) return monthKey;
  const label = date.locale("uk").format("MMMM YYYY");
  return label.charAt(0).toLocaleUpperCase("uk-UA") + label.slice(1);
};

const getDocumentFileCount = (document: BackendPersonDocument) => {
  const files = mergeDocumentFiles(document.files);
  if (document.type === "ubdReport") {
    return files.ubdScans?.length ?? 0;
  }
  if (document.type === "temporaryMilitaryId") {
    const hasPhoto = Boolean(
      files.ticketPhoto?.dataUrl ||
        (typeof document.fields?.photoData === "string" &&
          document.fields.photoData),
    );
    return (hasPhoto ? 1 : 0) + (files.ubdScans?.length ?? 0);
  }

  const workflow = mergeSalaryWorkflow(document.workflow);
  return [
    workflow.accountFileName,
    workflow.signedScanFileName,
    workflow.mergedPdfFileName,
  ].filter((value) => value.trim()).length;
};

const getDocumentFileSummary = (document: BackendPersonDocument) =>
  document.type === "ubdReport"
    ? `${getDocumentFileCount(document)} сканів`
    : document.type === "temporaryMilitaryId"
      ? getDocumentFileCount(document)
        ? "є фото"
        : "без фото"
      : `${getDocumentFileCount(document)} / 3`;

const personWorkflowKey = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
) =>
  String(
    summary.externalId ||
      (row?.__dbRowId && !isUnstablePersonExternalId(String(row.__dbRowId))
        ? row.__dbRowId
        : "") ||
      summary.name ||
      "unknown",
  );

const loadWorkflowMap = () => {
  try {
    const raw = window.localStorage.getItem(SALARY_WORKFLOW_STORAGE_KEY);
    return raw ? (JSON.parse(raw) as Record<string, SalaryWorkflowState>) : {};
  } catch {
    return {};
  }
};

const saveWorkflowForPerson = (key: string, workflow: SalaryWorkflowState) => {
  const current = loadWorkflowMap();
  current[key] = workflow;
  window.localStorage.setItem(
    SALARY_WORKFLOW_STORAGE_KEY,
    JSON.stringify(current),
  );
};

const highlightOrPlaceholder = (value: string, placeholder: string) =>
  value.trim() ? <mark>{value}</mark> : <mark>{placeholder}</mark>;

const salaryDocumentHtml = (fields: SalaryDocumentFields) => `
<!doctype html>
<html lang="uk">
<head>
  <meta charset="utf-8" />
  <title>Рапорт ГЗ ${safeFilePart(fields.fullName)}</title>
  <style>
    @page { size: A4; margin: 24mm 20mm; }
    body { font-family: "Times New Roman", serif; color: #000; font-size: 14pt; line-height: 1.35; }
    .to { text-align: right; margin-left: 42%; }
    h1 { margin: 28mm 0 12mm; text-align: center; font-size: 16pt; }
    p { margin: 0 0 9mm; }
    .sign-row { display: grid; grid-template-columns: 1fr 44mm; gap: 12mm; margin-top: 18mm; }
    .finance { margin-top: 18mm; }
    .small { font-size: 12pt; }
  </style>
</head>
<body>
  <div class="to">${fields.commander}</div>
  <h1>Рапорт</h1>
  <p>
    Прошу моє грошове забезпечення перераховувати на мій розрахунковий
    банківський рахунок:
  </p>
  <p><strong>${formatIban(fields.iban) || "UA___________________________"}</strong>, відкритий ${fields.bankName || "________________"}.</p>
  <p>Код РНОКПП <strong>${fields.rnokpp || "____________"}</strong>.</p>
  <p>Довідка з банківськими реквізитами додається.</p>
  <p>${fields.rank || "________________"} ${fields.fullName || "____________________________"}</p>
  <div class="sign-row">
    <span>${fields.date || "__.__.202__"}</span>
    <span>________________</span>
  </div>
  <div class="finance">
    <p>
      Начальнику фінансово-економічної служби – головному бухгалтеру
      військової частини А4862 - прошу прийняти в роботу
    </p>
    <p>Начальник групи персоналу штабу<br />1 піхотного батальйону<br />військової частини А4862</p>
    <p>${fields.personnelChief}</p>
    <p class="small">(підпис)</p>
  </div>
</body>
</html>`;

const downloadTextFile = (
  fileName: string,
  content: string,
  mimeType: string,
) => {
  const fileContent = mimeType.startsWith("application/msword")
    ? `\ufeff${content}`
    : content;
  const url = URL.createObjectURL(
    new Blob([fileContent], { type: mimeType }),
  );
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

const downloadBlob = (fileName: string, blob: Blob) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

const openDataUrlInNewTab = (dataUrl: string) => {
  const objectUrl = dataUrlToObjectUrl(dataUrl);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 60_000);
};

const readFileAsDataUrl = (file: File) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

const getDocumentRouteState = (): DocumentRouteState => {
  const urlParams = new URLSearchParams(window.location.search);
  const pathPersonId = window.location.pathname.startsWith("/documents/")
    ? decodeURIComponent(window.location.pathname.slice("/documents/".length))
        .split("/")[0]
        ?.trim() || ""
    : "";
  const requestedPersonId =
    pathPersonId ||
    urlParams.get("personId")?.trim() ||
    urlParams.get("externalId")?.trim() ||
    urlParams.get("personRowId")?.trim() ||
    "";

  return {
    requestedPersonId,
    requestedDocumentType: urlParams.get("type") || "",
    requestedDocumentId: urlParams.get("documentId")?.trim() || "",
    isPersonDocumentMode: Boolean(requestedPersonId),
  };
};

export function DocumentsPage(_props: {
  onNavigate?: (path: string) => void;
}) {
  const { canEditArea, user } = useAuth();
  const canEdit = canEditArea("documents");
  const isDocumentOwner = useCallback(
    (document: BackendPersonDocument | null | undefined) => {
      if (!user || !document) return false;
      if (document.createdByUserId) {
        return document.createdByUserId === user.id;
      }
      if (document.createdByEmail) {
        return (
          document.createdByEmail.toLowerCase() === user.email.toLowerCase()
        );
      }
      return false;
    },
    [user],
  );
  const canMutateDocument = useCallback(
    (document: BackendPersonDocument | null | undefined) =>
      Boolean(canEdit && isDocumentOwner(document)),
    [canEdit, isDocumentOwner],
  );
  const documentCreatorLabel = (document: BackendPersonDocument) =>
    document.createdByDisplayName?.trim() ||
    document.createdByEmail?.trim() ||
    "Невідомий автор";
  const {
    requestedPersonId,
    requestedDocumentType,
    requestedDocumentId,
    isPersonDocumentMode,
  } = getDocumentRouteState();
  const [selectedPerson, setSelectedPerson] =
    useState<EjournalPreviewRow | null>(null);
  const [mode, setMode] = useState<DocumentMode>("default");
  const [fields, setFields] = useState<DefaultDocumentFields>(() =>
    createDefaultFields(buildPersonSummary(null)),
  );
  const [salaryFields, setSalaryFields] = useState<SalaryDocumentFields>(() =>
    createSalaryFields(null, buildPersonSummary(null)),
  );
  const [ubdFields, setUbdFields] = useState<UbdReportFields>(() =>
    createUbdFields(null, buildPersonSummary(null)),
  );
  const [form6Fields, setForm6Fields] = useState<Form6ReportFields>(() =>
    createForm6Fields(null, buildPersonSummary(null)),
  );
  const [form12Fields, setForm12Fields] = useState<Form12ReportFields>(() =>
    createForm12Fields(null, buildPersonSummary(null)),
  );
  const [serviceCharacteristicFields, setServiceCharacteristicFields] =
    useState<ServiceCharacteristicFields>(() =>
      createServiceCharacteristicFields(null, buildPersonSummary(null)),
    );
  const [zhbdCertificateFields, setZhbdCertificateFields] =
    useState<ZhbdCertificateFields>(() =>
      createZhbdCertificateFields(null, buildPersonSummary(null)),
    );
  const [rosterResolvedFullPosition, setRosterResolvedFullPosition] =
    useState(() => readStoredSelectedPersonFullPosition());
  const [ubdRestoreFields, setUbdRestoreFields] =
    useState<UbdRestoreReportFields>(() =>
      createUbdRestoreFields(null, buildPersonSummary(null)),
    );
  const [ticketFields, setTicketFields] = useState<TemporaryMilitaryIdFields>(
    () => createTemporaryMilitaryIdFields(buildPersonSummary(null)),
  );
  const [commanderSignatureSrc, setCommanderSignatureSrc] = useState(
    UBD_COMMANDER_SIGNATURE_SRC,
  );
  const [documentFiles, setDocumentFiles] = useState<DocumentFiles>({});
  const [workflow, setWorkflow] =
    useState<SalaryWorkflowState>(createEmptyWorkflow);
  const [personDocuments, setPersonDocuments] = useState<
    BackendPersonDocument[]
  >([]);
  const [personQuestionnaire, setPersonQuestionnaire] =
    useState<BackendPersonQuestionnaireMeta | null>(null);
  const [isQuestionnairePreviewOpen, setIsQuestionnairePreviewOpen] =
    useState(false);
  const [questionnairePreviewUrl, setQuestionnairePreviewUrl] = useState("");
  const [questionnaireFileData, setQuestionnaireFileData] = useState("");
  const [questionnairePreviewPersonId, setQuestionnairePreviewPersonId] =
    useState("");
  const [isLoadingQuestionnairePreview, setIsLoadingQuestionnairePreview] =
    useState(false);
  const questionnaireLoadSeqRef = useRef(0);
  const autoSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const pendingFieldSaveRef = useRef<(() => void) | null>(null);
  const [selectedDocumentId, setSelectedDocumentId] = useState("");
  const [documentMessage, setDocumentMessage] = useState("");
  const [isSavingDocument, setIsSavingDocument] = useState(false);
  const [allPersonDocuments, setAllPersonDocuments] = useState<
    BackendPersonDocument[]
  >([]);
  const [personStatusById, setPersonStatusById] = useState<
    Record<string, PersonStatusSnapshot>
  >({});
  const [isLoadingDocumentJournal, setIsLoadingDocumentJournal] =
    useState(true);
  const [journalTypeFilter, setJournalTypeFilter] = useState("ALL");
  const [journalStatusFilter, setJournalStatusFilter] = useState("ALL");
  const [journalMonthFilter, setJournalMonthFilter] = useState("ALL");
  const [journalCreatorFilter, setJournalCreatorFilter] = useState("MINE");
  const [journalNameQuery, setJournalNameQuery] = useState("");
  const showUbdExitDateColumn = journalTypeFilter === "ubdReport";
  const [journalSortField, setJournalSortField] = useState<JournalSortField>(
    "createdAt",
  );
  const [journalSortDirection, setJournalSortDirection] = useState<
    "desc" | "asc"
  >("desc");
  const toggleJournalSort = (field: JournalSortField) => {
    if (journalSortField === field) {
      setJournalSortDirection((current) =>
        current === "desc" ? "asc" : "desc",
      );
      return;
    }
    setJournalSortField(field);
    setJournalSortDirection("desc");
  };
  const [isExportingDocumentJournal, setIsExportingDocumentJournal] =
    useState(false);
  const ubdWordFields = useMemo(
    () => ({
      ...ubdFields,
      signatories: ubdFields.signatories.map((signatory) =>
        signatory.signatureData === UBD_COMMANDER_SIGNATURE_SRC
          ? { ...signatory, signatureData: commanderSignatureSrc }
          : signatory,
      ),
    }),
    [commanderSignatureSrc, ubdFields],
  );
  const wordPreviewEnabled =
    mode === "ubdReport" ||
    mode === "form6Report" ||
    mode === "form12Report" ||
    mode === "serviceCharacteristic" ||
    mode === "zhbdCertificate" ||
    mode === "ubdRestoreReport";
  const buildActiveWordBlob = useCallback(() => {
    if (mode === "ubdReport") return createUbdWordBlob(ubdWordFields);
    if (mode === "form6Report") return createForm6WordBlob(form6Fields);
    if (mode === "form12Report") return createForm12WordBlob(form12Fields);
    if (mode === "serviceCharacteristic") {
      return createServiceCharacteristicWordBlob(serviceCharacteristicFields);
    }
    if (mode === "zhbdCertificate") {
      return createZhbdCertificateWordBlob(zhbdCertificateFields);
    }
    if (mode === "ubdRestoreReport") {
      return createUbdRestoreWordBlob(ubdRestoreFields);
    }
    return Promise.reject(new Error("Немає Word-шаблону для цього документа."));
  }, [
    form12Fields,
    form6Fields,
    mode,
    serviceCharacteristicFields,
    zhbdCertificateFields,
    ubdRestoreFields,
    ubdWordFields,
  ]);
  const wordPreview = useWordPreviewBlob(buildActiveWordBlob, wordPreviewEnabled);


  const journalStatusOptions = useMemo(() => {
    const steps =
      journalTypeFilter === "ALL"
        ? [
            ...ubdWorkflowSteps,
            ...ubdRestoreWorkflowSteps,
            ...form12WorkflowSteps,
            ...temporaryMilitaryIdWorkflowSteps,
            ...salaryWorkflowSteps,
          ]
        : documentWorkflowSteps(journalTypeFilter);
    const titles: string[] = [];
    for (const step of steps) {
      if (!titles.includes(step.title)) titles.push(step.title);
    }
    return titles;
  }, [journalTypeFilter]);

  const journalDocuments = useMemo(
    () =>
      allPersonDocuments.filter(
        (document) =>
          document.type !== PERSON_PHONES_DOCUMENT_TYPE &&
          document.type !== PERSON_SIGNATURE_DOCUMENT_TYPE,
      ),
    [allPersonDocuments],
  );

  const journalMonthOptions = useMemo(() => {
    const months = new Set<string>();
    for (const document of journalDocuments) {
      const monthKey = documentCreatedMonthKey(document.createdAt);
      if (monthKey) months.add(monthKey);
    }
    return [...months].sort((left, right) => right.localeCompare(left));
  }, [journalDocuments]);

  const journalCreatorOptions = useMemo(() => {
    const byId = new Map<
      string,
      { id: string; label: string; email: string }
    >();
    for (const document of journalDocuments) {
      const id =
        document.createdByUserId?.trim() ||
        document.createdByEmail?.trim() ||
        "";
      if (!id || byId.has(id)) continue;
      byId.set(id, {
        id,
        label: documentCreatorLabel(document),
        email: document.createdByEmail?.trim() || "",
      });
    }
    return [...byId.values()].sort((left, right) =>
      left.label.localeCompare(right.label, "uk"),
    );
  }, [journalDocuments]);

  const filteredJournalDocuments = useMemo(() => {
    const filtered = journalDocuments.filter((document) => {
      if (!documentMatchesJournalNameQuery(document, journalNameQuery)) {
        return false;
      }
      if (
        journalTypeFilter !== "ALL" &&
        document.type !== journalTypeFilter
      ) {
        return false;
      }
      if (
        journalStatusFilter !== "ALL" &&
        documentWorkflowStatusLabel(document) !== journalStatusFilter
      ) {
        return false;
      }
      if (
        journalMonthFilter !== "ALL" &&
        documentCreatedMonthKey(document.createdAt) !== journalMonthFilter
      ) {
        return false;
      }
      if (journalCreatorFilter === "MINE") {
        if (!isDocumentOwner(document)) return false;
      } else if (journalCreatorFilter !== "ALL") {
        const creatorKey =
          document.createdByUserId?.trim() ||
          document.createdByEmail?.trim() ||
          "";
        if (creatorKey !== journalCreatorFilter) return false;
      }
      return true;
    });

    const direction = journalSortDirection === "desc" ? -1 : 1;
    return [...filtered].sort((left, right) => {
      let compare = 0;
      if (journalSortField === "progress") {
        compare =
          getDocumentProgressPercent(left) - getDocumentProgressPercent(right);
      } else if (journalSortField === "taskPeriodEnd") {
        compare =
          readDocumentUbdTaskPeriodEndTimestamp(left) -
          readDocumentUbdTaskPeriodEndTimestamp(right);
      } else {
        const leftTs = dayjs(
          journalSortField === "updatedAt" ? left.updatedAt : left.createdAt,
        ).valueOf();
        const rightTs = dayjs(
          journalSortField === "updatedAt" ? right.updatedAt : right.createdAt,
        ).valueOf();
        const leftSafe = Number.isFinite(leftTs) ? leftTs : 0;
        const rightSafe = Number.isFinite(rightTs) ? rightTs : 0;
        compare = leftSafe - rightSafe;
      }
      if (compare !== 0) return compare * direction;
      return left.id.localeCompare(right.id) * direction;
    });
  }, [
    isDocumentOwner,
    journalCreatorFilter,
    journalDocuments,
    journalMonthFilter,
    journalNameQuery,
    journalSortDirection,
    journalSortField,
    journalStatusFilter,
    journalTypeFilter,
  ]);

  const liveJournalDocument = (document: BackendPersonDocument) =>
    document.id === selectedDocumentId
      ? {
          ...document,
          status: workflow.currentStatus || document.status,
          workflow: {
            ...(document.workflow && typeof document.workflow === "object"
              ? document.workflow
              : {}),
            ...workflow,
          },
        }
      : document;

  const buildFilteredUbdExportRows = async () => {
    const exportDocuments = filteredJournalDocuments.filter(
      (document) => !isUbdDocumentSentOrLater(liveJournalDocument(document)),
    );
    if (!exportDocuments.length) {
      throw new Error(
        "Немає рядків для експорту: усі відфільтровані вже від статусу «Відправили».",
      );
    }

    const normalizeName = (value: string) =>
      value
        .replace(/[ʼ’']/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .toLocaleLowerCase("uk-UA");

    const [rosterLatest, anketaSnapshot] = await Promise.all([
      fetchWithCache({
        key: CacheKeys.rosterLatest,
        fetcher: () => api.getLatestPersonnelRoster(),
        isChanged: jsonChanged,
      }).catch(() => null),
      loadAnketaSheetPreferCache().catch(() => null),
    ]);

    const rosterRows = mapRosterLatestToPreviewRows(rosterLatest);
    const rosterByName = new Map<string, (typeof rosterRows)[number]>();
    const rosterById = new Map<string, (typeof rosterRows)[number]>();
    for (const row of rosterRows) {
      const name = normalizeName(
        String(row["піб"] ?? row["ПІБ"] ?? row["прізвище"] ?? ""),
      );
      if (name) rosterByName.set(name, row);
      const id = String(row.__dbRowId ?? row.id ?? "").trim();
      if (id) rosterById.set(id, row);
    }

    const anketaEdits = await loadAnketaEdits().catch(() => ({}));
    const anketaRows =
      anketaSnapshot?.rows?.length
        ? applyAnketaEditsToRows(anketaSnapshot.rows, anketaEdits)
        : [];
    const anketaById = new Map<string, AnketaRow>();
    const anketaByRnokpp = new Map<string, AnketaRow>();
    const anketaByName = new Map<string, AnketaRow>();
    for (const row of anketaRows) {
      const externalId = String(row.externalId ?? "").trim();
      if (externalId) anketaById.set(externalId, row);
      const rnokpp = String(row.rnokpp ?? "").replace(/\D/g, "");
      if (rnokpp.length >= 8) anketaByRnokpp.set(rnokpp, row);
      const nameKey = normalizeAnketaNameKey(row.fullName);
      if (nameKey) anketaByName.set(nameKey, row);
    }

    const uniquePeople = [
      ...new Map(
        exportDocuments.map((document) => [
          document.personExternalId,
          {
            personId: document.personExternalId,
            fullName: getDocumentPersonName(document),
          },
        ]),
      ).values(),
    ];
    const profileById = new Map<
      string,
      Awaited<ReturnType<typeof api.getPersonnelProfile>> | null
    >();
    const concurrency = 4;
    for (let index = 0; index < uniquePeople.length; index += concurrency) {
      const chunk = uniquePeople.slice(index, index + concurrency);
      const results = await Promise.all(
        chunk.map(async ({ personId, fullName }) => {
          try {
            const profile = await api.getPersonnelProfile(
              personId,
              fullName && !fullName.startsWith("ID ") ? fullName : undefined,
            );
            return [personId, profile] as const;
          } catch {
            return [personId, null] as const;
          }
        }),
      );
      for (const [personId, profile] of results) {
        profileById.set(personId, profile);
      }
    }

    return exportDocuments.map((document) => {
      const journalDocument = liveJournalDocument(document);
      const personName = getDocumentPersonName(journalDocument);
      const profile = profileById.get(journalDocument.personExternalId) ?? null;
      const rosterRow =
        rosterById.get(journalDocument.personExternalId) ||
        rosterByName.get(normalizeName(personName)) ||
        null;
      const rosterCallSign = rosterRow
        ? readRosterColumnValue(rosterRow, 15) ||
          String(rosterRow["позивний"] ?? rosterRow["Позивний"] ?? "").trim()
        : "";

      const fields = (journalDocument.fields || {}) as Record<string, unknown>;
      const rnokppDigits = String(fields.rnokpp ?? "").replace(/\D/g, "");
      const anketaRow =
        anketaById.get(journalDocument.personExternalId) ||
        (rnokppDigits.length >= 8 ? anketaByRnokpp.get(rnokppDigits) : null) ||
        anketaByName.get(normalizeAnketaNameKey(personName)) ||
        null;

      return buildUbdNotSubmittedRowFromDocument({
        document: journalDocument,
        profile,
        personnelStatus: readDocumentPersonStatus(
          journalDocument,
          personStatusById,
        ),
        rosterRow,
        rosterCallSign,
        anketaRow,
      });
    });
  };

  const journalPeriodFilterLabel =
    journalMonthFilter === "ALL"
      ? "Усі місяці"
      : formatJournalMonthLabel(journalMonthFilter);

  const exportDocumentJournal = async () => {
    if (!filteredJournalDocuments.length) {
      setDocumentMessage("Немає рядків для експорту за поточними фільтрами.");
      return;
    }

    setIsExportingDocumentJournal(true);
    try {
      const { fileName, rowCount } = await exportDocumentsJournalExcel({
        rows: filteredJournalDocuments.map((document) => {
          const journalDocument = liveJournalDocument(document);
          return {
            personName: getDocumentPersonName(journalDocument),
            personId: journalDocument.personExternalId,
            personStatus: readDocumentPersonStatus(
              journalDocument,
              personStatusById,
            ),
            documentType: salaryDocumentTypeLabel(journalDocument.type),
            progressPercent: getDocumentProgressPercent(journalDocument),
            status: documentWorkflowStatusLabel(journalDocument),
            note: liveDocumentStatusNote(journalDocument),
            files: getDocumentFileSummary(journalDocument),
            createdAt: journalDocument.createdAt,
            updatedAt: journalDocument.updatedAt,
            taskPeriodEnd: readDocumentUbdTaskPeriodEndLabel(journalDocument),
          };
        }),
        includeUbdExitDate: showUbdExitDateColumn,
        typeFilterLabel:
          JOURNAL_DOCUMENT_TYPE_FILTERS.find(
            (item) => item.value === journalTypeFilter,
          )?.label ?? "Усі документи",
        statusFilterLabel:
          journalStatusFilter === "ALL"
            ? "Усі статуси"
            : journalStatusFilter,
        periodFilterLabel: journalPeriodFilterLabel,
        totalCount: journalDocuments.length,
      });
      setDocumentMessage(`Експортовано журнал: ${rowCount} рядків · ${fileName}`);
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося експортувати Excel: ${error.message}`
          : "Не вдалося експортувати Excel.",
      );
    } finally {
      setIsExportingDocumentJournal(false);
    }
  };

  const exportUbdRegistryTable = async () => {
    if (!filteredJournalDocuments.length) {
      setDocumentMessage("Немає рядків для експорту за поточними фільтрами.");
      return;
    }
    setIsExportingDocumentJournal(true);
    try {
      setDocumentMessage("Готую таблицю УБД «Не подавалися»…");
      const rows = await buildFilteredUbdExportRows();
      const { fileName, rowCount } = await exportUbdNotSubmittedExcel({
        rows,
        periodFilterLabel: journalPeriodFilterLabel,
      });
      setDocumentMessage(
        `Експортовано УБД «Не подавалися»: ${rowCount} рядків · ${fileName}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося експортувати УБД: ${error.message}`
          : "Не вдалося експортувати УБД.",
      );
    } finally {
      setIsExportingDocumentJournal(false);
    }
  };

  const exportUbdStatusTable = async () => {
    if (!filteredJournalDocuments.length) {
      setDocumentMessage("Немає рядків для експорту за поточними фільтрами.");
      return;
    }
    setIsExportingDocumentJournal(true);
    try {
      setDocumentMessage("Готую експорт статусу УБД…");
      const rows = await buildFilteredUbdExportRows();
      const { fileName, rowCount } = await exportUbdStatusExcel({
        rows,
        periodFilterLabel: journalPeriodFilterLabel,
      });
      setDocumentMessage(
        `Експортовано статус УБД: ${rowCount} рядків · ${fileName}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося експортувати статус УБД: ${error.message}`
          : "Не вдалося експортувати статус УБД.",
      );
    } finally {
      setIsExportingDocumentJournal(false);
    }
  };

  useEffect(() => {
    let cancelled = false;
    void getCommanderSignatureTransparent(UBD_COMMANDER_SIGNATURE_SRC).then(
      (processed) => {
        if (!cancelled) setCommanderSignatureSrc(processed);
      },
    );
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    return () => {
      if (questionnairePreviewUrl) {
        revokeQuestionnairePreviewUrl(questionnairePreviewUrl);
      }
    };
  }, [questionnairePreviewUrl]);

  useEffect(() => {
    const rawPerson = window.localStorage.getItem("army-grid:selected-person");
    const savedMode =
      window.localStorage.getItem("army-grid:selected-document-mode") ||
      (requestedDocumentType === "salary-power-attorney"
        ? "salaryPowerAttorney"
        : "default");
    setMode(
      requestedDocumentType === "ubd-report"
        ? "ubdReport"
        : requestedDocumentType === "ubd-restore-report"
          ? "ubdRestoreReport"
          : requestedDocumentType === "form6-report"
            ? "form6Report"
            : requestedDocumentType === "form12-report"
              ? "form12Report"
            : requestedDocumentType === "service-characteristic"
              ? "serviceCharacteristic"
            : requestedDocumentType === "zhbd-certificate"
              ? "zhbdCertificate"
            : requestedDocumentType === "temporary-military-id"
              ? "temporaryMilitaryId"
              : requestedDocumentType === "salary-power-attorney" ||
                  savedMode === "salaryPowerAttorney"
                ? "salaryPowerAttorney"
                : savedMode === "ubdReport"
                  ? "ubdReport"
                  : savedMode === "ubdRestoreReport"
                    ? "ubdRestoreReport"
                    : savedMode === "form6Report"
                      ? "form6Report"
                      : savedMode === "form12Report"
                        ? "form12Report"
                      : savedMode === "serviceCharacteristic"
                        ? "serviceCharacteristic"
                      : savedMode === "zhbdCertificate"
                        ? "zhbdCertificate"
                      : savedMode === "temporaryMilitaryId"
                        ? "temporaryMilitaryId"
                        : "default",
    );

    if (!rawPerson || !isPersonDocumentMode) return;

    try {
      const person = JSON.parse(rawPerson) as EjournalPreviewRow;
      const nextSummary = buildPersonSummary(person);
      const nextPersonId = String(
        nextSummary.externalId ||
          (person.__dbRowId &&
          !isUnstablePersonExternalId(String(person.__dbRowId))
            ? person.__dbRowId
            : ""),
      );
      const allowedIds = new Set(
        [
          nextPersonId,
          nextSummary.externalId,
          ...collectPersonExternalIdCandidates(person),
        ].filter(Boolean),
      );
      if (requestedPersonId && !allowedIds.has(requestedPersonId)) {
        setDocumentMessage(
          `Не знайшов локальні дані службовця ID ${requestedPersonId}. Відкрийте документ із картки особи.`,
        );
        return;
      }
      setSelectedPerson(person);
      const zhbdDefaults = createZhbdCertificateFields(person, nextSummary);
      const storedFull = readStoredSelectedPersonFullPosition();
      const personFull = getPersonFullPositionTitle(person);
      setZhbdCertificateFields({
        ...zhbdDefaults,
        actualFullPosition:
          zhbdDefaults.actualFullPosition.trim() || personFull || storedFull,
      });
      setFields(createDefaultFields(nextSummary));
      setSalaryFields(createSalaryFields(person, nextSummary));
      setUbdFields(createUbdFields(person, nextSummary));
      setForm6Fields(createForm6Fields(person, nextSummary));
      setForm12Fields(createForm12Fields(person, nextSummary));
      setServiceCharacteristicFields(
        createServiceCharacteristicFields(person, nextSummary),
      );
      setUbdRestoreFields(createUbdRestoreFields(person, nextSummary));
      setTicketFields(createTemporaryMilitaryIdFields(nextSummary));
      if (personFull || storedFull) {
        storeSelectedPersonFullPosition(personFull || storedFull);
        setRosterResolvedFullPosition(personFull || storedFull);
      }
      const key = personWorkflowKey(person, nextSummary);
      setWorkflow(loadWorkflowMap()[key] ?? createEmptyWorkflow());
    } catch {
      setSelectedPerson(null);
    }
  }, [isPersonDocumentMode, requestedDocumentType, requestedPersonId]);

  const summary = useMemo(
    () => buildPersonSummary(selectedPerson),
    [selectedPerson],
  );
  const zhbdDisplayedFullPosition =
    zhbdCertificateFields.actualFullPosition?.trim() ||
    rosterResolvedFullPosition ||
    getPersonFullPositionTitle(selectedPerson) ||
    readStoredSelectedPersonFullPosition() ||
    "";

  const selectedDocument = useMemo(
    () =>
      personDocuments.find((document) => document.id === selectedDocumentId) ??
      allPersonDocuments.find((document) => document.id === selectedDocumentId) ??
      null,
    [allPersonDocuments, personDocuments, selectedDocumentId],
  );
  const canMutateSelected = canMutateDocument(selectedDocument);
  const selectedDocumentAuthor = selectedDocument
    ? documentCreatorLabel(selectedDocument)
    : "";
  const foreignReadonlyClass = canMutateSelected ? "" : "is-foreign-readonly";
  const personExternalId = String(
    summary.externalId ||
      (requestedPersonId && !isUnstablePersonExternalId(requestedPersonId)
        ? requestedPersonId
        : "") ||
      (selectedPerson?.__dbRowId &&
      !isUnstablePersonExternalId(String(selectedPerson.__dbRowId))
        ? selectedPerson.__dbRowId
        : "") ||
      requestedPersonId ||
      selectedDocument?.personExternalId ||
      "",
  );

  useEffect(() => {
    if (!isPersonDocumentMode || mode !== "zhbdCertificate") return;

    const lookupName =
      zhbdCertificateFields.fullName.trim() ||
      getPersonDisplayName(selectedPerson) ||
      summary.name;

    const applyFull = (full: string) => {
      const text = full.trim();
      if (!text) return false;
      setRosterResolvedFullPosition(text);
      storeSelectedPersonFullPosition(text);
      setZhbdCertificateFields((current) => {
        const nextStaff = resolveZhbdCombatStaffPosition(current.rank, text);
        const staffLooksGeneric =
          !current.staffPosition.trim() ||
          /^командир\s+(відділення|взводу)$/i.test(current.staffPosition.trim());
        const next = {
          ...current,
          actualFullPosition: text,
          staffPosition: staffLooksGeneric ? nextStaff : current.staffPosition,
        };
        if (staffLooksGeneric || current.actualFullPosition !== text) {
          next.bodyParagraph = buildZhbdBodyParagraph(next);
        }
        return next;
      });
      return true;
    };

    const fromStore = readStoredSelectedPersonFullPosition();
    const fromPerson = getPersonFullPositionTitle(selectedPerson);
    const fromFields = zhbdCertificateFields.actualFullPosition?.trim() || "";
    if (applyFull(fromFields || fromPerson || fromStore)) return;

    if (!lookupName || lookupName === "Особа не вибрана") return;

    let cancelled = false;
    const normalize = (value: string) =>
      value
        .replace(/[ʼ’']/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .toLocaleLowerCase("uk-UA");

    void (async () => {
      try {
        // 1) Той самий ранковий «Загальний список», що на картці особи
        const latest = await fetchWithCache({
          key: CacheKeys.rosterLatest,
          fetcher: () => api.getLatestPersonnelRoster(),
          isChanged: jsonChanged,
        });
        if (cancelled) return;
        const rows = mapRosterLatestToPreviewRows(latest);
        const wanted = normalize(lookupName);
        const wantedSurname = wanted.split(" ")[0] || "";
        const match =
          rows.find((row) => {
            const rowName = normalize(String(row["піб"] ?? row["ПІБ"] ?? ""));
            return rowName && rowName === wanted;
          }) ||
          rows.find((row) => {
            const rowName = normalize(
              getPersonDisplayName(row) || String(row["піб"] ?? ""),
            );
            return rowName && rowName === wanted;
          }) ||
          rows.find((row) => {
            const rowName = normalize(String(row["піб"] ?? ""));
            return (
              wantedSurname.length >= 3 &&
              rowName.startsWith(wantedSurname) &&
              wanted.split(" ").every((part) => !part || rowName.includes(part))
            );
          });

        const fromRosterRow = match
          ? String(
              match["повна_посада"] ??
                match["Повна посада"] ??
                match["roster__повна_посада"] ??
                "",
            ).trim() || getPersonFullPositionTitle(match)
          : "";
        if (applyFull(fromRosterRow)) return;

        // 2) Профіль з БД (roster.values)
        const profileId = personExternalId || summary.externalId;
        if (profileId) {
          const profile = await api.getPersonnelProfile(
            profileId,
            lookupName || undefined,
          );
          if (cancelled) return;
          if (
            applyFull(getFullPositionFromPersonnelProfileRoster(profile?.roster))
          ) {
            return;
          }
        }
      } catch (error) {
        console.warn("[Довідка ЖБД] не вдалося взяти повну посаду з ранкового:", error);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [
    isPersonDocumentMode,
    mode,
    personExternalId,
    selectedPerson,
    summary.externalId,
    summary.name,
    zhbdCertificateFields.actualFullPosition,
    zhbdCertificateFields.fullName,
  ]);

  const workflowKey = useMemo(
    () =>
      [personWorkflowKey(selectedPerson, summary), selectedDocumentId]
        .filter(Boolean)
        .join(":"),
    [selectedDocumentId, selectedPerson, summary],
  );
  const documentSavePersonId =
    selectedDocument?.personExternalId || personExternalId;
  const activeWorkflowSteps = documentWorkflowSteps(
    selectedDocument?.type || mode,
  );
  const targetDocumentType: DocumentMode =
    requestedDocumentType === "ubd-report"
      ? "ubdReport"
      : requestedDocumentType === "ubd-restore-report"
        ? "ubdRestoreReport"
        : requestedDocumentType === "form6-report"
          ? "form6Report"
          : requestedDocumentType === "form12-report"
            ? "form12Report"
          : requestedDocumentType === "service-characteristic"
            ? "serviceCharacteristic"
          : requestedDocumentType === "zhbd-certificate"
            ? "zhbdCertificate"
          : requestedDocumentType === "temporary-military-id"
            ? "temporaryMilitaryId"
            : requestedDocumentType === "salary-power-attorney"
              ? "salaryPowerAttorney"
              : mode;

  useEffect(() => {
    if (!isPersonDocumentMode) return;
    if (
      (targetDocumentType !== "salaryPowerAttorney" &&
        targetDocumentType !== "ubdReport" &&
        targetDocumentType !== "ubdRestoreReport" &&
        targetDocumentType !== "form6Report" &&
        targetDocumentType !== "form12Report" &&
        targetDocumentType !== "serviceCharacteristic" &&
        targetDocumentType !== "zhbdCertificate" &&
        targetDocumentType !== "temporaryMilitaryId") ||
      !personExternalId
    )
      return;

    let cancelled = false;
    const salaryDefaults = createSalaryFields(selectedPerson, summary);

    const loadDocuments = async () => {
      setDocumentMessage("Завантажую документи службовця...");
      setIsQuestionnairePreviewOpen(false);
      setIsLoadingQuestionnairePreview(false);
      setQuestionnaireFileData("");
      setQuestionnairePreviewPersonId("");
      questionnaireLoadSeqRef.current += 1;
      setQuestionnairePreviewUrl((current) => {
        if (current) revokeQuestionnairePreviewUrl(current);
        return "";
      });
      try {
        if (selectedPerson && summary.externalId) {
          const pairs = collectPersonExternalIdCandidates(selectedPerson)
            .filter(
              (fromExternalId) =>
                fromExternalId !== summary.externalId &&
                isUnstablePersonExternalId(fromExternalId),
            )
            .map((fromExternalId) => ({
              name: summary.name,
              fromExternalId,
              toExternalId: summary.externalId,
            }));
          if (pairs.length) {
            migrateStoredPersonSignatures(pairs);
            await migratePersonAttachmentsBetweenIds(pairs, {
              includeDocuments: true,
            });
          }
        }
        const [documents, configuredRecords, personPhoto, questionnaires] =
          await Promise.all([
            api.listPersonDocuments(personExternalId),
            targetDocumentType === "ubdRestoreReport"
              ? loadUbdRestoreSignatoryRecords()
              : targetDocumentType === "ubdReport"
              ? api.listDocumentSignatories("ubdReport")
              : targetDocumentType === "form6Report"
                ? loadForm6SignatoryRecords()
              : targetDocumentType === "form12Report"
                ? loadForm12SignatoryRecords()
              : targetDocumentType === "serviceCharacteristic"
                ? loadServiceCharacteristicSignatoryRecords()
              : targetDocumentType === "zhbdCertificate"
                ? loadZhbdCertificateSignatoryRecords()
                : Promise.resolve([]),
            api.getPersonPhoto(personExternalId).catch(() => null),
            api.listPersonQuestionnaires().catch(() => []),
          ]);
        const questionnaire =
          questionnaires.find(
            (item) => item.personExternalId === personExternalId,
          ) ?? null;
        const configured = snapshotSignatories(configuredRecords);
        const ubdDefaults = createUbdFields(
          selectedPerson,
          summary,
          configured.length ? configured : legacyUbdSignatories(),
        );
        const ticketDefaults = createTemporaryMilitaryIdFields(
          summary,
          personPhoto?.photoData || "",
        );
        const form6Defaults = createForm6Fields(
          selectedPerson,
          summary,
          toForm6Signatories(
            configured.length ? configured : legacyUbdSignatories(),
          ),
        );
        const form12Defaults = createForm12Fields(
          selectedPerson,
          summary,
          toForm12Signatories(
            configured.length ? configured : legacyUbdSignatories(),
          ),
        );
        const serviceCharacteristicDefaults = createServiceCharacteristicFields(
          selectedPerson,
          summary,
          toServiceCharacteristicSignatories(
            configured.length ? configured : legacyUbdSignatories(),
          ),
        );
        const zhbdCertificateDefaults = createZhbdCertificateFields(
          selectedPerson,
          summary,
          toZhbdCertificateSignatories(
            configured.length ? configured : legacyUbdSignatories(),
          ),
        );
        const restoreDefaults = createUbdRestoreFields(
          selectedPerson,
          summary,
          toUbdRestoreSignatories(
            configured.length ? configured : legacyUbdSignatories(),
          ),
        );
        let nextDocuments = documents;
        let active =
          nextDocuments.find(
            (document) => document.id === requestedDocumentId,
          ) ??
          nextDocuments.find(
            (document) => document.type === targetDocumentType,
          ) ??
          null;

        if (cancelled) return;
        setPersonDocuments(nextDocuments);
        setPersonQuestionnaire(questionnaire);
        if (!active) {
          const autoCreateKey = `${personExternalId}:${targetDocumentType}`;
          if (autoCreateInFlight.has(autoCreateKey)) {
            const latest = await api.listPersonDocuments(personExternalId);
            if (cancelled) return;
            nextDocuments = latest;
            active =
              latest.find((document) => document.id === requestedDocumentId) ??
              latest.find((document) => document.type === targetDocumentType) ??
              null;
            setPersonDocuments(nextDocuments);
          }
        }
        if (!active) {
          const autoCreateKey = `${personExternalId}:${targetDocumentType}`;
          const draft = buildPersonDocumentDraft(targetDocumentType, {
            personStatus: createPersonStatusSnapshot(selectedPerson, summary)
              .label,
            salaryFields: salaryDefaults,
            ubdFields: ubdDefaults,
            form6Fields: form6Defaults,
            form12Fields: form12Defaults,
            serviceCharacteristicFields: serviceCharacteristicDefaults,
            zhbdCertificateFields: zhbdCertificateDefaults,
            restoreFields: restoreDefaults,
            ticketFields: ticketDefaults,
          });
          if (!draft) {
            setSelectedDocumentId("");
            setSalaryFields(salaryDefaults);
            setUbdFields(ubdDefaults);
            setForm6Fields(form6Defaults);
            setForm12Fields(form12Defaults);
            setServiceCharacteristicFields(serviceCharacteristicDefaults);
            setZhbdCertificateFields(zhbdCertificateDefaults);
            setUbdRestoreFields(restoreDefaults);
            setTicketFields(ticketDefaults);
            setDocumentFiles({});
            setWorkflow(createEmptyWorkflow());
            setDocumentMessage(
              `Для ID ${personExternalId} ще немає документа "${salaryDocumentTypeLabel(targetDocumentType)}". Створіть його вручну.`,
            );
            return;
          }
          setDocumentMessage(
            `Створюю документ "${salaryDocumentTypeLabel(targetDocumentType)}"...`,
          );
          autoCreateInFlight.add(autoCreateKey);
          try {
            const created = await api.createPersonDocument(
              personExternalId,
              draft,
            );
            if (cancelled) return;
            nextDocuments = [
              created,
              ...nextDocuments.filter((document) => document.id !== created.id),
            ];
            active = created;
            setPersonDocuments(nextDocuments);
            setAllPersonDocuments((current) => [
              created,
              ...current.filter((document) => document.id !== created.id),
            ]);
          } catch (error) {
            if (cancelled) return;
            setSelectedDocumentId("");
            setSalaryFields(salaryDefaults);
            setUbdFields(ubdDefaults);
            setForm6Fields(form6Defaults);
            setForm12Fields(form12Defaults);
            setServiceCharacteristicFields(serviceCharacteristicDefaults);
            setZhbdCertificateFields(zhbdCertificateDefaults);
            setUbdRestoreFields(restoreDefaults);
            setTicketFields(ticketDefaults);
            setDocumentFiles({});
            setWorkflow(createEmptyWorkflow());
            setDocumentMessage(
              error instanceof Error
                ? `Не вдалося створити документ: ${error.message}`
                : `Для ID ${personExternalId} ще немає документа "${salaryDocumentTypeLabel(targetDocumentType)}". Створіть його вручну.`,
            );
            return;
          } finally {
            autoCreateInFlight.delete(autoCreateKey);
          }
        }
        setSelectedDocumentId(active.id);
        setSalaryFields(mergeSalaryFields(salaryDefaults, active.fields));
        const mergedUbdFields = mergeUbdFields(ubdDefaults, active.fields);
        setUbdFields({
          ...mergedUbdFields,
          signatories: ubdDefaults.signatories,
        });
        setForm6Fields(
          withSavedPersonPhone(
            {
              ...mergeForm6Fields(form6Defaults, active.fields),
              signatories: form6Defaults.signatories,
            },
            personExternalId,
          ),
        );
        const personSignature = extractPersonSignature(
          nextDocuments,
          personExternalId,
        ).signature;
        if (personSignature) {
          void persistPersonSignature(
            personExternalId,
            personSignature,
            nextDocuments,
          ).catch(() => undefined);
        }
        const mergedForm12 = withPersonSignature(
          mergeForm12Fields(form12Defaults, active.fields),
          personSignature,
        );
        setForm12Fields({
          ...mergedForm12,
          signatories: form12Defaults.signatories,
        });
        setServiceCharacteristicFields({
          ...mergeServiceCharacteristicFields(
            serviceCharacteristicDefaults,
            active.fields,
          ),
          signatories: serviceCharacteristicDefaults.signatories,
        });
        setZhbdCertificateFields(() => {
          const merged = mergeZhbdCertificateFields(
            zhbdCertificateDefaults,
            active.fields,
          );
          return {
            ...merged,
            actualFullPosition:
              merged.actualFullPosition.trim() ||
              getPersonFullPositionTitle(selectedPerson) ||
              readStoredSelectedPersonFullPosition(),
            signatories: zhbdCertificateDefaults.signatories,
          };
        });
        const mergedRestore = withPersonSignature(
          mergeUbdRestoreFields(restoreDefaults, active.fields),
          personSignature,
        );
        setUbdRestoreFields(mergedRestore);
        const mergedTicket = mergeTemporaryMilitaryIdFields(
          ticketDefaults,
          active.fields,
        );
        const ticketFiles = mergeDocumentFiles(active.files);
        setTicketFields({
          ...mergedTicket,
          photoData:
            mergedTicket.photoData ||
            ticketFiles.ticketPhoto?.dataUrl ||
            ticketDefaults.photoData,
        });
        setDocumentFiles(ticketFiles);
        setWorkflow(mergeSalaryWorkflow(active.workflow));
        setDocumentMessage(
          `Відкрито документ: ${active.title} · ${new Date(active.updatedAt).toLocaleString("uk-UA")}`,
        );
      } catch (error) {
        if (cancelled) return;
        setPersonQuestionnaire(null);
        setDocumentMessage(
          error instanceof Error
            ? `БД документів недоступна: ${error.message}`
            : "БД документів недоступна.",
        );
      }
    };

    void loadDocuments();

    return () => {
      cancelled = true;
    };
  }, [
    isPersonDocumentMode,
    personExternalId,
    requestedDocumentId,
    selectedPerson,
    summary,
    targetDocumentType,
  ]);

  const loadDocumentJournal = useCallback(async (message: string) => {
    setIsLoadingDocumentJournal(true);
    setDocumentMessage(message);
    try {
      const applyJournal = (
        documents: BackendPersonDocument[],
        overview: BackendPersonnelOverview | null,
        fromCache = false,
      ) => {
        setAllPersonDocuments(documents);
        setPersonStatusById(
          Object.fromEntries(
            (overview?.rows ?? [])
              .filter((row) => row.externalId)
              .map((row) => [row.externalId, overviewStatusSnapshot(row)]),
          ),
        );
        setDocumentMessage(
          fromCache
            ? `Кеш журналу: ${documents.length} · оновлюю з БД…`
            : `Документів у журналі: ${documents.length}.`,
        );
      };

      const [cachedDocs, cachedOverview] = await Promise.all([
        readDataCache<BackendPersonDocument[]>(CacheKeys.documentsAll),
        readDataCache<BackendPersonnelOverview>(CacheKeys.overview),
      ]);
      if (cachedDocs) {
        applyJournal(cachedDocs, cachedOverview, true);
        setIsLoadingDocumentJournal(false);
      }

      const [documents, overview] = await Promise.all([
        fetchWithCache({
          key: CacheKeys.documentsAll,
          fetcher: () => api.listAllPersonDocuments(),
          isChanged: jsonChanged,
        }),
        fetchWithCache({
          key: CacheKeys.overview,
          fetcher: () => api.getPersonnelOverview(),
          isChanged: jsonChanged,
        }).catch(() => null),
      ]);
      applyJournal(documents, overview);
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося завантажити журнал: ${error.message}`
          : "Не вдалося завантажити журнал документів.",
      );
    } finally {
      setIsLoadingDocumentJournal(false);
    }
  }, []);

  useEffect(() => {
    if (isPersonDocumentMode) return;

    void loadDocumentJournal("Завантажую журнал документів...");
  }, [isPersonDocumentMode, loadDocumentJournal]);
  const currentStepIndex = Math.max(
    0,
    activeWorkflowSteps.findIndex(
      (step) =>
        step.key ===
        resolveDocumentWorkflowStatus(
          selectedDocument?.type || mode,
          workflow.currentStatus,
        ),
    ),
  );
  const completedCount = currentStepIndex + 1;
  const progressPercent = Math.round(
    (completedCount / activeWorkflowSteps.length) * 100,
  );

  const updateDefaultField = (key: DefaultDocumentFieldKey, value: string) => {
    setFields((current) => ({ ...current, [key]: value }));
  };

  const openPersonDocument = async (document: BackendPersonDocument) => {
    const nextPersonId = String(document.personExternalId || "").trim();
    if (
      nextPersonId &&
      nextPersonId !== questionnairePreviewPersonId
    ) {
      questionnaireLoadSeqRef.current += 1;
      setIsQuestionnairePreviewOpen(false);
      setIsLoadingQuestionnairePreview(false);
      setQuestionnaireFileData("");
      setQuestionnairePreviewPersonId("");
      setQuestionnairePreviewUrl((current) => {
        if (current) revokeQuestionnairePreviewUrl(current);
        return "";
      });
    }

    setSelectedDocumentId(document.id);
    setPersonDocuments((current) =>
      current.some((item) => item.id === document.id)
        ? current
        : [document, ...current],
    );
    void api
      .getPersonQuestionnaire(document.personExternalId)
      .then((item) => {
        setPersonQuestionnaire(
          item?.personExternalId
            ? {
                personExternalId: item.personExternalId,
                fileName: item.fileName,
              }
            : null,
        );
      })
      .catch(() => {
        setPersonQuestionnaire(null);
      });

    if (document.type === "ubdReport") {
      setMode("ubdReport");
      const configured = snapshotSignatories(
        await api.listDocumentSignatories("ubdReport"),
      );
      const defaults = createUbdFields(
        selectedPerson,
        summary,
        configured.length ? configured : legacyUbdSignatories(),
      );
      const merged = mergeUbdFields(defaults, document.fields);
      setUbdFields({ ...merged, signatories: defaults.signatories });
      setDocumentFiles(mergeDocumentFiles(document.files));
    } else if (document.type === "form6Report") {
      setMode("form6Report");
      const configured = snapshotSignatories(await loadForm6SignatoryRecords());
      const defaults = createForm6Fields(
        selectedPerson,
        summary,
        toForm6Signatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      setForm6Fields(
        withSavedPersonPhone(
          {
            ...mergeForm6Fields(defaults, document.fields),
            signatories: defaults.signatories,
          },
          nextPersonId || personExternalId,
        ),
      );
      setDocumentFiles(mergeDocumentFiles(document.files));
    } else if (document.type === "form12Report") {
      setMode("form12Report");
      const configured = snapshotSignatories(await loadForm12SignatoryRecords());
      const defaults = createForm12Fields(
        selectedPerson,
        summary,
        toForm12Signatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      const personDocs =
        personDocuments.some((item) => item.id === document.id)
          ? personDocuments
          : [document, ...personDocuments];
      const personSignature = extractPersonSignature(
        personDocs,
        nextPersonId || personExternalId,
      ).signature;
      const merged = withPersonSignature(
        mergeForm12Fields(defaults, document.fields),
        personSignature,
      );
      setForm12Fields({ ...merged, signatories: defaults.signatories });
      setDocumentFiles(mergeDocumentFiles(document.files));
    } else if (document.type === "serviceCharacteristic") {
      setMode("serviceCharacteristic");
      const configured = snapshotSignatories(
        await loadServiceCharacteristicSignatoryRecords(),
      );
      const defaults = createServiceCharacteristicFields(
        selectedPerson,
        summary,
        toServiceCharacteristicSignatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      setServiceCharacteristicFields({
        ...mergeServiceCharacteristicFields(defaults, document.fields),
        signatories: defaults.signatories,
      });
      setDocumentFiles(mergeDocumentFiles(document.files));
    } else if (document.type === "zhbdCertificate") {
      setMode("zhbdCertificate");
      const configured = snapshotSignatories(
        await loadZhbdCertificateSignatoryRecords(),
      );
      const defaults = createZhbdCertificateFields(
        selectedPerson,
        summary,
        toZhbdCertificateSignatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      setZhbdCertificateFields(() => {
        const merged = mergeZhbdCertificateFields(defaults, document.fields);
        return {
          ...merged,
          actualFullPosition:
            merged.actualFullPosition.trim() ||
            getPersonFullPositionTitle(selectedPerson) ||
            readStoredSelectedPersonFullPosition(),
          signatories: defaults.signatories,
        };
      });
      setDocumentFiles(mergeDocumentFiles(document.files));
    } else if (document.type === "ubdRestoreReport") {
      setMode("ubdRestoreReport");
      const configured = snapshotSignatories(
        await loadUbdRestoreSignatoryRecords(),
      );
      const defaults = createUbdRestoreFields(
        selectedPerson,
        summary,
        toUbdRestoreSignatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      const personDocs =
        personDocuments.some((item) => item.id === document.id)
          ? personDocuments
          : [document, ...personDocuments];
      const personSignature = extractPersonSignature(
        personDocs,
        nextPersonId || personExternalId,
      ).signature;
      setUbdRestoreFields(
        withPersonSignature(
          mergeUbdRestoreFields(defaults, document.fields),
          personSignature,
        ),
      );
      setDocumentFiles(mergeDocumentFiles(document.files));
    } else if (document.type === "temporaryMilitaryId") {
      const ticketPersonId = document.personExternalId || personExternalId;
      const personPhoto = ticketPersonId
        ? await api.getPersonPhoto(ticketPersonId).catch(() => null)
        : null;
      const defaults = createTemporaryMilitaryIdFields(
        summary,
        personPhoto?.photoData || "",
      );
      const merged = mergeTemporaryMilitaryIdFields(defaults, document.fields);
      const files = mergeDocumentFiles(document.files);
      setMode("temporaryMilitaryId");
      setTicketFields({
        ...merged,
        photoData: merged.photoData || files.ticketPhoto?.dataUrl || defaults.photoData,
      });
      setDocumentFiles(files);
    } else {
      const defaults = createSalaryFields(selectedPerson, summary);
      setMode("salaryPowerAttorney");
      setSalaryFields(mergeSalaryFields(defaults, document.fields));
      setDocumentFiles({});
    }

    setWorkflow(mergeSalaryWorkflow(document.workflow));
    setDocumentMessage(
      `Відкрито документ: ${document.title} · ${new Date(document.updatedAt).toLocaleString("uk-UA")}`,
    );
  };

  const createSalaryPowerAttorneyDocument = async () => {
    if (!personExternalId) return;
    const defaults = createSalaryFields(selectedPerson, summary);
    const fieldPayload = {
      ...defaults,
      personStatus: createPersonStatusSnapshot(selectedPerson, summary).label,
    };
    const nextWorkflow = createEmptyWorkflow();
    setIsSavingDocument(true);
    setDocumentMessage("Створюю документ службовця...");
    try {
      const created = await api.createPersonDocument(personExternalId, {
        type: "salaryPowerAttorney",
        title: `Довіреність зарплати · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: nextWorkflow.currentStatus,
        fields: fieldPayload as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
      });
      setPersonDocuments((current) => [created, ...current]);
      setAllPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setMode("salaryPowerAttorney");
      setSalaryFields(mergeSalaryFields(defaults, created.fields));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити документ: ${error.message}`
          : "Не вдалося створити документ.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const createUbdReportDocument = async () => {
    if (!personExternalId) return;
    const nextWorkflow = {
      ...createEmptyWorkflow(),
      currentStatus: "document",
    };
    setIsSavingDocument(true);
    setDocumentMessage("Створюю УБД рапорт...");
    try {
      const configured = snapshotSignatories(
        await api.listDocumentSignatories("ubdReport"),
      );
      const defaults = createUbdFields(
        selectedPerson,
        summary,
        configured.length ? configured : legacyUbdSignatories(),
      );
      const fieldPayload = {
        ...defaults,
        personStatus: createPersonStatusSnapshot(selectedPerson, summary).label,
      };
      const created = await api.createPersonDocument(personExternalId, {
        type: "ubdReport",
        title: `Рапорт на УБД · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: "document",
        fields: fieldPayload as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
        files: { ubdScans: [] },
      });
      setPersonDocuments((current) => [created, ...current]);
      setAllPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setMode("ubdReport");
      setUbdFields(mergeUbdFields(defaults, created.fields));
      setDocumentFiles(mergeDocumentFiles(created.files));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити УБД рапорт: ${error.message}`
          : "Не вдалося створити УБД рапорт.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const createForm6ReportDocument = async () => {
    if (!personExternalId) return;
    const nextWorkflow = {
      ...createEmptyWorkflow(),
      currentStatus: "document",
    };
    setIsSavingDocument(true);
    setDocumentMessage("Створюю Форму 6...");
    try {
      const configured = snapshotSignatories(await loadForm6SignatoryRecords());
      const defaults = createForm6Fields(
        selectedPerson,
        summary,
        toForm6Signatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      const seeded = withSavedPersonPhone(defaults, personExternalId);
      const fieldPayload = {
        ...seeded,
        personStatus: createPersonStatusSnapshot(selectedPerson, summary).label,
      };
      const created = await api.createPersonDocument(personExternalId, {
        type: "form6Report",
        title: `Форма 6 · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: "document",
        fields: fieldPayload as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
        files: { ubdScans: [] },
      });
      setPersonDocuments((current) => [created, ...current]);
      setAllPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setForm6Fields(
        withSavedPersonPhone(
          mergeForm6Fields(seeded, created.fields),
          personExternalId,
        ),
      );
      setDocumentFiles(mergeDocumentFiles(created.files));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setMode("form6Report");
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити Форму 6: ${error.message}`
          : "Не вдалося створити Форму 6.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };


  const createServiceCharacteristicDocument = async () => {
    if (!personExternalId) return;
    const nextWorkflow = {
      ...createEmptyWorkflow(),
      currentStatus: "created",
    };
    setIsSavingDocument(true);
    setDocumentMessage("Створюю службову характеристику...");
    try {
      const configured = snapshotSignatories(
        await loadServiceCharacteristicSignatoryRecords(),
      );
      const defaults = createServiceCharacteristicFields(
        selectedPerson,
        summary,
        toServiceCharacteristicSignatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      const fieldPayload = {
        ...defaults,
        personStatus: createPersonStatusSnapshot(selectedPerson, summary).label,
      };
      const created = await api.createPersonDocument(personExternalId, {
        type: "serviceCharacteristic",
        title: `Службова характеристика · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: "created",
        fields: fieldPayload as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
        files: { ubdScans: [] },
      });
      setPersonDocuments((current) => [created, ...current]);
      setAllPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setServiceCharacteristicFields({
        ...mergeServiceCharacteristicFields(defaults, created.fields),
        signatories: defaults.signatories,
      });
      setDocumentFiles(mergeDocumentFiles(created.files));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setMode("serviceCharacteristic");
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити службову характеристику: ${error.message}`
          : "Не вдалося створити службову характеристику.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const createZhbdCertificateDocument = async () => {
    if (!personExternalId) return;
    const nextWorkflow = {
      ...createEmptyWorkflow(),
      currentStatus: "created",
    };
    setIsSavingDocument(true);
    setDocumentMessage("Створюю довідку ЖБД...");
    try {
      const configured = snapshotSignatories(
        await loadZhbdCertificateSignatoryRecords(),
      );
      const defaults = createZhbdCertificateFields(
        selectedPerson,
        summary,
        toZhbdCertificateSignatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      const fieldPayload = {
        ...defaults,
        personStatus: createPersonStatusSnapshot(selectedPerson, summary).label,
      };
      const created = await api.createPersonDocument(personExternalId, {
        type: "zhbdCertificate",
        title: `Довідка ЖБД · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: "created",
        fields: fieldPayload as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
        files: { ubdScans: [] },
      });
      setPersonDocuments((current) => [created, ...current]);
      setAllPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setZhbdCertificateFields({
        ...mergeZhbdCertificateFields(defaults, created.fields),
        // Always use fresh presets (API may strip large signatureData on create).
        signatories: defaults.signatories,
      });
      setDocumentFiles(mergeDocumentFiles(created.files));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setMode("zhbdCertificate");
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити довідку ЖБД: ${error.message}`
          : "Не вдалося створити довідку ЖБД.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const createForm12ReportDocument = async () => {
    if (!personExternalId) return;
    const nextWorkflow = {
      ...createEmptyWorkflow(),
      currentStatus: "document",
    };
    setIsSavingDocument(true);
    setDocumentMessage("Створюю Форму 12...");
    try {
      const configured = snapshotSignatories(await loadForm12SignatoryRecords());
      const defaults = createForm12Fields(
        selectedPerson,
        summary,
        toForm12Signatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      const fieldPayload = {
        ...defaults,
        personStatus: createPersonStatusSnapshot(selectedPerson, summary).label,
      };
      const created = await api.createPersonDocument(personExternalId, {
        type: "form12Report",
        title: `Форма 12 · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: "document",
        fields: fieldPayload as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
        files: { ubdScans: [] },
      });
      setPersonDocuments((current) => [created, ...current]);
      setAllPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setForm12Fields(mergeForm12Fields(defaults, created.fields));
      setDocumentFiles(mergeDocumentFiles(created.files));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setMode("form12Report");
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити Форму 12: ${error.message}`
          : "Не вдалося створити Форму 12.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const createUbdRestoreReportDocument = async () => {
    if (!personExternalId) return;
    const nextWorkflow = {
      ...createEmptyWorkflow(),
      currentStatus: "document",
    };
    setIsSavingDocument(true);
    setDocumentMessage("Створюю рапорт на відновлення УБД...");
    try {
      const configured = snapshotSignatories(
        await loadUbdRestoreSignatoryRecords(),
      );
      const defaults = createUbdRestoreFields(
        selectedPerson,
        summary,
        toUbdRestoreSignatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      const fieldPayload = {
        ...defaults,
        personStatus: createPersonStatusSnapshot(selectedPerson, summary).label,
      };
      const created = await api.createPersonDocument(personExternalId, {
        type: "ubdRestoreReport",
        title: `Рапорт на відновлення УБД · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: "document",
        fields: fieldPayload as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
        files: { ubdScans: [] },
      });
      setPersonDocuments((current) => [created, ...current]);
      setAllPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setUbdRestoreFields(mergeUbdRestoreFields(defaults, created.fields));
      setDocumentFiles(mergeDocumentFiles(created.files));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setMode("ubdRestoreReport");
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити рапорт на відновлення УБД: ${error.message}`
          : "Не вдалося створити рапорт на відновлення УБД.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const createTemporaryMilitaryIdDocument = async () => {
    if (!personExternalId) return;
    const nextWorkflow = {
      ...createEmptyWorkflow(),
      currentStatus: "photo",
    };
    setIsSavingDocument(true);
    setDocumentMessage("Створюю тимчасовий військовий квиток...");
    try {
      const personPhoto = await api
        .getPersonPhoto(personExternalId)
        .catch(() => null);
      const defaults = createTemporaryMilitaryIdFields(
        summary,
        personPhoto?.photoData || "",
      );
      const ticketPhoto = defaults.photoData
        ? {
            id: `photo-${Date.now()}`,
            name: "фото.jpg",
            type: "image/jpeg",
            dataUrl: defaults.photoData,
            uploadedAt: new Date().toISOString(),
          }
        : null;
      const nextFiles = { ticketPhoto, ubdScans: [] };
      const fieldPayload = {
        ...defaults,
        personStatus: createPersonStatusSnapshot(selectedPerson, summary).label,
      };
      const created = await api.createPersonDocument(personExternalId, {
        type: "temporaryMilitaryId",
        title: `Тимчасовий військовий квиток · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: "photo",
        fields: fieldPayload as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
        files: nextFiles as unknown as Record<string, unknown>,
      });
      setPersonDocuments((current) => [created, ...current]);
      setAllPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setMode("temporaryMilitaryId");
      setTicketFields(mergeTemporaryMilitaryIdFields(defaults, created.fields));
      setDocumentFiles(mergeDocumentFiles(created.files));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити квиток: ${error.message}`
          : "Не вдалося створити квиток.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const addUbdScanFiles = async (files: FileList | null) => {
    if (!files?.length || !selectedDocument) return;

    setIsSavingDocument(true);
    try {
      const uploaded = await Promise.all(
        Array.from(files).map(async (file) => ({
          id: `${Date.now()}-${file.name}-${Math.random().toString(16).slice(2)}`,
          name: file.name,
          type: file.type || "application/octet-stream",
          dataUrl: await readFileAsDataUrl(file),
          uploadedAt: new Date().toISOString(),
        })),
      );
      const nextFiles = {
        ...documentFiles,
        ubdScans: [...(documentFiles.ubdScans ?? []), ...uploaded],
      };
      setDocumentFiles(nextFiles);
      if (selectedDocument.type === "form6Report" || mode === "form6Report") {
        await saveForm6Document(form6Fields, nextFiles);
      } else if (selectedDocument.type === "form12Report" || mode === "form12Report") {
        await saveForm12Document(form12Fields, nextFiles);
      } else if (
        selectedDocument.type === "serviceCharacteristic" ||
        mode === "serviceCharacteristic"
      ) {
        await saveServiceCharacteristicDocument(
          serviceCharacteristicFields,
          nextFiles,
        );
      } else if (
        selectedDocument.type === "zhbdCertificate" ||
        mode === "zhbdCertificate"
      ) {
        await saveZhbdCertificateDocument(zhbdCertificateFields, nextFiles);
      } else if (
        selectedDocument.type === "ubdRestoreReport" ||
        mode === "ubdRestoreReport"
      ) {
        await saveUbdRestoreDocument(ubdRestoreFields, nextFiles);
      } else if (
        selectedDocument.type === "temporaryMilitaryId" ||
        mode === "temporaryMilitaryId"
      ) {
        await saveTicketDocument(ticketFields, nextFiles);
      } else {
        await saveUbdDocument(ubdFields, nextFiles);
      }
      window.requestAnimationFrame(() => {
        document
          .getElementById(
            selectedDocument.type === "temporaryMilitaryId"
              ? "ticket-tvk-scans"
              : "ubd-scans-panel",
          )
          ?.scrollIntoView({
          block: "end",
          behavior: "smooth",
        });
      });
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося додати скани: ${error.message}`
          : "Не вдалося додати скани.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const deleteUbdScanFile = (fileId: string) => {
    const nextFiles = {
      ...documentFiles,
      ubdScans: (documentFiles.ubdScans ?? []).filter(
        (file) => file.id !== fileId,
      ),
    };
    setDocumentFiles(nextFiles);
    if (selectedDocument?.type === "form6Report" || mode === "form6Report") {
      void saveForm6Document(form6Fields, nextFiles);
    } else if (selectedDocument?.type === "form12Report" || mode === "form12Report") {
      void saveForm12Document(form12Fields, nextFiles);
    } else if (
      selectedDocument?.type === "serviceCharacteristic" ||
      mode === "serviceCharacteristic"
    ) {
      void saveServiceCharacteristicDocument(
        serviceCharacteristicFields,
        nextFiles,
      );
    } else if (
      selectedDocument?.type === "zhbdCertificate" ||
      mode === "zhbdCertificate"
    ) {
      void saveZhbdCertificateDocument(zhbdCertificateFields, nextFiles);
    } else if (
      selectedDocument?.type === "ubdRestoreReport" ||
      mode === "ubdRestoreReport"
    ) {
      void saveUbdRestoreDocument(ubdRestoreFields, nextFiles);
    } else if (
      selectedDocument?.type === "temporaryMilitaryId" ||
      mode === "temporaryMilitaryId"
    ) {
      void saveTicketDocument(ticketFields, nextFiles);
    } else {
      void saveUbdDocument(ubdFields, nextFiles);
    }
  };

  const nextPersonStatusLabel = () => {
    const snapshot = createPersonStatusSnapshot(selectedPerson, summary).label;
    const existing =
      typeof selectedDocument?.fields?.personStatus === "string"
        ? selectedDocument.fields.personStatus.trim()
        : "";
    if (selectedPerson && snapshot && snapshot !== "—") return snapshot;
    return existing || snapshot;
  };

  const applyUpdatedDocument = (updated: BackendPersonDocument) => {
    setPersonDocuments((current) =>
      current.some((document) => document.id === updated.id)
        ? current.map((document) =>
            document.id === updated.id ? updated : document,
          )
        : [updated, ...current],
    );
    setAllPersonDocuments((current) =>
      current.some((document) => document.id === updated.id)
        ? current.map((document) =>
            document.id === updated.id ? updated : document,
          )
        : [updated, ...current],
    );
  };

  const patchSelectedDocumentWorkflow = (nextWorkflow: SalaryWorkflowState) => {
    if (!selectedDocumentId) return;
    const patch = (document: BackendPersonDocument) =>
      document.id === selectedDocumentId
        ? {
            ...document,
            status: nextWorkflow.currentStatus,
            workflow: nextWorkflow as unknown as Record<string, unknown>,
          }
        : document;
    setPersonDocuments((current) => current.map(patch));
    setAllPersonDocuments((current) => current.map(patch));
  };

  const saveSalaryDocument = async (
    nextFields: SalaryDocumentFields,
    nextWorkflow: SalaryWorkflowState,
  ) => {
    if (!canMutateSelected) return;
    if (!documentSavePersonId || !selectedDocumentId) {
      saveWorkflowForPerson(workflowKey, nextWorkflow);
      return;
    }

    setIsSavingDocument(true);
    try {
      const fieldPayload = {
        ...nextFields,
        personStatus: nextPersonStatusLabel(),
      };
      const updated = await api.updatePersonDocument(
        documentSavePersonId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Довіреність зарплати",
          status: nextWorkflow.currentStatus,
          fields: fieldPayload as unknown as Record<string, unknown>,
          workflow: nextWorkflow as unknown as Record<string, unknown>,
        },
      );
      applyUpdatedDocument(updated);
      void syncEnrichmentToPerson({
        personExternalId: documentSavePersonId,
        rowId: selectedPerson?.__dbRowId
          ? String(selectedPerson.__dbRowId)
          : null,
        row: selectedPerson,
        patch: { rnokpp: nextFields.rnokpp },
        existingPhones: readStoredPersonPhones()[documentSavePersonId] ?? [],
      })
        .then((enrichment) => {
          if (
            enrichment &&
            Object.keys(enrichment.fieldUpdates).length &&
            selectedPerson
          ) {
            setSelectedPerson(
              applyEnrichmentToPreviewRow(
                selectedPerson,
                enrichment.fieldUpdates,
              ),
            );
          }
        })
        .catch(() => undefined);
      setDocumentMessage(
        `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      saveWorkflowForPerson(workflowKey, nextWorkflow);
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося зберегти в БД, залишив локально: ${error.message}`
          : "Не вдалося зберегти в БД, залишив локально.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const saveUbdDocument = async (
    nextFields: UbdReportFields,
    nextFiles = documentFiles,
    nextWorkflow = workflow,
  ) => {
    if (!canMutateSelected || !documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = {
        ...nextFields,
        personStatus: nextPersonStatusLabel(),
      };
      const updated = await api.updatePersonDocument(
        documentSavePersonId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Рапорт на УБД",
          status: nextWorkflow.currentStatus,
          fields: fieldPayload as unknown as Record<string, unknown>,
          workflow: nextWorkflow as unknown as Record<string, unknown>,
          files: nextFiles as unknown as Record<string, unknown>,
        },
      );
      applyUpdatedDocument(updated);
      void syncEnrichmentToPerson({
        personExternalId: documentSavePersonId,
        rowId: selectedPerson?.__dbRowId
          ? String(selectedPerson.__dbRowId)
          : null,
        row: selectedPerson,
        patch: { rnokpp: nextFields.rnokpp },
        existingPhones: readStoredPersonPhones()[documentSavePersonId] ?? [],
      })
        .then((enrichment) => {
          if (
            enrichment &&
            Object.keys(enrichment.fieldUpdates).length &&
            selectedPerson
          ) {
            setSelectedPerson(
              applyEnrichmentToPreviewRow(
                selectedPerson,
                enrichment.fieldUpdates,
              ),
            );
          }
        })
        .catch(() => undefined);
      setDocumentMessage(
        `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося зберегти УБД рапорт: ${error.message}`
          : "Не вдалося зберегти УБД рапорт.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const saveForm6Document = async (
    nextFields: Form6ReportFields,
    nextFiles = documentFiles,
    nextWorkflow = workflow,
  ) => {
    if (!canMutateSelected || !documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = {
        ...nextFields,
        personStatus: nextPersonStatusLabel(),
      };
      const updated = await api.updatePersonDocument(
        documentSavePersonId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Форма 6",
          status: nextWorkflow.currentStatus,
          fields: fieldPayload as unknown as Record<string, unknown>,
          workflow: nextWorkflow as unknown as Record<string, unknown>,
          files: nextFiles as unknown as Record<string, unknown>,
        },
      );
      applyUpdatedDocument(updated);

      const enrichment = await syncEnrichmentToPerson({
        personExternalId: documentSavePersonId,
        rowId: selectedPerson?.__dbRowId
          ? String(selectedPerson.__dbRowId)
          : null,
        row: selectedPerson,
        patch: {
          rnokpp: nextFields.rnokpp,
          address: nextFields.address,
          phone: nextFields.phone,
        },
        existingPhones: readStoredPersonPhones()[documentSavePersonId] ?? [],
      }).catch(() => null);

      if (
        enrichment &&
        Object.keys(enrichment.fieldUpdates).length &&
        selectedPerson
      ) {
        setSelectedPerson(
          applyEnrichmentToPreviewRow(selectedPerson, enrichment.fieldUpdates),
        );
      }

      const enrichmentNote = enrichment
        ? [
            enrichment.phonesAdded.length
              ? `телефон +${enrichment.phonesAdded.length}`
              : "",
            Object.keys(enrichment.fieldUpdates).length
              ? "порожні поля ОС доповнено"
              : "",
          ]
            .filter(Boolean)
            .join(", ")
        : "";

      setDocumentMessage(
        enrichmentNote
          ? `Збережено в БД · ${enrichmentNote} · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`
          : `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося зберегти Форму 6: ${error.message}`
          : "Не вдалося зберегти Форму 6.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };


  const saveServiceCharacteristicDocument = async (
    nextFields: ServiceCharacteristicFields,
    nextFiles = documentFiles,
    nextWorkflow = workflow,
  ) => {
    if (!canMutateSelected || !documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = {
        ...nextFields,
        personStatus: nextPersonStatusLabel(),
      };
      const updated = await api.updatePersonDocument(
        documentSavePersonId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Службова характеристика",
          status: nextWorkflow.currentStatus,
          fields: fieldPayload as unknown as Record<string, unknown>,
          workflow: nextWorkflow as unknown as Record<string, unknown>,
          files: nextFiles as unknown as Record<string, unknown>,
        },
      );
      applyUpdatedDocument(updated);
      setDocumentMessage(
        `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося зберегти службову характеристику: ${error.message}`
          : "Не вдалося зберегти службову характеристику.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const saveZhbdCertificateDocument = async (
    nextFields: ZhbdCertificateFields,
    nextFiles = documentFiles,
    nextWorkflow = workflow,
  ) => {
    if (!canMutateSelected || !documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = {
        ...nextFields,
        personStatus: nextPersonStatusLabel(),
      };
      const updated = await api.updatePersonDocument(
        documentSavePersonId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Довідка ЖБД",
          status: nextWorkflow.currentStatus,
          fields: fieldPayload as unknown as Record<string, unknown>,
          workflow: nextWorkflow as unknown as Record<string, unknown>,
          files: nextFiles as unknown as Record<string, unknown>,
        },
      );
      applyUpdatedDocument(updated);
      setDocumentMessage(
        `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося зберегти довідку ЖБД: ${error.message}`
          : "Не вдалося зберегти довідку ЖБД.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const saveForm12Document = async (
    nextFields: Form12ReportFields,
    nextFiles = documentFiles,
    nextWorkflow = workflow,
  ) => {
    if (!canMutateSelected || !documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = {
        ...nextFields,
        personStatus: nextPersonStatusLabel(),
      };
      const updated = await api.updatePersonDocument(
        documentSavePersonId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Форма 12",
          status: nextWorkflow.currentStatus,
          fields: fieldPayload as unknown as Record<string, unknown>,
          workflow: nextWorkflow as unknown as Record<string, unknown>,
          files: nextFiles as unknown as Record<string, unknown>,
        },
      );
      applyUpdatedDocument(updated);
      setDocumentMessage(
        `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося зберегти Форму 12: ${error.message}`
          : "Не вдалося зберегти Форму 12.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const saveUbdRestoreDocument = async (
    nextFields: UbdRestoreReportFields,
    nextFiles = documentFiles,
    nextWorkflow = workflow,
  ) => {
    if (!canMutateSelected || !documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = {
        ...nextFields,
        personStatus: nextPersonStatusLabel(),
      };
      const updated = await api.updatePersonDocument(
        documentSavePersonId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Рапорт на відновлення УБД",
          status: nextWorkflow.currentStatus,
          fields: fieldPayload as unknown as Record<string, unknown>,
          workflow: nextWorkflow as unknown as Record<string, unknown>,
          files: nextFiles as unknown as Record<string, unknown>,
        },
      );
      applyUpdatedDocument(updated);
      setDocumentMessage(
        `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося зберегти рапорт на відновлення УБД: ${error.message}`
          : "Не вдалося зберегти рапорт на відновлення УБД.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const saveTicketDocument = async (
    nextFields: TemporaryMilitaryIdFields,
    nextFiles = documentFiles,
    nextWorkflow = workflow,
  ) => {
    if (!canMutateSelected || !documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = {
        ...nextFields,
        personStatus: nextPersonStatusLabel(),
      };
      const updated = await api.updatePersonDocument(
        documentSavePersonId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Тимчасовий військовий квиток",
          status: nextWorkflow.currentStatus,
          fields: fieldPayload as unknown as Record<string, unknown>,
          workflow: nextWorkflow as unknown as Record<string, unknown>,
          files: nextFiles as unknown as Record<string, unknown>,
        },
      );
      applyUpdatedDocument(updated);
      setDocumentMessage(
        `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося зберегти квиток: ${error.message}`
          : "Не вдалося зберегти квиток.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const flushDocumentFieldSave = () => {
    if (autoSaveTimerRef.current) {
      clearTimeout(autoSaveTimerRef.current);
      autoSaveTimerRef.current = null;
    }
    const pending = pendingFieldSaveRef.current;
    pendingFieldSaveRef.current = null;
    pending?.();
  };

  const scheduleDocumentFieldSave = (run: () => void) => {
    pendingFieldSaveRef.current = run;
    if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);
    autoSaveTimerRef.current = setTimeout(() => {
      autoSaveTimerRef.current = null;
      const pending = pendingFieldSaveRef.current;
      pendingFieldSaveRef.current = null;
      pending?.();
    }, 700);
  };

  useEffect(() => {
    return () => {
      flushDocumentFieldSave();
    };
  }, [selectedDocumentId]);

  const updateUbdField = (key: keyof UbdReportFields, value: string) => {
    setUbdFields((current) => {
      const next = { ...current, [key]: value };
      scheduleDocumentFieldSave(() => {
        void saveUbdDocument(next);
      });
      return next;
    });
  };

  const updateUbdBasisPart = (
    key: "basisNumber" | "basisDate",
    value: string,
  ) => {
    setUbdFields((current) => {
      const basisNumber =
        key === "basisNumber" ? value : current.basisNumber;
      const basisDate = key === "basisDate" ? value : current.basisDate;
      const next = {
        ...current,
        basisNumber,
        basisDate,
        basis: formatUbdBasisText(basisNumber, basisDate),
      };
      scheduleDocumentFieldSave(() => {
        void saveUbdDocument(next);
      });
      return next;
    });
  };

  const updateForm6Field = (
    key: Exclude<keyof Form6ReportFields, "signatories">,
    value: string,
  ) => {
    setForm6Fields((current) => {
      const next = { ...current, [key]: value };
      scheduleDocumentFieldSave(() => {
        void saveForm6Document(next);
      });
      return next;
    });
  };


  const updateServiceCharacteristicField = (
    key: Exclude<keyof ServiceCharacteristicFields, "signatories">,
    value: string,
  ) => {
    setServiceCharacteristicFields((current) => {
      const next = { ...current, [key]: value };
      scheduleDocumentFieldSave(() => {
        void saveServiceCharacteristicDocument(next);
      });
      return next;
    });
  };

  const updateZhbdCertificateField = (
    key: Exclude<keyof ZhbdCertificateFields, "signatories">,
    value: string,
  ) => {
    setZhbdCertificateFields((current) => {
      const next = { ...current, [key]: value };
      if (key === "rank") {
        next.staffPosition = resolveZhbdCombatStaffPosition(
          value,
          current.staffPosition,
        );
      }
      if (
        key === "rank" ||
        key === "fullName" ||
        key === "staffPosition" ||
        key === "periodFrom" ||
        key === "periodTo" ||
        key === "periodNote"
      ) {
        next.bodyParagraph = buildZhbdBodyParagraph(next);
      }
      scheduleDocumentFieldSave(() => {
        void saveZhbdCertificateDocument(next);
      });
      return next;
    });
  };

  const updateForm12Field = (
    key: Exclude<keyof Form12ReportFields, "signatories">,
    value: string,
  ) => {
    setForm12Fields((current) => {
      const next = { ...current, [key]: value };
      scheduleDocumentFieldSave(() => {
        void saveForm12Document(next);
      });
      return next;
    });
  };

  const savePersonSignatureForCurrent = async (
    signature: PersonSignatureRecord | null,
  ) => {
    const personId = (documentSavePersonId || personExternalId || "").trim();
    if (!personId) return;
    try {
      const saved = await persistPersonSignature(
        personId,
        signature,
        personDocuments,
      );
      if (saved) {
        setPersonDocuments((current) => [
          saved,
          ...current.filter(
            (item) =>
              item.id !== saved.id &&
              item.type !== PERSON_SIGNATURE_DOCUMENT_TYPE,
          ),
        ]);
        setAllPersonDocuments((current) => [
          saved,
          ...current.filter(
            (item) =>
              item.id !== saved.id &&
              !(
                item.type === PERSON_SIGNATURE_DOCUMENT_TYPE &&
                item.personExternalId === personId
              ),
          ),
        ]);
      } else if (!signature) {
        setPersonDocuments((current) =>
          current.filter((item) => item.type !== PERSON_SIGNATURE_DOCUMENT_TYPE),
        );
        setAllPersonDocuments((current) =>
          current.filter(
            (item) =>
              !(
                item.type === PERSON_SIGNATURE_DOCUMENT_TYPE &&
                item.personExternalId === personId
              ),
          ),
        );
      }
    } catch {
      // Local storage already updated inside persistPersonSignature.
    }
  };

  const uploadForm12Signature = async (file: File | null) => {
    if (!file) return;
    try {
      const dataUrl = await readFileAsDataUrl(file);
      const signatureData = await processSignatureTransparentBackground(dataUrl);
      const signatureFileName = file.name.replace(/\.[^.]+$/, "") + ".png";
      const signature: PersonSignatureRecord = {
        signatureData,
        signatureFileName,
      };
      setForm12Fields((current) => {
        const next = {
          ...current,
          signatureData,
          signatureFileName,
        };
        void saveForm12Document(next);
        return next;
      });
      setUbdRestoreFields((current) =>
        current.signatureData
          ? current
          : { ...current, signatureData, signatureFileName },
      );
      await savePersonSignatureForCurrent(signature);
      setDocumentMessage(
        "Підпис службовця збережено — доступний у всіх його документах.",
      );
    } catch {
      setDocumentMessage("Не вдалося обробити зображення підпису.");
    }
  };

  const clearForm12Signature = () => {
    setForm12Fields((current) => {
      const next = { ...current, signatureData: "", signatureFileName: "" };
      void saveForm12Document(next);
      return next;
    });
    setUbdRestoreFields((current) => ({
      ...current,
      signatureData: "",
      signatureFileName: "",
    }));
    void savePersonSignatureForCurrent(null);
  };

  const copySignatureImage = async (dataUrl: string) => {
    if (!dataUrl) return;
    try {
      await copyPngDataUrlToClipboard(dataUrl);
      setDocumentMessage("Підпис скопійовано як зображення.");
    } catch {
      setDocumentMessage("Не вдалося скопіювати підпис у буфер.");
    }
  };

  const updateUbdRestoreField = (
    key: Exclude<
      keyof UbdRestoreReportFields,
      "signatories" | "signatureData" | "signatureFileName"
    >,
    value: string,
  ) => {
    setUbdRestoreFields((current) => {
      const next = { ...current, [key]: value };
      scheduleDocumentFieldSave(() => {
        void saveUbdRestoreDocument(next);
      });
      return next;
    });
  };

  const uploadUbdRestoreSignature = async (file: File | null) => {
    if (!file) return;
    try {
      const dataUrl = await readFileAsDataUrl(file);
      const signatureData = await processSignatureTransparentBackground(dataUrl);
      const signatureFileName = file.name.replace(/\.[^.]+$/, "") + ".png";
      const signature: PersonSignatureRecord = {
        signatureData,
        signatureFileName,
      };
      setUbdRestoreFields((current) => {
        const next = {
          ...current,
          signatureData,
          signatureFileName,
        };
        void saveUbdRestoreDocument(next);
        return next;
      });
      setForm12Fields((current) =>
        current.signatureData
          ? current
          : { ...current, signatureData, signatureFileName },
      );
      await savePersonSignatureForCurrent(signature);
      setDocumentMessage(
        "Підпис службовця збережено — доступний у всіх його документах.",
      );
    } catch {
      setDocumentMessage("Не вдалося обробити зображення підпису.");
    }
  };

  const clearUbdRestoreSignature = () => {
    setUbdRestoreFields((current) => {
      const next = { ...current, signatureData: "", signatureFileName: "" };
      void saveUbdRestoreDocument(next);
      return next;
    });
    setForm12Fields((current) => ({
      ...current,
      signatureData: "",
      signatureFileName: "",
    }));
    void savePersonSignatureForCurrent(null);
  };

  const updateTicketField = (
    key: keyof TemporaryMilitaryIdFields,
    value: string,
  ) => {
    setTicketFields((current) => {
      const next = { ...current, [key]: value };
      scheduleDocumentFieldSave(() => {
        void saveTicketDocument(next);
      });
      return next;
    });
  };

  const addTicketPhoto = async (files: FileList | null) => {
    const file = files?.[0];
    if (!file || !selectedDocument) return;

    setIsSavingDocument(true);
    try {
      const dataUrl = await readFileAsDataUrl(file);
      const ticketPhoto: UbdScanFile = {
        id: `photo-${Date.now()}`,
        name: file.name,
        type: file.type || "image/jpeg",
        dataUrl,
        uploadedAt: new Date().toISOString(),
      };
      const nextFields = { ...ticketFields, photoData: dataUrl };
      const nextFiles = { ...documentFiles, ticketPhoto };
      setTicketFields(nextFields);
      setDocumentFiles(nextFiles);
      if (personExternalId) {
        await api
          .upsertPersonPhoto(personExternalId, {
            photoData: dataUrl,
            fileName: file.name,
            mimeType: file.type || "image/jpeg",
          })
          .catch(() => null);
      }
      await saveTicketDocument(nextFields, nextFiles);
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося додати фото: ${error.message}`
          : "Не вдалося додати фото.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const copyTicketDispatchLine = async () => {
    const text = formatTemporaryIdDispatchText(ticketFields);
    try {
      await navigator.clipboard.writeText(text);
      setDocumentMessage("Текст для сповіщення скопійовано.");
    } catch {
      setDocumentMessage("Не вдалося скопіювати. Виділіть текст вручну.");
    }
  };

  const questionnairePersonName =
    (summary.name && summary.name !== "Особа не вибрана"
      ? summary.name
      : "") ||
    (selectedDocument ? getDocumentPersonName(selectedDocument) : "") ||
    "Особа не вибрана";
  const questionnaireFileName = buildQuestionnaireExportFileName(
    questionnairePersonName,
    ticketFields.callSign || summary.callSign,
  );

  const closeQuestionnairePreview = () => {
    setIsQuestionnairePreviewOpen(false);
  };

  const openPersonQuestionnaire = () => {
    const targetId = String(
      selectedDocument?.personExternalId ||
        personQuestionnaire?.personExternalId ||
        personExternalId ||
        "",
    ).trim();
    if (!targetId) {
      setDocumentMessage("Немає ID службовця, щоб відкрити анкету.");
      return;
    }
    if (!personQuestionnaire) {
      setDocumentMessage("Анкета ще не додана в картці особи.");
      return;
    }

    setIsQuestionnairePreviewOpen(true);

    const cacheMatches =
      Boolean(questionnairePreviewUrl) &&
      questionnairePreviewPersonId === targetId &&
      Boolean(questionnaireFileData);
    if (cacheMatches) return;
    if (
      isLoadingQuestionnairePreview &&
      questionnairePreviewPersonId === targetId
    ) {
      return;
    }

    const requestSeq = questionnaireLoadSeqRef.current + 1;
    questionnaireLoadSeqRef.current = requestSeq;
    setQuestionnairePreviewPersonId(targetId);
    setIsLoadingQuestionnairePreview(true);
    setQuestionnaireFileData("");
    setQuestionnairePreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      return "";
    });

    void api
      .getPersonQuestionnaire(targetId)
      .then((full) => {
        if (requestSeq !== questionnaireLoadSeqRef.current) return;
        if (!full?.fileData) {
          setDocumentMessage("Не вдалося завантажити PDF анкети.");
          return;
        }
        const nextUrl = dataUrlToObjectUrl(full.fileData);
        setQuestionnaireFileData(full.fileData);
        setQuestionnairePreviewPersonId(full.personExternalId || targetId);
        setQuestionnairePreviewUrl((current) => {
          if (current) revokeQuestionnairePreviewUrl(current);
          return nextUrl;
        });
      })
      .catch((error) => {
        if (requestSeq !== questionnaireLoadSeqRef.current) return;
        setDocumentMessage(
          error instanceof Error
            ? `Не вдалося відкрити анкету: ${error.message}`
            : "Не вдалося відкрити анкету.",
        );
      })
      .finally(() => {
        if (requestSeq !== questionnaireLoadSeqRef.current) return;
        setIsLoadingQuestionnairePreview(false);
      });
  };

  const openPersonQuestionnaireInNewTab = async () => {
    const targetId = String(
      selectedDocument?.personExternalId ||
        personQuestionnaire?.personExternalId ||
        personExternalId ||
        "",
    ).trim();
    if (
      questionnairePreviewUrl &&
      questionnairePreviewPersonId === targetId
    ) {
      window.open(questionnairePreviewUrl, "_blank", "noopener,noreferrer");
      return;
    }
    if (!targetId || !personQuestionnaire) return;
    try {
      const url = await api.getPersonQuestionnaireObjectUrl(
        targetId,
        personQuestionnaire.fileName || questionnaireFileName,
      );
      window.open(url, "_blank", "noopener,noreferrer");
      window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося відкрити анкету: ${error.message}`
          : "Не вдалося відкрити анкету.",
      );
    }
  };

  const updateSalaryField = (
    key: keyof SalaryDocumentFields,
    value: string,
  ) => {
    setSalaryFields((current) => {
      if (key === "iban") {
        const iban = normalizeIban(value);
        const mfo = getBankMfoFromIban(iban);
        const bank = findBankByMfo(mfo);
        const next = {
          ...current,
          iban,
          bankMfo: mfo || current.bankMfo,
          bankName: bank?.name || current.bankName,
        };
        scheduleDocumentFieldSave(() => {
          void saveSalaryDocument(next, workflow);
        });
        return next;
      }

      if (key === "bankMfo") {
        const bank = findBankByMfo(value);
        const next = {
          ...current,
          bankMfo: value,
          bankName: value === "custom" ? current.bankName : bank?.name || "",
        };
        scheduleDocumentFieldSave(() => {
          void saveSalaryDocument(next, workflow);
        });
        return next;
      }

      const next = { ...current, [key]: value };
      scheduleDocumentFieldSave(() => {
        void saveSalaryDocument(next, workflow);
      });
      return next;
    });
  };

  const saveActiveDocumentWorkflow = (
    nextFields:
      | SalaryDocumentFields
      | UbdReportFields
      | Form6ReportFields
      | Form12ReportFields
      | ServiceCharacteristicFields
      | ZhbdCertificateFields
      | UbdRestoreReportFields
      | TemporaryMilitaryIdFields,
    nextWorkflow: SalaryWorkflowState,
    nextFiles = documentFiles,
  ) => {
    if (!canMutateSelected) {
      setDocumentMessage(
        selectedDocument
          ? `Лише перегляд · автор: ${documentCreatorLabel(selectedDocument)}`
          : "Можна змінювати лише свої документи.",
      );
      return;
    }
    if (selectedDocument?.type === "ubdReport" || mode === "ubdReport") {
      void saveUbdDocument(
        nextFields as UbdReportFields,
        nextFiles,
        nextWorkflow,
      );
      return;
    }

    if (selectedDocument?.type === "form6Report" || mode === "form6Report") {
      void saveForm6Document(
        nextFields as Form6ReportFields,
        nextFiles,
        nextWorkflow,
      );
      return;
    }

    if (selectedDocument?.type === "form12Report" || mode === "form12Report") {
      void saveForm12Document(
        nextFields as Form12ReportFields,
        nextFiles,
        nextWorkflow,
      );
      return;
    }

    if (
      selectedDocument?.type === "serviceCharacteristic" ||
      mode === "serviceCharacteristic"
    ) {
      void saveServiceCharacteristicDocument(
        nextFields as ServiceCharacteristicFields,
        nextFiles,
        nextWorkflow,
      );
      return;
    }

    if (
      selectedDocument?.type === "zhbdCertificate" ||
      mode === "zhbdCertificate"
    ) {
      void saveZhbdCertificateDocument(
        nextFields as ZhbdCertificateFields,
        nextFiles,
        nextWorkflow,
      );
      return;
    }

    if (
      selectedDocument?.type === "ubdRestoreReport" ||
      mode === "ubdRestoreReport"
    ) {
      void saveUbdRestoreDocument(
        nextFields as UbdRestoreReportFields,
        nextFiles,
        nextWorkflow,
      );
      return;
    }

    if (
      selectedDocument?.type === "temporaryMilitaryId" ||
      mode === "temporaryMilitaryId"
    ) {
      void saveTicketDocument(ticketFields, nextFiles, nextWorkflow);
      return;
    }

    void saveSalaryDocument(nextFields as SalaryDocumentFields, nextWorkflow);
  };

  const updateWorkflow = (next: SalaryWorkflowState) => {
    if (!canMutateSelected) {
      setDocumentMessage(
        selectedDocument
          ? `Лише перегляд · автор: ${documentCreatorLabel(selectedDocument)}`
          : "Можна змінювати лише свої документи.",
      );
      return;
    }
    setWorkflow(next);
    patchSelectedDocumentWorkflow(next);
    saveWorkflowForPerson(workflowKey, next);
    saveActiveDocumentWorkflow(
      mode === "ubdReport"
        ? ubdFields
        : mode === "form6Report"
          ? form6Fields
          : mode === "form12Report"
            ? form12Fields
          : mode === "serviceCharacteristic"
            ? serviceCharacteristicFields
          : mode === "zhbdCertificate"
            ? zhbdCertificateFields
          : mode === "ubdRestoreReport"
            ? ubdRestoreFields
            : mode === "temporaryMilitaryId"
            ? ticketFields
            : salaryFields,
      next,
    );
  };

  const setWorkflowStep = (stepKey: string) => {
    const selectedIndex = activeWorkflowSteps.findIndex(
      (step) => step.key === stepKey,
    );
    if (selectedIndex < 0) return;
    const completed = Object.fromEntries(
      activeWorkflowSteps.map((step, index) => [
        step.key,
        index <= selectedIndex,
      ]),
    ) as Record<string, boolean>;

    updateWorkflow({
      ...workflow,
      completed,
      currentStatus: stepKey,
    });
    if (stepKey === "fighterSign") {
      window.requestAnimationFrame(() => {
        document.getElementById("fighter-signature-panel")?.scrollIntoView({
          block: "center",
          behavior: "smooth",
        });
      });
    }
  };

  const generateSalaryDoc = () => {
    const content = salaryDocumentHtml(salaryFields);
    downloadTextFile(
      `${safeFilePart(salaryFields.folderName)}.doc`,
      content,
      "application/msword;charset=utf-8",
    );
    setWorkflowStep("document");
  };

  const printSalaryDocument = () => {
    const printWindow = window.open("", "_blank", "noopener,noreferrer");
    if (!printWindow) return;
    printWindow.document.write(salaryDocumentHtml(salaryFields));
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
    setWorkflowStep("print");
  };

  const saveUbdAsWord = async () => {
    setDocumentMessage("Формую Word-документ...");
    const blob = await createUbdWordBlob(ubdWordFields);
    const fileName = `${safeFilePart(ubdFields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(
      `Word-файл «${fileName}» збережено.`,
    );
    setWorkflowStep("document");
    void saveUbdDocument(ubdFields);
  };

  const saveForm6AsWord = async () => {
    setDocumentMessage("Формую Word-документ Форми 6...");
    const blob = await createForm6WordBlob(form6Fields);
    const fileName = `${safeFilePart(form6Fields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(`Word-файл «${fileName}» збережено.`);
    setWorkflowStep("document");
    void saveForm6Document(form6Fields);
  };


  const saveServiceCharacteristicAsWord = async () => {
    setDocumentMessage("Формую Word-документ службової характеристики...");
    const blob = await createServiceCharacteristicWordBlob(
      serviceCharacteristicFields,
    );
    const fileName = `${safeFilePart(serviceCharacteristicFields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(`Word-файл «${fileName}» збережено.`);
    setWorkflowStep("document");
    void saveServiceCharacteristicDocument(serviceCharacteristicFields);
  };

  const saveZhbdCertificateAsWord = async () => {
    setDocumentMessage("Формую Word-документ довідки ЖБД...");
    const blob = await createZhbdCertificateWordBlob(zhbdCertificateFields);
    const fileName = `${safeFilePart(zhbdCertificateFields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(`Word-файл «${fileName}» збережено.`);
    setWorkflowStep("document");
    void saveZhbdCertificateDocument(zhbdCertificateFields);
  };

  const saveForm12AsWord = async () => {
    setDocumentMessage("Формую Word-документ Форми 12...");
    const blob = await createForm12WordBlob(form12Fields);
    const fileName = `${safeFilePart(form12Fields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(`Word-файл «${fileName}» збережено.`);
    setWorkflowStep("document");
    void saveForm12Document(form12Fields);
  };

  const saveUbdRestoreAsWord = async () => {
    setDocumentMessage("Формую Word-документ рапорта на відновлення УБД...");
    const blob = await createUbdRestoreWordBlob(ubdRestoreFields);
    const fileName = `${safeFilePart(ubdRestoreFields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(`Word-файл «${fileName}» збережено.`);
    setWorkflowStep("document");
    void saveUbdRestoreDocument(ubdRestoreFields);
  };

  const printUbdDocument = () => {
    printRenderedWordPreview();
    void saveUbdDocument(ubdFields);
  };

  const documentStepProgress = selectedDocument ? (
    <div className="salary-step-progress-header">
      <div className="salary-progress">
        <span>{progressPercent}%</span>
        <div>
          <i style={{ width: `${progressPercent}%` }} />
        </div>
      </div>
      <div className="salary-step-bar">
        {activeWorkflowSteps.map((step, index) => {
          const isDone = index <= currentStepIndex;
          const isCurrent = index === currentStepIndex;
          return (
            <button
              className={[
                "salary-step-node",
                isDone ? "done" : "",
                isCurrent ? "current" : "",
                canMutateSelected ? "" : "is-readonly",
              ]
                .filter(Boolean)
                .join(" ")}
              key={step.key}
              type="button"
              disabled={!canMutateSelected}
              onClick={() => {
                if (!canMutateSelected) return;
                setWorkflowStep(step.key);
              }}
              title={
                canMutateSelected
                  ? step.title
                  : `${step.title} · лише перегляд`
              }
            >
              <span>{index + 1}</span>
              <i />
              <strong>{step.title}</strong>
            </button>
          );
        })}
      </div>
    </div>
  ) : null;

  const deleteDocument = async (document: BackendPersonDocument) => {
    if (!canMutateDocument(document)) {
      setDocumentMessage("Можна видаляти лише документи, які ви створили.");
      return;
    }
    const confirmed = window.confirm(
      `Видалити документ "${document.title}" для ID ${document.personExternalId}?`,
    );
    if (!confirmed) return;

    setIsSavingDocument(true);
    setDocumentMessage("Видаляю документ...");
    try {
      await api.deletePersonDocument(document.personExternalId, document.id);
      setAllPersonDocuments((current) =>
        current.filter((item) => item.id !== document.id),
      );
      setPersonDocuments((current) => {
        const next = current.filter((item) => item.id !== document.id);
        if (document.id === selectedDocumentId) {
          const replacement = next[0] ?? null;
          setSelectedDocumentId(replacement?.id ?? "");
          if (replacement) {
            openPersonDocument(replacement);
          } else {
            setSalaryFields(createSalaryFields(selectedPerson, summary));
            setWorkflow(createEmptyWorkflow());
          }
        }
        return next;
      });
      setDocumentMessage(`Документ видалено: ${document.title}.`);
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося видалити документ: ${error.message}`
          : "Не вдалося видалити документ.",
      );
    } finally {
      setIsSavingDocument(false);
    }
  };

  const questionnaireHeaderButton = (
    <Button
      variant="outlined"
      startIcon={<PictureAsPdfOutlinedIcon />}
      onClick={openPersonQuestionnaire}
    >
      Анкета
    </Button>
  );

  const createdDocumentTypes = new Set(
    [...personDocuments, ...allPersonDocuments]
      .filter(
        (document) =>
          document.personExternalId === personExternalId &&
          document.type !== PERSON_PHONES_DOCUMENT_TYPE &&
          document.type !== PERSON_SIGNATURE_DOCUMENT_TYPE,
      )
      .map((document) => document.type),
  );

  // Internal storage docs (phones / signature) stay in state for sync, but not in UI lists.
  const visiblePersonDocuments = personDocuments.filter(
    (document) =>
      document.type !== PERSON_PHONES_DOCUMENT_TYPE &&
      document.type !== PERSON_SIGNATURE_DOCUMENT_TYPE,
  );

  const createPersonDocumentByType = (type: string) => {
    if (!canEdit) return;
    if (type === "salaryPowerAttorney") return createSalaryPowerAttorneyDocument();
    if (type === "ubdReport") return createUbdReportDocument();
    if (type === "ubdRestoreReport") return createUbdRestoreReportDocument();
    if (type === "form6Report") return createForm6ReportDocument();
    if (type === "form12Report") return createForm12ReportDocument();
    if (type === "serviceCharacteristic")
      return createServiceCharacteristicDocument();
    if (type === "zhbdCertificate") return createZhbdCertificateDocument();
    if (type === "temporaryMilitaryId")
      return createTemporaryMilitaryIdDocument();
  };

  const personDocumentCreateSelect = (
    <TextField
      select
      size="small"
      label="Створити документ"
      value="__none"
      disabled={isSavingDocument || !personExternalId || !canEdit}
      onChange={(event) => {
        const nextType = event.target.value;
        if (!nextType || nextType === "__none") return;
        void createPersonDocumentByType(nextType);
      }}
    >
      <MenuItem value="__none">Оберіть тип документа</MenuItem>
      {CREATE_PERSON_DOCUMENT_TYPES.map((item) => {
        const exists = createdDocumentTypes.has(item.value);
        return (
          <MenuItem
            key={item.value}
            value={item.value}
            className={exists ? "is-ready" : undefined}
          >
            {exists ? `${item.label} · є` : item.label}
          </MenuItem>
        );
      })}
    </TextField>
  );

  const personQuestionnaireRow = (
    <article className="salary-person-document-shell ticket-questionnaire-shell">
      <button
        className={
          personQuestionnaire
            ? "salary-person-document"
            : "salary-person-document is-empty"
        }
        type="button"
        onClick={openPersonQuestionnaire}
      >
        <strong>Анкета службовця (PDF)</strong>
        <span>
          {isLoadingQuestionnairePreview
            ? "Завантажую PDF..."
            : personQuestionnaire
              ? "Переглянути PDF"
              : "Анкета ще не додана в картці особи"}
        </span>
        <small>{personQuestionnaire?.fileName || questionnaireFileName}</small>
      </button>
    </article>
  );

  const documentStatusNotePanel = (
    value: string,
    onChange: (value: string) => void,
  ) => (
    <div className="document-status-note-panel">
      <div className="panel-heading">Коментар до статусу</div>
      <TextField
        size="small"
        fullWidth
        multiline
        rows={3}
        value={value}
        disabled={!canMutateSelected}
        InputProps={{ readOnly: !canMutateSelected }}
        onChange={(event) => {
          if (!canMutateSelected) return;
          onChange(event.target.value);
        }}
        placeholder={
          canMutateSelected
            ? "Проблема з документами, уточнення по статусу або що треба доробити"
            : "Лише перегляд"
        }
      />
    </div>
  );

  const liveDocumentStatusNote = (document: BackendPersonDocument) => {
    if (document.id !== selectedDocumentId) {
      return readDocumentStatusNote(document);
    }
    if (mode === "ubdReport") return ubdFields.statusNote.trim();
    if (mode === "form6Report") return form6Fields.statusNote.trim();
    if (mode === "form12Report") return form12Fields.statusNote.trim();
    if (mode === "serviceCharacteristic")
      return serviceCharacteristicFields.statusNote.trim();
    if (mode === "zhbdCertificate")
      return zhbdCertificateFields.statusNote.trim();
    if (mode === "ubdRestoreReport") return ubdRestoreFields.statusNote.trim();
    if (mode === "temporaryMilitaryId") return ticketFields.statusNote.trim();
    if (mode === "salaryPowerAttorney") return salaryFields.statusNote.trim();
    return readDocumentStatusNote(document);
  };

  const documentListStatusNote = (document: BackendPersonDocument) => {
    const note = liveDocumentStatusNote(document);
    if (!note) return null;
    return (
      <small className="document-status-note-preview" title={note}>
        {note}
      </small>
    );
  };

  const questionnairePreviewDialog = (
    <FloatingQuestionnairePreview
      open={isQuestionnairePreviewOpen}
      title={`Анкета · ${questionnairePersonName}${
        personQuestionnaire?.fileName || questionnaireFileName
          ? ` · ${personQuestionnaire?.fileName || questionnaireFileName}`
          : ""
      }`}
      previewUrl={questionnairePreviewUrl}
      pendingFile={false}
      isUploading={isLoadingQuestionnairePreview}
      placement="left"
      shareFileName={questionnaireFileName}
      sharePersonName={questionnairePersonName}
      shareSource={
        questionnaireFileData ? { fileData: questionnaireFileData } : null
      }
      onShareNotify={setDocumentMessage}
      onClose={closeQuestionnairePreview}
      onOpenTab={openPersonQuestionnaireInNewTab}
      childrenHint={
        isLoadingQuestionnairePreview ? "Завантажую PDF анкети..." : undefined
      }
      onDownload={
        questionnaireFileData
          ? () =>
              downloadQuestionnairePdf(questionnaireFileName, {
                fileData: questionnaireFileData,
              })
          : undefined
      }
    />
  );

  const wordPreviewPanel = selectedDocument ? (
    <section className="analytics-panel document-preview">
      <div className="panel-heading">Попередній перегляд Word</div>
      <WordDocumentPreview
        blob={wordPreview.blob}
        error={wordPreview.error}
        isLoading={wordPreview.isLoading}
      />
    </section>
  ) : null;

  const ubdWorkspace = (
    <>
          {selectedDocument ? (
            <section className="analytics-panel document-fields">
              <div className="panel-heading">Дані для УБД рапорта</div>
              <div className="document-placeholder-map salary-document-map ubd-document-map">
                {[
                  ["fullName", "ПІБ"],
                  ["rank", "Звання"],
                  ["staffPosition", "Посада"],
                  ["birthDate", "Дата народження"],
                  ["rnokpp", "РНОКПП"],
                  ["taskPeriod", "Період завдань"],
                  ["taskPlace", "Місце завдань"],
                  ["date", "Дата"],
                  ["folderName", "Папка"],
                ].map(([key, label]) => (
                  <label key={key}>
                    <code>{label}</code>
                    <TextField
                      size="small"
                      fullWidth
                      value={String(
                        ubdFields[key as keyof UbdReportFields] ?? "",
                      )}
                      onChange={(event) =>
                        updateUbdField(
                          key as keyof UbdReportFields,
                          event.target.value,
                        )
                      }
                    />
                  </label>
                ))}
                <label>
                  <code>№ розпорядження</code>
                  <TextField
                    size="small"
                    fullWidth
                    value={ubdFields.basisNumber}
                    onChange={(event) =>
                      updateUbdBasisPart("basisNumber", event.target.value)
                    }
                  />
                </label>
                <label>
                  <code>Дата розпорядження</code>
                  <TextField
                    size="small"
                    fullWidth
                    value={ubdFields.basisDate}
                    onChange={(event) =>
                      updateUbdBasisPart("basisDate", event.target.value)
                    }
                  />
                </label>
                <label className="wide">
                  <code>Кому</code>
                  <TextField
                    size="small"
                    fullWidth
                    value={ubdFields.commander}
                    onChange={(event) =>
                      updateUbdField("commander", event.target.value)
                    }
                  />
                </label>
                <div className="wide document-default-signatories">
                  <code>Дефолтні записи (редагуються у налаштуваннях)</code>
                  {ubdFields.signatories.map((signatory, index) => (
                    <article key={`${signatory.sourceId ?? signatory.label}-${index}`}>
                      <strong>
                        {signatory.blockType === "APPROVAL"
                          ? "ЗАТВЕРДЖУЮ · "
                          : ""}
                        {signatory.label}
                      </strong>
                      <span>{signatory.title}</span>
                      <small>
                        {signatory.rank} · {signatory.fullName}
                      </small>
                    </article>
                  ))}
                </div>
              </div>

              {documentStatusNotePanel(ubdFields.statusNote, (value) =>
                updateUbdField("statusNote", value),
              )}

              {personQuestionnaireRow}

              <div className="ubd-scans-panel" id="ubd-scans-panel">
                <div className="panel-heading">Скани для УБД</div>
                <Button component="label" variant="outlined">
                  Додати скани
                  <input
                    hidden
                    multiple
                    type="file"
                    accept="image/*,application/pdf,.pdf"
                    onChange={(event) => {
                      void addUbdScanFiles(event.target.files);
                      event.target.value = "";
                    }}
                  />
                </Button>
                <div className="ubd-scan-list">
                  {(documentFiles.ubdScans ?? []).length ? (
                    (documentFiles.ubdScans ?? []).map((file) => (
                      <article className="ubd-scan-item" key={file.id}>
                        <a
                          href={file.dataUrl}
                          target="_blank"
                          rel="noreferrer"
                          onClick={(event) => {
                            event.preventDefault();
                            openDataUrlInNewTab(file.dataUrl);
                          }}
                        >
                          {file.type.startsWith("image/") ? (
                            <img alt={file.name} src={file.dataUrl} />
                          ) : (
                            <PictureAsPdfOutlinedIcon />
                          )}
                          <span>{file.name}</span>
                        </a>
                        <button
                          aria-label="Видалити скан"
                          className="documents-journal-delete"
                          onClick={() => deleteUbdScanFile(file.id)}
                          type="button"
                        >
                          <DeleteOutlineOutlinedIcon />
                        </button>
                      </article>
                    ))
                  ) : (
                    <p className="salary-empty-documents">
                      Додайте фото, ІНН, копію паспорта або затверджений рапорт.
                    </p>
                  )}
                </div>
              </div>
            </section>
          ) : null}

          {wordPreviewPanel}
    </>
  );

  const form6Workspace = (
    <>
      {selectedDocument ? (
        <section className="analytics-panel document-fields">
          <div className="panel-heading">Дані для Форми 6</div>
          <div className="document-placeholder-map salary-document-map ubd-document-map">
            {(
              [
                ["commander", "Кому"],
                ["rank", "Звання"],
                ["fullName", "ПІБ"],
                ["staffPosition", "Посада згідно штату"],
                ["birthDate", "Дата народження"],
                ["idDocument", "Документ, що посвідчує особу"],
                ["rnokpp", "РНОКПП"],
                ["address", "Адреса проживання"],
                ["phone", "Телефон"],
                ["taskPeriod", "Період завдань"],
                ["taskPlace", "Місце завдань"],
                ["basis", "Підстава"],
                ["date", "Дата"],
                ["folderName", "Папка"],
              ] as Array<
                [Exclude<keyof Form6ReportFields, "signatories">, string]
              >
            ).map(([key, label]) => (
              <label key={key} className={key === "staffPosition" || key === "address" || key === "basis" || key === "commander" ? "wide" : undefined}>
                <code>{label}</code>
                <TextField
                  size="small"
                  fullWidth
                  multiline={
                    key === "staffPosition" ||
                    key === "address" ||
                    key === "basis" ||
                    key === "commander"
                  }
                  rows={
                    key === "staffPosition" ||
                    key === "address" ||
                    key === "basis" ||
                    key === "commander"
                      ? 2
                      : undefined
                  }
                  value={form6Fields[key]}
                  onChange={(event) =>
                    updateForm6Field(key, event.target.value)
                  }
                />
              </label>
            ))}
          </div>
          {documentStatusNotePanel(form6Fields.statusNote, (value) =>
            updateForm6Field("statusNote", value),
          )}
          {personQuestionnaireRow}
          <div className="ubd-scans-panel" id="ubd-scans-panel">
            <div className="panel-heading">Скани для Форми 6</div>
            <Button component="label" variant="outlined">
              Додати скани
              <input
                hidden
                multiple
                type="file"
                accept="image/*,application/pdf,.pdf"
                onChange={(event) => {
                  void addUbdScanFiles(event.target.files);
                  event.target.value = "";
                }}
              />
            </Button>
            <div className="ubd-scan-list">
              {(documentFiles.ubdScans ?? []).length ? (
                (documentFiles.ubdScans ?? []).map((file) => (
                  <article className="ubd-scan-item" key={file.id}>
                    <a
                      href={file.dataUrl}
                      target="_blank"
                      rel="noreferrer"
                      onClick={(event) => {
                        event.preventDefault();
                        openDataUrlInNewTab(file.dataUrl);
                      }}
                    >
                      {file.type.startsWith("image/") ? (
                        <img alt={file.name} src={file.dataUrl} />
                      ) : (
                        <PictureAsPdfOutlinedIcon />
                      )}
                      <span>{file.name}</span>
                    </a>
                    <button
                      aria-label="Видалити скан"
                      className="documents-journal-delete"
                      onClick={() => deleteUbdScanFile(file.id)}
                      type="button"
                    >
                      <DeleteOutlineOutlinedIcon />
                    </button>
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Додайте паспорт, РНОКПП, витяг про проживання або фото 3х4.
                </p>
              )}
            </div>
          </div>
        </section>
      ) : null}

      {wordPreviewPanel}
    </>
  );

  const form12Workspace = (
    <>
      {selectedDocument ? (
        <section className="analytics-panel document-fields">
          <div className="panel-heading">Дані для Форми 12</div>
          <div className="document-placeholder-map salary-document-map ubd-document-map">
            {(
              [
                ["commander", "Кому"],
                ["rank", "Звання"],
                ["fullName", "ПІБ"],
                ["staffPosition", "Посада згідно штату"],
                ["date", "Дата"],
                ["folderName", "Папка"],
              ] as Array<
                [Exclude<keyof Form12ReportFields, "signatories" | "signatureData" | "signatureFileName" | "statusNote">, string]
              >
            ).map(([key, label]) => (
              <label
                key={key}
                className={
                  key === "staffPosition" || key === "commander" ? "wide" : undefined
                }
              >
                <code>{label}</code>
                <TextField
                  size="small"
                  fullWidth
                  multiline={key === "staffPosition" || key === "commander"}
                  rows={key === "staffPosition" || key === "commander" ? 2 : undefined}
                  value={form12Fields[key]}
                  onChange={(event) =>
                    updateForm12Field(key, event.target.value)
                  }
                />
              </label>
            ))}
          </div>
          <div className="form12-signature-upload" id="fighter-signature-panel">
            <div className="panel-heading">Підпис бійця (PNG)</div>
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              Зберігається для службовця і підставляється в інші його документи.
            </Typography>
            <Stack direction="row" spacing={1}>
              <Button component="label" variant="outlined">
                Завантажити PNG
                <input
                  hidden
                  type="file"
                  accept="image/png,image/jpeg,image/webp,.png,.jpg,.jpeg"
                  onChange={(event) => {
                    void uploadForm12Signature(event.target.files?.[0] ?? null);
                    event.target.value = "";
                  }}
                />
              </Button>
              {form12Fields.signatureData ? (
                <>
                  <Button
                    variant="outlined"
                    startIcon={<ContentCopyOutlinedIcon />}
                    onClick={() =>
                      void copySignatureImage(form12Fields.signatureData)
                    }
                  >
                    Копіювати
                  </Button>
                  <Button variant="text" onClick={clearForm12Signature}>
                    Прибрати
                  </Button>
                </>
              ) : null}
            </Stack>
            {form12Fields.signatureData ? (
              <div className="form12-signature-preview">
                <img alt="Підпис бійця" src={form12Fields.signatureData} />
                <small>{form12Fields.signatureFileName || "підпис.png"}</small>
              </div>
            ) : (
              <p className="salary-empty-documents">
                Завантажте підпис бійця — білий фон прибереться, у Word піде PNG.
              </p>
            )}
          </div>
          {documentStatusNotePanel(form12Fields.statusNote, (value) =>
            updateForm12Field("statusNote", value),
          )}
          {personQuestionnaireRow}
          <div className="ubd-scans-panel" id="ubd-scans-panel">
            <div className="panel-heading">Скани для Форми 12</div>
            <Button component="label" variant="outlined">
              Додати скани
              <input
                hidden
                multiple
                type="file"
                accept="image/*,application/pdf,.pdf"
                onChange={(event) => {
                  void addUbdScanFiles(event.target.files);
                  event.target.value = "";
                }}
              />
            </Button>
            <div className="ubd-scan-list">
              {(documentFiles.ubdScans ?? []).length ? (
                (documentFiles.ubdScans ?? []).map((file) => (
                  <article className="ubd-scan-item" key={file.id}>
                    <a
                      href={file.dataUrl}
                      target="_blank"
                      rel="noreferrer"
                      onClick={(event) => {
                        event.preventDefault();
                        openDataUrlInNewTab(file.dataUrl);
                      }}
                    >
                      {file.type.startsWith("image/") ? (
                        <img alt={file.name} src={file.dataUrl} />
                      ) : (
                        <PictureAsPdfOutlinedIcon />
                      )}
                      <span>{file.name}</span>
                    </a>
                    <button
                      aria-label="Видалити скан"
                      className="documents-journal-delete"
                      onClick={() => deleteUbdScanFile(file.id)}
                      type="button"
                    >
                      <DeleteOutlineOutlinedIcon />
                    </button>
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  За потреби додайте скани до рапорту.
                </p>
              )}
            </div>
          </div>
        </section>
      ) : null}

      {wordPreviewPanel}
    </>
  );


  const serviceCharacteristicWorkspace = (
    <>
      {selectedDocument ? (
        <section className="analytics-panel document-fields">
          <div className="panel-heading">Дані для службової характеристики</div>
          <div className="document-placeholder-map salary-document-map ubd-document-map">
            {(
              [
                ["rank", "Звання"],
                ["lastName", "Прізвище"],
                ["firstName", "Ім’я"],
                ["patronymic", "По батькові"],
                ["staffPosition", "Займає посаду"],
                ["introParagraph", "Вступний абзац"],
                ["professionalParagraph", "Професійна підготовка"],
                ["combatParagraph", "Дії в складних умовах"],
                ["moralParagraph", "Моральні якості"],
                ["drillParagraph", "Стройова / фізична підготовка"],
                ["conclusion", "Висновок"],
                ["date", "Дата"],
                ["folderName", "Папка"],
              ] as Array<
                [
                  Exclude<keyof ServiceCharacteristicFields, "signatories">,
                  string,
                ]
              >
            ).map(([key, label]) => (
              <label
                key={key}
                className={
                  key === "staffPosition" ||
                  key.endsWith("Paragraph") ||
                  key === "conclusion"
                    ? "wide"
                    : undefined
                }
              >
                <code>{label}</code>
                <TextField
                  size="small"
                  fullWidth
                  multiline={
                    key === "staffPosition" ||
                    key.endsWith("Paragraph") ||
                    key === "conclusion"
                  }
                  rows={
                    key.endsWith("Paragraph")
                      ? 4
                      : key === "staffPosition" || key === "conclusion"
                        ? 2
                        : undefined
                  }
                  value={serviceCharacteristicFields[key]}
                  onChange={(event) =>
                    updateServiceCharacteristicField(key, event.target.value)
                  }
                />
              </label>
            ))}
          </div>
          {documentStatusNotePanel(
            serviceCharacteristicFields.statusNote,
            (value) => updateServiceCharacteristicField("statusNote", value),
          )}
          {personQuestionnaireRow}
        </section>
      ) : null}

      {wordPreviewPanel}
    </>
  );

  const zhbdCertificateWorkspace = (
    <>
      {selectedDocument ? (
        <section className="analytics-panel document-fields">
          <div className="panel-heading">Дані для довідки ЖБД</div>
          <div className="zhbd-full-position-banner">
            <code>Повна посада (з загального списку)</code>
            <strong>
              {zhbdDisplayedFullPosition || "немає в ранковому списку"}
            </strong>
          </div>
          <div className="document-placeholder-map salary-document-map ubd-document-map">
            {(
              [
                ["rank", "Звання"],
                ["fullName", "ПІБ"],
                [
                  "staffPosition",
                  "Посада на момент БЗ",
                ],
                ["periodFrom", "Період з"],
                ["periodTo", "Період по"],
                ["periodNote", "Примітка до періоду"],
                ["bodyParagraph", "Текст довідки"],
                ["basisOrders", "Підстава · розпорядження"],
                ["basisJournal", "Підстава · журнал"],
                ["headerDate", "Дата в шапці"],
                ["documentNumber", "Номер документа"],
                ["date", "Дата"],
                ["folderName", "Папка"],
              ] as Array<
                [
                  Exclude<
                    keyof ZhbdCertificateFields,
                    "signatories" | "actualFullPosition"
                  >,
                  string,
                ]
              >
            ).map(([key, label]) => (
              <label
                key={key}
                className={
                  key === "staffPosition" ||
                  key === "bodyParagraph" ||
                  key === "basisOrders" ||
                  key === "basisJournal" ||
                  key === "periodNote"
                    ? "wide"
                    : undefined
                }
              >
                <code>{label}</code>
                <div className="document-field-with-hint">
                  <TextField
                    size="small"
                    fullWidth
                    multiline={
                      key === "staffPosition" ||
                      key === "bodyParagraph" ||
                      key === "basisOrders" ||
                      key === "basisJournal" ||
                      key === "periodNote"
                    }
                    rows={
                      key === "bodyParagraph"
                        ? 4
                        : key === "basisOrders" ||
                            key === "basisJournal" ||
                            key === "staffPosition"
                          ? 2
                          : undefined
                    }
                    value={zhbdCertificateFields[key]}
                    onChange={(event) =>
                      updateZhbdCertificateField(key, event.target.value)
                    }
                  />
                  {key === "staffPosition" ? (
                    <span className="document-field-hint">
                      Повна посада:{" "}
                      <strong>
                        {zhbdDisplayedFullPosition || "немає в ранковому списку"}
                      </strong>
                    </span>
                  ) : null}
                </div>
              </label>
            ))}
            <div className="wide document-default-signatories">
              <code>Підписант (з налаштувань)</code>
              {zhbdCertificateFields.signatories.length ? (
                zhbdCertificateFields.signatories.map((signatory, index) => (
                  <article
                    key={`${signatory.blockType}-${signatory.fullName}-${index}`}
                  >
                    <strong>
                      {signatory.blockType === "APPROVAL"
                        ? "ЗАТВЕРДЖУЮ · "
                        : "Підписант · "}
                      {signatory.fullName || "—"}
                    </strong>
                    <span>{signatory.title}</span>
                    <small>
                      {signatory.rank}
                      {signatory.signatureData
                        ? " · є зображення підпису"
                        : " · немає зображення підпису"}
                    </small>
                  </article>
                ))
              ) : (
                <p>
                  Немає записів. Увімкни «Довідка ЖБД» у налаштуваннях підписів.
                </p>
              )}
            </div>
          </div>
          {documentStatusNotePanel(zhbdCertificateFields.statusNote, (value) =>
            updateZhbdCertificateField("statusNote", value),
          )}
          {personQuestionnaireRow}
        </section>
      ) : null}

      {wordPreviewPanel}
    </>
  );

  const ubdRestoreWorkspace = (
    <>
      {selectedDocument ? (
        <section className="analytics-panel document-fields">
          <div className="panel-heading">Дані для рапорта на відновлення УБД</div>
          <div className="document-placeholder-map salary-document-map ubd-document-map">
            {(
              [
                ["commander", "Кому (1 сторінка)"],
                ["rank", "Звання"],
                ["fullName", "ПІБ"],
                ["staffPosition", "Посада"],
                ["signerTitle", "Підпис, посада"],
                ["certificateSeries", "Серія посвідчення УБД"],
                ["circumstances", "Обставини пошкодження"],
                ["requestText", "Прошу"],
                ["coveringCommander", "Кому (клопотання)"],
                ["date", "Дата"],
                ["folderName", "Папка"],
              ] as Array<
                [Exclude<keyof UbdRestoreReportFields, "signatories" | "signatureData" | "signatureFileName" | "statusNote">, string]
              >
            ).map(([key, label]) => (
              <label
                key={key}
                className={
                  key === "staffPosition" ||
                  key === "signerTitle" ||
                  key === "circumstances" ||
                  key === "requestText" ||
                  key === "commander" ||
                  key === "coveringCommander"
                    ? "wide"
                    : undefined
                }
              >
                <code>{label}</code>
                <TextField
                  size="small"
                  fullWidth
                  multiline={
                    key === "staffPosition" ||
                    key === "signerTitle" ||
                    key === "circumstances" ||
                    key === "requestText" ||
                    key === "commander" ||
                    key === "coveringCommander"
                  }
                  rows={
                    key === "circumstances" || key === "requestText"
                      ? 3
                      : key === "staffPosition" ||
                          key === "signerTitle" ||
                          key === "commander" ||
                          key === "coveringCommander"
                        ? 2
                        : undefined
                  }
                  value={String(ubdRestoreFields[key] ?? "")}
                  onChange={(event) =>
                    updateUbdRestoreField(key, event.target.value)
                  }
                />
              </label>
            ))}
          </div>
          <div className="form12-signature-upload" id="fighter-signature-panel">
            <div className="panel-heading">Підпис службовця (PNG)</div>
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              Зберігається для службовця і підставляється в інші його документи.
            </Typography>
            <Stack direction="row" spacing={1}>
              <Button component="label" variant="outlined">
                Завантажити PNG
                <input
                  hidden
                  type="file"
                  accept="image/png,image/jpeg,image/webp,.png,.jpg,.jpeg"
                  onChange={(event) => {
                    void uploadUbdRestoreSignature(event.target.files?.[0] ?? null);
                    event.target.value = "";
                  }}
                />
              </Button>
              {ubdRestoreFields.signatureData ? (
                <>
                  <Button
                    variant="outlined"
                    startIcon={<ContentCopyOutlinedIcon />}
                    onClick={() =>
                      void copySignatureImage(ubdRestoreFields.signatureData)
                    }
                  >
                    Копіювати
                  </Button>
                  <Button variant="text" onClick={clearUbdRestoreSignature}>
                    Прибрати
                  </Button>
                </>
              ) : null}
            </Stack>
            {ubdRestoreFields.signatureData ? (
              <div className="form12-signature-preview">
                <img alt="Підпис службовця" src={ubdRestoreFields.signatureData} />
                <small>{ubdRestoreFields.signatureFileName || "підпис.png"}</small>
              </div>
            ) : (
              <p className="salary-empty-documents">
                Завантажте PNG підпису — білий фон прибереться, у Word картинка
                стане прозорою над прізвищем службовця.
              </p>
            )}
          </div>
          {documentStatusNotePanel(ubdRestoreFields.statusNote, (value) =>
            updateUbdRestoreField("statusNote", value),
          )}
          {personQuestionnaireRow}
          <div className="ubd-scans-panel" id="ubd-scans-panel">
            <div className="panel-heading">Скани пошкодженого посвідчення</div>
            <Button component="label" variant="outlined">
              Додати скани
              <input
                hidden
                multiple
                type="file"
                accept="image/*,application/pdf,.pdf"
                onChange={(event) => {
                  void addUbdScanFiles(event.target.files);
                  event.target.value = "";
                }}
              />
            </Button>
            <div className="ubd-scan-list">
              {(documentFiles.ubdScans ?? []).length ? (
                (documentFiles.ubdScans ?? []).map((file) => (
                  <article className="ubd-scan-item" key={file.id}>
                    <a
                      href={file.dataUrl}
                      target="_blank"
                      rel="noreferrer"
                      onClick={(event) => {
                        event.preventDefault();
                        openDataUrlInNewTab(file.dataUrl);
                      }}
                    >
                      {file.type.startsWith("image/") ? (
                        <img alt={file.name} src={file.dataUrl} />
                      ) : (
                        <PictureAsPdfOutlinedIcon />
                      )}
                      <span>{file.name}</span>
                    </a>
                    <button
                      aria-label="Видалити скан"
                      className="documents-journal-delete"
                      onClick={() => deleteUbdScanFile(file.id)}
                      type="button"
                    >
                      <DeleteOutlineOutlinedIcon />
                    </button>
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Додайте скан пошкодженого посвідчення УБД.
                </p>
              )}
            </div>
          </div>
        </section>
      ) : null}

      {wordPreviewPanel}
    </>
  );

  const ticketDispatchText = formatTemporaryIdDispatchText(ticketFields);
  const ticketWorkspace = (
    <>
      {selectedDocument ? (
        <section className="analytics-panel document-fields">
          <div className="panel-heading">
            Дані для надсилання
            <IconButton
              aria-label="Скопіювати для сповіщення"
              className="ticket-copy-icon"
              size="small"
              title="Скопіювати для сповіщення"
              type="button"
              onClick={() => void copyTicketDispatchLine()}
            >
              <ContentCopyOutlinedIcon />
            </IconButton>
          </div>
          <div className="document-placeholder-map salary-document-map">
            <label>
              <span>Звання</span>
              <TextField
                size="small"
                fullWidth
                value={ticketFields.rank}
                onChange={(event) =>
                  updateTicketField("rank", event.target.value)
                }
              />
            </label>
            <label>
              <span>ПІБ</span>
              <TextField
                size="small"
                fullWidth
                value={ticketFields.fullName}
                onChange={(event) =>
                  updateTicketField("fullName", event.target.value)
                }
              />
            </label>
            <label>
              <span>Позивний</span>
              <TextField
                size="small"
                fullWidth
                value={ticketFields.callSign}
                onChange={(event) =>
                  updateTicketField("callSign", event.target.value)
                }
              />
            </label>
            <label>
              <span>Дата народження</span>
              <TextField
                size="small"
                fullWidth
                value={ticketFields.birthDate}
                onChange={(event) =>
                  updateTicketField("birthDate", event.target.value)
                }
                placeholder="дд.мм.рррр"
              />
            </label>
            <label className="wide">
              <span>Текст для сповіщення</span>
              <TextField
                size="small"
                fullWidth
                multiline
                rows={4}
                value={ticketDispatchText}
                readOnly
              />
            </label>
          </div>
          {documentStatusNotePanel(ticketFields.statusNote, (value) =>
            updateTicketField("statusNote", value),
          )}
          <div className="ticket-id-photo-actions">
            <Button
              component="label"
              variant="outlined"
              startIcon={<FileUploadOutlinedIcon />}
              disabled={isSavingDocument}
            >
              {ticketFields.photoData ? "Замінити фото" : "Додати фото"}
              <input
                hidden
                type="file"
                accept="image/*"
                onChange={(event) => {
                  void addTicketPhoto(event.target.files);
                  event.target.value = "";
                }}
              />
            </Button>
            <Button
              component="label"
              variant="outlined"
              startIcon={<FileUploadOutlinedIcon />}
              disabled={isSavingDocument}
            >
              Додати скан ТВК
              <input
                hidden
                multiple
                type="file"
                accept="image/*,application/pdf,.pdf"
                onChange={(event) => {
                  void addUbdScanFiles(event.target.files);
                  event.target.value = "";
                }}
              />
            </Button>
            <div className="ubd-scan-list" id="ticket-tvk-scans">
              {(documentFiles.ubdScans ?? []).length ? (
                (documentFiles.ubdScans ?? []).map((file) => (
                  <article className="ubd-scan-item" key={file.id}>
                    <a
                      href={file.dataUrl}
                      target="_blank"
                      rel="noreferrer"
                      onClick={(event) => {
                        event.preventDefault();
                        openDataUrlInNewTab(file.dataUrl);
                      }}
                    >
                      {file.type.startsWith("image/") ? (
                        <img alt={file.name} src={file.dataUrl} />
                      ) : (
                        <PictureAsPdfOutlinedIcon />
                      )}
                      <span>{file.name}</span>
                    </a>
                    <button
                      aria-label="Видалити скан ТВК"
                      className="documents-journal-delete"
                      onClick={() => deleteUbdScanFile(file.id)}
                      type="button"
                    >
                      <DeleteOutlineOutlinedIcon />
                    </button>
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Додайте скан тимчасового військового квитка.
                </p>
              )}
            </div>
          </div>
        </section>
      ) : null}

      {selectedDocument ? (
        <section className="analytics-panel document-preview">
          <div className="panel-heading">Фото та рядок</div>
          <div className="ticket-id-preview">
            <div className="ticket-id-photo-frame">
              {ticketFields.photoData ? (
                <img alt={ticketFields.fullName} src={ticketFields.photoData} />
              ) : (
                <span>Фото ще не додано</span>
              )}
            </div>
            <p className="ticket-id-dispatch-line">{ticketDispatchText}</p>
            {(documentFiles.ubdScans ?? []).length ? (
              <div className="ticket-tvk-preview">
                <div className="panel-heading">Скан ТВК</div>
                {(documentFiles.ubdScans ?? []).map((file) => (
                  <a
                    className="ticket-tvk-preview-item"
                    href={file.dataUrl}
                    key={file.id}
                    rel="noreferrer"
                    target="_blank"
                    onClick={(event) => {
                      event.preventDefault();
                      openDataUrlInNewTab(file.dataUrl);
                    }}
                  >
                    {file.type.startsWith("image/") ? (
                      <img alt={file.name} src={file.dataUrl} />
                    ) : (
                      <span className="ticket-tvk-preview-pdf">
                        <PictureAsPdfOutlinedIcon />
                        {file.name}
                      </span>
                    )}
                  </a>
                ))}
              </div>
            ) : null}
          </div>
        </section>
      ) : null}
    </>
  );

  const salaryWorkspace = (
    <>
      {selectedDocument ? (
        <section className="analytics-panel document-fields">
          <div className="panel-heading">Дані для рапорта</div>
          <div className="document-placeholder-map salary-document-map">
            <label>
              <code>ФИО</code>
              <TextField
                size="small"
                fullWidth
                value={salaryFields.fullName}
                onChange={(event) =>
                  updateSalaryField("fullName", event.target.value)
                }
                placeholder="ПРІЗВИЩЕ Ім'я По батькові"
              />
            </label>
            <label>
              <code>ІНН</code>
              <TextField
                size="small"
                fullWidth
                value={salaryFields.rnokpp}
                onChange={(event) =>
                  updateSalaryField("rnokpp", event.target.value)
                }
                placeholder="РНОКПП"
              />
            </label>
            <label>
              <code>IBAN</code>
              <TextField
                size="small"
                fullWidth
                value={formatIban(salaryFields.iban)}
                onChange={(event) =>
                  updateSalaryField("iban", event.target.value)
                }
                placeholder="UA..."
              />
            </label>
            <label>
              <code>БАНК</code>
              <TextField
                select
                size="small"
                fullWidth
                value={salaryFields.bankMfo}
                onChange={(event) =>
                  updateSalaryField("bankMfo", event.target.value)
                }
              >
                {ukrainianBanks.map((bank) => (
                  <MenuItem key={bank.mfo} value={bank.mfo}>
                    {bank.mfo === "custom"
                      ? bank.name
                      : `${bank.name} · ${bank.mfo}`}
                  </MenuItem>
                ))}
              </TextField>
            </label>
            <label>
              <code>Назва банку</code>
              <TextField
                size="small"
                fullWidth
                value={salaryFields.bankName}
                onChange={(event) =>
                  updateSalaryField("bankName", event.target.value)
                }
                placeholder="АТ «...»"
              />
            </label>
            <label>
              <code>Дата</code>
              <TextField
                size="small"
                fullWidth
                value={salaryFields.date}
                onChange={(event) =>
                  updateSalaryField("date", event.target.value)
                }
                placeholder="ДД.ММ.РРРР"
              />
            </label>
            <label>
              <code>Папка</code>
              <TextField
                size="small"
                fullWidth
                value={salaryFields.folderName}
                onChange={(event) =>
                  updateSalaryField("folderName", event.target.value)
                }
                placeholder="Назва папки"
              />
            </label>
          </div>
          {documentStatusNotePanel(salaryFields.statusNote, (value) =>
            updateSalaryField("statusNote", value),
          )}
        </section>
      ) : null}

      {selectedDocument ? (
        <section className="analytics-panel document-preview">
          <div className="panel-heading">Попередній перегляд</div>
          <div className="document-page-preview salary-document-preview">
            <div className="salary-preview-to">{salaryFields.commander}</div>
            <h2>Рапорт</h2>
            <p>
              Прошу моє грошове забезпечення перераховувати на мій розрахунковий
              банківський рахунок:
            </p>
            <p>
              {highlightOrPlaceholder(
                formatIban(salaryFields.iban),
                "UA___________________________",
              )}{" "}
              відкритий {highlightOrPlaceholder(salaryFields.bankName, "банк")}.
            </p>
            <p>
              Код РНОКПП{" "}
              {highlightOrPlaceholder(salaryFields.rnokpp, "__________")}.
            </p>
            <p>Довідка з банківськими реквізитами додається.</p>
            <p>
              {highlightOrPlaceholder(salaryFields.rank, "звання")}{" "}
              {highlightOrPlaceholder(
                salaryFields.fullName,
                "ПРІЗВИЩЕ Ім'я По батькові",
              )}
            </p>
            <footer>
              <span>{salaryFields.date || "__.__.202__"}</span>
              <span>________________</span>
            </footer>
            <p className="salary-finance-note">
              Начальнику фінансово-економічної служби – головному бухгалтеру
              військової частини А4862 - прошу прийняти в роботу
            </p>
            <p>
              Начальник групи персоналу штабу
              <br />1 піхотного батальйону
              <br />
              військової частини А4862
            </p>
            <p>{salaryFields.personnelChief}</p>
          </div>
        </section>
      ) : null}
    </>
  );

  const journalOpenDocument =
    (mode === "ubdReport" && selectedDocument?.type === "ubdReport") ||
    (mode === "ubdRestoreReport" &&
      selectedDocument?.type === "ubdRestoreReport") ||
    (mode === "form6Report" && selectedDocument?.type === "form6Report") ||
    (mode === "form12Report" && selectedDocument?.type === "form12Report") ||
    (mode === "serviceCharacteristic" &&
      selectedDocument?.type === "serviceCharacteristic") ||
    (mode === "zhbdCertificate" &&
      selectedDocument?.type === "zhbdCertificate") ||
    (mode === "temporaryMilitaryId" &&
      selectedDocument?.type === "temporaryMilitaryId") ||
    (mode === "salaryPowerAttorney" &&
      selectedDocument?.type === "salaryPowerAttorney");

  if (!isPersonDocumentMode) {
    return (
      <>
      <main
        className={
          journalOpenDocument
            ? "main-panel documents-journal-page documents-journal-with-ubd"
            : "main-panel documents-journal-page"
        }
      >
        <header className="topbar analytics-topbar salary-document-topbar">
          <Box className="salary-document-title">
            <Typography component="h1" variant="h4">
              Документи
            </Typography>
            <Typography variant="body2" color="text.secondary">
              Загальний журнал рапортів · прогрес по всіх службовцях
            </Typography>
          </Box>
          <Stack direction="row" spacing={1}>
            <Button
              variant="outlined"
              startIcon={<ArticleOutlinedIcon />}
              disabled={isLoadingDocumentJournal}
              onClick={() =>
                void loadDocumentJournal("Оновлюю журнал документів...")
              }
            >
              Оновити
            </Button>
            <Button
              variant="contained"
              startIcon={<FileDownloadOutlinedIcon />}
              disabled={
                isLoadingDocumentJournal ||
                isExportingDocumentJournal ||
                !filteredJournalDocuments.length
              }
              onClick={() => void exportDocumentJournal()}
              sx={{ color: "#1a1a14" }}
            >
              {isExportingDocumentJournal ? "Експорт..." : "Excel"}
            </Button>
            {journalTypeFilter === "ubdReport" ? (
              <>
                <Button
                  variant="outlined"
                  startIcon={<FileDownloadOutlinedIcon />}
                  disabled={
                    isLoadingDocumentJournal ||
                    isExportingDocumentJournal ||
                    !filteredJournalDocuments.length
                  }
                  onClick={() => void exportUbdRegistryTable()}
                >
                  УБД таблиця
                </Button>
                <Button
                  variant="outlined"
                  startIcon={<FileDownloadOutlinedIcon />}
                  disabled={
                    isLoadingDocumentJournal ||
                    isExportingDocumentJournal ||
                    !filteredJournalDocuments.length
                  }
                  onClick={() => void exportUbdStatusTable()}
                >
                  Статус
                </Button>
              </>
            ) : null}
          </Stack>
        </header>

        {isLoadingDocumentJournal ? <LinearProgress color="primary" /> : null}

        <Alert
          severity={isLoadingDocumentJournal ? "info" : "success"}
          variant="outlined"
          className="personnel-page-alert"
        >
          {documentMessage || "Журнал документів готовий."}
        </Alert>
        {selectedDocument && !canMutateSelected ? (
          <Alert severity="info" variant="outlined" className="personnel-page-alert">
            Лише перегляд документа · автор: {selectedDocumentAuthor}. Зміни,
            прогрес і видалення доступні лише автору.
          </Alert>
        ) : null}

        <section className="analytics-panel documents-journal-panel">
          <div className="panel-heading">
            {journalCreatorFilter === "MINE"
              ? "Мої документи"
              : journalCreatorFilter === "ALL"
                ? "Усі документи"
                : "Документи автора"}{" "}
            · {filteredJournalDocuments.length}
            {filteredJournalDocuments.length !== journalDocuments.length
              ? ` / ${journalDocuments.length}`
              : ""}
          </div>
          <div className="documents-journal-toolbar">
            <TextField
              size="small"
              label="ПІБ"
              placeholder="Пошук за прізвищем / імʼям"
              value={journalNameQuery}
              onChange={(event) => setJournalNameQuery(event.target.value)}
              className="documents-journal-search"
            />
            <TextField
              select
              size="small"
              label="Документ"
              value={journalTypeFilter}
              onChange={(event) => {
                const nextType = event.target.value;
                setJournalTypeFilter(nextType);
                const nextSteps =
                  nextType === "ALL"
                    ? [
                        ...ubdWorkflowSteps,
                        ...ubdRestoreWorkflowSteps,
                        ...form12WorkflowSteps,
                        ...temporaryMilitaryIdWorkflowSteps,
                        ...salaryWorkflowSteps,
                      ]
                    : documentWorkflowSteps(nextType);
                if (
                  journalStatusFilter !== "ALL" &&
                  !nextSteps.some((step) => step.title === journalStatusFilter)
                ) {
                  setJournalStatusFilter("ALL");
                }
              }}
            >
              {JOURNAL_DOCUMENT_TYPE_FILTERS.map((item) => (
                <MenuItem key={item.value} value={item.value}>
                  {item.label}
                </MenuItem>
              ))}
            </TextField>
            <TextField
              select
              size="small"
              label="Автор"
              value={journalCreatorFilter}
              onChange={(event) => setJournalCreatorFilter(event.target.value)}
            >
              <MenuItem value="MINE">Лише мої</MenuItem>
              <MenuItem value="ALL">Усі автори</MenuItem>
              {journalCreatorOptions
                .filter((creator) => {
                  const isSelf =
                    user?.id === creator.id ||
                    user?.email?.toLowerCase() ===
                      creator.email.toLowerCase();
                  return !isSelf;
                })
                .map((creator) => (
                  <MenuItem key={creator.id} value={creator.id}>
                    {creator.label}
                  </MenuItem>
                ))}
            </TextField>
            <TextField
              select
              size="small"
              label="Статус"
              value={journalStatusFilter}
              onChange={(event) => setJournalStatusFilter(event.target.value)}
            >
              <MenuItem value="ALL">Усі статуси</MenuItem>
              {journalStatusOptions.map((title) => (
                <MenuItem key={title} value={title}>
                  {title}
                </MenuItem>
              ))}
            </TextField>
            <TextField
              select
              size="small"
              label="Місяць створення"
              value={journalMonthFilter}
              onChange={(event) => setJournalMonthFilter(event.target.value)}
            >
              <MenuItem value="ALL">Усі місяці</MenuItem>
              {(journalMonthOptions.includes(journalMonthFilter) ||
              journalMonthFilter === "ALL"
                ? journalMonthOptions
                : [journalMonthFilter, ...journalMonthOptions]
              ).map((monthKey) => (
                <MenuItem key={monthKey} value={monthKey}>
                  {formatJournalMonthLabel(monthKey)}
                </MenuItem>
              ))}
            </TextField>
          </div>
          <div
            className={[
              "documents-journal-table",
              showUbdExitDateColumn ? "documents-journal-table-with-exit-date" : "",
            ]
              .filter(Boolean)
              .join(" ")}
          >
            <div
              className={[
                "documents-journal-row",
                "header",
                showUbdExitDateColumn ? "documents-journal-row-with-exit-date" : "",
              ]
                .filter(Boolean)
                .join(" ")}
            >
              <span>Службовець</span>
              <span>Статус службовця</span>
              <span>Документ</span>
              <span>Автор</span>
              <button
                type="button"
                className="documents-journal-sort"
                aria-label={
                  journalSortField === "progress" &&
                  journalSortDirection === "desc"
                    ? "Сортувати за прогресом: спочатку менший"
                    : "Сортувати за прогресом: спочатку більший"
                }
                title="Сортувати за прогресом"
                onClick={() => toggleJournalSort("progress")}
              >
                Прогрес
                {journalSortField === "progress"
                  ? journalSortDirection === "desc"
                    ? " ↓"
                    : " ↑"
                  : ""}
              </button>
              <span>Статус</span>
              <span>Коментар</span>
              <span>Файли</span>
              {showUbdExitDateColumn ? (
                <button
                  type="button"
                  className="documents-journal-sort"
                  aria-label={
                    journalSortField === "taskPeriodEnd" &&
                    journalSortDirection === "desc"
                      ? "Сортувати за датою виходу: спочатку старіші"
                      : "Сортувати за датою виходу: спочатку новіші"
                  }
                  title="Сортувати за датою виходу"
                  onClick={() => toggleJournalSort("taskPeriodEnd")}
                >
                  Вихід до
                  {journalSortField === "taskPeriodEnd"
                    ? journalSortDirection === "desc"
                      ? " ↓"
                      : " ↑"
                    : ""}
                </button>
              ) : null}
              <button
                type="button"
                className="documents-journal-sort"
                aria-label={
                  journalSortField === "createdAt" &&
                  journalSortDirection === "desc"
                    ? "Сортувати за датою створення: спочатку старіші"
                    : "Сортувати за датою створення: спочатку новіші"
                }
                title="Сортувати за датою створення"
                onClick={() => toggleJournalSort("createdAt")}
              >
                Створено
                {journalSortField === "createdAt"
                  ? journalSortDirection === "desc"
                    ? " ↓"
                    : " ↑"
                  : ""}
              </button>
              <button
                type="button"
                className="documents-journal-sort"
                aria-label={
                  journalSortField === "updatedAt" &&
                  journalSortDirection === "desc"
                    ? "Сортувати за датою оновлення: спочатку старіші"
                    : "Сортувати за датою оновлення: спочатку новіші"
                }
                title="Сортувати за датою оновлення"
                onClick={() => toggleJournalSort("updatedAt")}
              >
                Оновлено
                {journalSortField === "updatedAt"
                  ? journalSortDirection === "desc"
                    ? " ↓"
                    : " ↑"
                  : ""}
              </button>
              <span />
            </div>
            <div className="documents-journal-table-body">
            {isLoadingDocumentJournal ? (
              <div className="documents-journal-loader" role="status">
                <Spinner size="LG" label="ЗАВАНТАЖЕННЯ ЖУРНАЛУ" />
              </div>
            ) : filteredJournalDocuments.length ? (
              filteredJournalDocuments.map((document) => {
                const journalDocument =
                  document.id === selectedDocumentId
                    ? {
                        ...document,
                        status: workflow.currentStatus || document.status,
                        workflow: {
                          ...(document.workflow &&
                          typeof document.workflow === "object"
                            ? document.workflow
                            : {}),
                          ...workflow,
                        },
                      }
                    : document;
                const progress = getDocumentProgressPercent(journalDocument);
                const openDocument = () => {
                  void openPersonDocument(journalDocument).then(() => {
                    window.requestAnimationFrame(() => {
                      window.document
                        .getElementById("documents-journal-ubd")
                        ?.scrollIntoView({
                          behavior: "smooth",
                          block: "start",
                        });
                    });
                  });
                };
                return (
                  <div
                    className={[
                      "documents-journal-row",
                      showUbdExitDateColumn
                        ? "documents-journal-row-with-exit-date"
                        : "",
                      document.id === selectedDocumentId ? "active" : "",
                    ]
                      .filter(Boolean)
                      .join(" ")}
                    key={document.id}
                    role="button"
                    tabIndex={0}
                    onClick={openDocument}
                    onKeyDown={(event) => {
                      if (event.key !== "Enter" && event.key !== " ") return;
                      event.preventDefault();
                      openDocument();
                    }}
                  >
                    <strong>{getDocumentPersonName(document)}</strong>
                    <span
                      className="documents-journal-person-status"
                      title={readDocumentPersonStatus(
                        document,
                        personStatusById,
                      )}
                    >
                      {readDocumentPersonStatus(document, personStatusById)}
                    </span>
                    <span>{salaryDocumentTypeLabel(document.type)}</span>
                    <span
                      className="documents-journal-author"
                      title={document.createdByEmail || undefined}
                    >
                      {documentCreatorLabel(document)}
                      {isDocumentOwner(document) ? " · ви" : ""}
                    </span>
                    <span className="documents-journal-progress">
                      <i style={{ width: `${progress}%` }} />
                      <b>{progress}%</b>
                    </span>
                    <span>{documentWorkflowStatusLabel(journalDocument)}</span>
                    <span
                      className="documents-journal-note"
                      title={liveDocumentStatusNote(journalDocument)}
                    >
                      {liveDocumentStatusNote(journalDocument) || "—"}
                    </span>
                    <span>{getDocumentFileSummary(document)}</span>
                    {showUbdExitDateColumn ? (
                      <span title={readDocumentUbdTaskPeriod(document) || undefined}>
                        {readDocumentUbdTaskPeriodEndLabel(document) || "—"}
                      </span>
                    ) : null}
                    <span>
                      {new Date(document.createdAt).toLocaleString("uk-UA")}
                    </span>
                    <span>
                      {new Date(document.updatedAt).toLocaleString("uk-UA")}
                    </span>
                    {canMutateDocument(document) ? (
                      <button
                        aria-label="Видалити документ"
                        className="documents-journal-delete"
                        disabled={isSavingDocument}
                        onClick={(event) => {
                          event.stopPropagation();
                          void deleteDocument(document);
                        }}
                        title="Видалити документ"
                        type="button"
                      >
                        <DeleteOutlineOutlinedIcon />
                      </button>
                    ) : (
                      <span className="documents-journal-delete-slot" aria-hidden />
                    )}
                  </div>
                );
              })
            ) : (
              <div className="documents-journal-empty">
                {journalDocuments.length
                  ? journalCreatorFilter === "MINE"
                    ? "У вас ще немає документів у цьому фільтрі. Оберіть «Усі автори» або іншого автора для перегляду."
                    : "Немає документів за вибраними фільтрами."
                  : "Документи ще не створювались."}
              </div>
            )}
            </div>
          </div>
        </section>
        {mode === "ubdReport" && selectedDocument?.type === "ubdReport" ? (
          <section
            className={["documents-journal-ubd", foreignReadonlyClass]
              .filter(Boolean)
              .join(" ")}
            id="documents-journal-ubd"
          >
            <header className="topbar analytics-topbar salary-document-topbar">
              <div className="salary-document-main-row">
                <Box className="salary-document-title">
                  <Typography component="h1" variant="h4">
                    Рапорт на УБД
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    {getDocumentPersonName(selectedDocument)}
                    {selectedDocument.personExternalId
                      ? ` · ID ${selectedDocument.personExternalId}`
                      : ""}
                    {selectedDocument.title ? ` · ${selectedDocument.title}` : ""}
                    {` · автор: ${selectedDocumentAuthor}`}
                    {canMutateSelected ? "" : " · лише перегляд"}
                  </Typography>
                </Box>
                <Stack direction="row" spacing={1}>
                  {questionnaireHeaderButton}
                  <Button
                    variant="outlined"
                    startIcon={<ArticleOutlinedIcon />}
                    onClick={() => void saveUbdAsWord()}
                  >
                    Зберегти у Word
                  </Button>
                  <Button
                    variant="contained"
                    startIcon={<PictureAsPdfOutlinedIcon />}
                    onClick={printUbdDocument}
                    sx={{ color: "#1a1a14" }}
                  >
                    Друк / PDF
                  </Button>
                </Stack>
              </div>
              {documentStepProgress}
            </header>
            <section className="salary-documents-layout ubd-report-layout documents-journal-ubd-layout">
              {ubdWorkspace}
            </section>
          </section>
        ) : null}
        {mode === "form6Report" && selectedDocument?.type === "form6Report" ? (
          <section
            className={["documents-journal-ubd", foreignReadonlyClass]
              .filter(Boolean)
              .join(" ")}
            id="documents-journal-ubd"
          >
            <header className="topbar analytics-topbar salary-document-topbar">
              <div className="salary-document-main-row">
                <Box className="salary-document-title">
                  <Typography component="h1" variant="h4">
                    Форма 6
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    {getDocumentPersonName(selectedDocument)}
                    {selectedDocument.personExternalId
                      ? ` · ID ${selectedDocument.personExternalId}`
                      : ""}
                    {selectedDocument.title ? ` · ${selectedDocument.title}` : ""}
                  </Typography>
                </Box>
                <Stack direction="row" spacing={1}>
                  {questionnaireHeaderButton}
                  <Button
                    variant="outlined"
                    startIcon={<ArticleOutlinedIcon />}
                    onClick={() => void saveForm6AsWord()}
                  >
                    Зберегти у Word
                  </Button>
                </Stack>
              </div>
              {documentStepProgress}
            </header>
            <section className="salary-documents-layout ubd-report-layout documents-journal-ubd-layout">
              {form6Workspace}
            </section>
          </section>
        ) : null}
        {mode === "form12Report" && selectedDocument?.type === "form12Report" ? (
          <section
            className={["documents-journal-ubd", foreignReadonlyClass]
              .filter(Boolean)
              .join(" ")}
            id="documents-journal-ubd"
          >
            <header className="topbar analytics-topbar salary-document-topbar">
              <div className="salary-document-main-row">
                <Box className="salary-document-title">
                  <Typography component="h1" variant="h4">
                    Форма 12
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    {getDocumentPersonName(selectedDocument)}
                    {selectedDocument.personExternalId
                      ? ` · ID ${selectedDocument.personExternalId}`
                      : ""}
                    {selectedDocument.title ? ` · ${selectedDocument.title}` : ""}
                  </Typography>
                </Box>
                <Stack direction="row" spacing={1}>
                  {questionnaireHeaderButton}
                  <Button
                    variant="outlined"
                    startIcon={<ArticleOutlinedIcon />}
                    onClick={() => void saveForm12AsWord()}
                  >
                    Зберегти у Word
                  </Button>
                </Stack>
              </div>
              {documentStepProgress}
            </header>
            <section className="salary-documents-layout ubd-report-layout documents-journal-ubd-layout">
              {form12Workspace}
            </section>
          </section>
        ) : null}
        {mode === "serviceCharacteristic" &&
        selectedDocument?.type === "serviceCharacteristic" ? (
          <section
            className={["documents-journal-ubd", foreignReadonlyClass]
              .filter(Boolean)
              .join(" ")}
            id="documents-journal-ubd"
          >
            <header className="topbar analytics-topbar salary-document-topbar">
              <div className="salary-document-main-row">
                <Box className="salary-document-title">
                  <Typography component="h1" variant="h4">
                    Службова характеристика
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    {getDocumentPersonName(selectedDocument)}
                    {selectedDocument.personExternalId
                      ? ` · ID ${selectedDocument.personExternalId}`
                      : ""}
                    {selectedDocument.title ? ` · ${selectedDocument.title}` : ""}
                  </Typography>
                </Box>
                <Stack direction="row" spacing={1}>
                  {questionnaireHeaderButton}
                  <Button
                    variant="outlined"
                    startIcon={<ArticleOutlinedIcon />}
                    onClick={() => void saveServiceCharacteristicAsWord()}
                  >
                    Зберегти у Word
                  </Button>
                </Stack>
              </div>
              {documentStepProgress}
            </header>
            <section className="salary-documents-layout ubd-report-layout documents-journal-ubd-layout">
              {serviceCharacteristicWorkspace}
            </section>
          </section>
        ) : null}
        {mode === "zhbdCertificate" &&
        selectedDocument?.type === "zhbdCertificate" ? (
          <section
            className={["documents-journal-ubd", foreignReadonlyClass]
              .filter(Boolean)
              .join(" ")}
            id="documents-journal-ubd"
          >
            <header className="topbar analytics-topbar salary-document-topbar">
              <div className="salary-document-main-row">
                <Box className="salary-document-title">
                  <Typography component="h1" variant="h4">
                    Довідка ЖБД
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    {getDocumentPersonName(selectedDocument)}
                    {selectedDocument.personExternalId
                      ? ` · ID ${selectedDocument.personExternalId}`
                      : ""}
                    {selectedDocument.title ? ` · ${selectedDocument.title}` : ""}
                  </Typography>
                </Box>
                <Stack direction="row" spacing={1}>
                  {questionnaireHeaderButton}
                  <Button
                    variant="outlined"
                    startIcon={<ArticleOutlinedIcon />}
                    onClick={() => void saveZhbdCertificateAsWord()}
                  >
                    Зберегти у Word
                  </Button>
                </Stack>
              </div>
              {documentStepProgress}
            </header>
            <section className="salary-documents-layout ubd-report-layout documents-journal-ubd-layout">
              {zhbdCertificateWorkspace}
            </section>
          </section>
        ) : null}
        {mode === "ubdRestoreReport" &&
        selectedDocument?.type === "ubdRestoreReport" ? (
          <section
            className={["documents-journal-ubd", foreignReadonlyClass]
              .filter(Boolean)
              .join(" ")}
            id="documents-journal-ubd"
          >
            <header className="topbar analytics-topbar salary-document-topbar">
              <div className="salary-document-main-row">
                <Box className="salary-document-title">
                  <Typography component="h1" variant="h4">
                    Рапорт на відновлення УБД
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    {getDocumentPersonName(selectedDocument)}
                    {selectedDocument.personExternalId
                      ? ` · ID ${selectedDocument.personExternalId}`
                      : ""}
                    {selectedDocument.title ? ` · ${selectedDocument.title}` : ""}
                  </Typography>
                </Box>
                <Stack direction="row" spacing={1}>
                  {questionnaireHeaderButton}
                  <Button
                    variant="outlined"
                    startIcon={<ArticleOutlinedIcon />}
                    onClick={() => void saveUbdRestoreAsWord()}
                  >
                    Зберегти у Word
                  </Button>
                </Stack>
              </div>
              {documentStepProgress}
            </header>
            <section className="salary-documents-layout ubd-report-layout documents-journal-ubd-layout">
              {ubdRestoreWorkspace}
            </section>
          </section>
        ) : null}
        {mode === "temporaryMilitaryId" &&
        selectedDocument?.type === "temporaryMilitaryId" ? (
          <section
            className={["documents-journal-ubd", foreignReadonlyClass]
              .filter(Boolean)
              .join(" ")}
            id="documents-journal-ubd"
          >
            <header className="topbar analytics-topbar salary-document-topbar">
              <div className="salary-document-main-row">
                <Box className="salary-document-title">
                  <Typography component="h1" variant="h4">
                    Тимчасовий військовий квиток
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    {getDocumentPersonName(selectedDocument)}
                    {selectedDocument.personExternalId
                      ? ` · ID ${selectedDocument.personExternalId}`
                      : ""}
                    {selectedDocument.title ? ` · ${selectedDocument.title}` : ""}
                  </Typography>
                </Box>
                {questionnaireHeaderButton}
              </div>
              {documentStepProgress}
            </header>
            <section className="salary-documents-layout ticket-id-layout documents-journal-ubd-layout">
              {ticketWorkspace}
            </section>
          </section>
        ) : null}
        {mode === "salaryPowerAttorney" &&
        selectedDocument?.type === "salaryPowerAttorney" ? (
          <section
            className={["documents-journal-ubd", foreignReadonlyClass]
              .filter(Boolean)
              .join(" ")}
            id="documents-journal-ubd"
          >
            <header className="topbar analytics-topbar salary-document-topbar">
              <div className="salary-document-main-row">
                <Box className="salary-document-title">
                  <Typography component="h1" variant="h4">
                    Довіреність зарплати
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    {getDocumentPersonName(selectedDocument)}
                    {selectedDocument.personExternalId
                      ? ` · ID ${selectedDocument.personExternalId}`
                      : ""}
                    {selectedDocument.title ? ` · ${selectedDocument.title}` : ""}
                  </Typography>
                </Box>
                <Stack direction="row" spacing={1}>
                  {questionnaireHeaderButton}
                  <Button
                    variant="outlined"
                    startIcon={<ArticleOutlinedIcon />}
                    onClick={generateSalaryDoc}
                  >
                    Завантажити DOC
                  </Button>
                  <Button
                    variant="contained"
                    startIcon={<PictureAsPdfOutlinedIcon />}
                    onClick={printSalaryDocument}
                    sx={{ color: "#1a1a14" }}
                  >
                    Друк / PDF
                  </Button>
                </Stack>
              </div>
              {documentStepProgress}
            </header>
            <section className="salary-documents-layout documents-journal-ubd-layout">
              {salaryWorkspace}
            </section>
          </section>
        ) : null}
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  if (mode === "temporaryMilitaryId") {
    return (
      <>
      <main
        className={["main-panel", "documents-ubd-page", foreignReadonlyClass]
          .filter(Boolean)
          .join(" ")}
      >
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Тимчасовий військовий квиток
              </Typography>
              <Typography variant="body2" color="text.secondary">
                {summary.name} ·{" "}
                {summary.externalId
                  ? `ID ${summary.externalId}`
                  : "дані з картки особи"}
                {selectedDocument ? ` · ${selectedDocument.title}` : ""}
              </Typography>
            </Box>
            {questionnaireHeaderButton}
          </div>
          {documentStepProgress}
        </header>

        <section
          className={
            selectedDocument
              ? "salary-documents-layout ticket-id-layout"
              : "salary-documents-layout empty"
          }
        >
          <section className="analytics-panel salary-person-documents-panel">
            <div className="panel-heading">Документи службовця</div>
            <div className="salary-document-sync">
              <span>
                {isSavingDocument ? "збереження..." : "БД синхронізована"}
              </span>
              <p>{documentMessage || "Виберіть документ зі списку."}</p>
            </div>
            <div className="salary-person-documents-list">
              {personQuestionnaireRow}
              {visiblePersonDocuments.length ? (
                visiblePersonDocuments.map((document) => (
                  <article
                    className={
                      document.id === selectedDocumentId
                        ? "salary-person-document-shell active"
                        : "salary-person-document-shell"
                    }
                    key={document.id}
                  >
                    <button
                      className="salary-person-document"
                      type="button"
                      onClick={() => openPersonDocument(document)}
                    >
                      <strong>{document.title}</strong>
                      <span>{salaryDocumentTypeLabel(document.type)}</span>
                      <em>{documentWorkflowStatusLabel(document)}</em>
                      {documentListStatusNote(document)}
                      <small>
                        {new Date(document.updatedAt).toLocaleString("uk-UA")}
                      </small>
                    </button>
                    {canMutateDocument(document) ? (
                      <button
                        aria-label="Видалити документ"
                        className="salary-person-document-delete"
                        disabled={isSavingDocument}
                        onClick={() => void deleteDocument(document)}
                        title="Видалити документ"
                        type="button"
                      >
                        <DeleteOutlineOutlinedIcon />
                      </button>
                    ) : null}
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Документів по цьому службовцю ще немає.
                </p>
              )}
            </div>
            {canEdit ? personDocumentCreateSelect : null}
          </section>

          {ticketWorkspace}
        </section>
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  if (mode === "ubdReport") {
    return (
      <>
      <main
        className={["main-panel", "documents-ubd-page", foreignReadonlyClass]
          .filter(Boolean)
          .join(" ")}
      >
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Рапорт на УБД
              </Typography>
              <Typography variant="body2" color="text.secondary">
                {summary.name} ·{" "}
                {summary.externalId
                  ? `ID ${summary.externalId}`
                  : "дані з картки особи"}
                {selectedDocument ? ` · ${selectedDocument.title}` : ""}
              </Typography>
            </Box>
            {selectedDocument ? (
              <Stack direction="row" spacing={1}>
                {questionnaireHeaderButton}
                <Button
                  variant="outlined"
                  startIcon={<ArticleOutlinedIcon />}
                  onClick={() => void saveUbdAsWord()}
                >
                  Зберегти у Word
                </Button>
                <Button
                  variant="contained"
                  startIcon={<PictureAsPdfOutlinedIcon />}
                  onClick={printUbdDocument}
                  sx={{ color: "#1a1a14" }}
                >
                  Друк / PDF
                </Button>
              </Stack>
            ) : questionnaireHeaderButton}
          </div>
          {documentStepProgress}
        </header>

        <section
          className={
            selectedDocument
              ? "salary-documents-layout ubd-report-layout"
              : "salary-documents-layout empty"
          }
        >
          <section className="analytics-panel salary-person-documents-panel">
            <div className="panel-heading">Документи службовця</div>
            <div className="salary-document-sync">
              <span>
                {isSavingDocument ? "збереження..." : "БД синхронізована"}
              </span>
              <p>{documentMessage || "Виберіть документ зі списку."}</p>
            </div>
            <div className="salary-person-documents-list">
              {personQuestionnaireRow}
              {visiblePersonDocuments.length ? (
                visiblePersonDocuments.map((document) => (
                  <article
                    className={
                      document.id === selectedDocumentId
                        ? "salary-person-document-shell active"
                        : "salary-person-document-shell"
                    }
                    key={document.id}
                  >
                    <button
                      className="salary-person-document"
                      type="button"
                      onClick={() => openPersonDocument(document)}
                    >
                      <strong>{document.title}</strong>
                      <span>{salaryDocumentTypeLabel(document.type)}</span>
                      <em>{documentWorkflowStatusLabel(document)}</em>
                      {documentListStatusNote(document)}
                      <small>
                        {new Date(document.updatedAt).toLocaleString("uk-UA")}
                      </small>
                    </button>
                    {canMutateDocument(document) ? (
                      <button
                        aria-label="Видалити документ"
                        className="salary-person-document-delete"
                        disabled={isSavingDocument}
                        onClick={() => void deleteDocument(document)}
                        title="Видалити документ"
                        type="button"
                      >
                        <DeleteOutlineOutlinedIcon />
                      </button>
                    ) : null}
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Документів по цьому службовцю ще немає.
                </p>
              )}
            </div>
            {canEdit ? personDocumentCreateSelect : null}
          </section>

          {ubdWorkspace}
        </section>
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  if (mode === "form6Report") {
    return (
      <>
      <main
        className={["main-panel", "documents-ubd-page", foreignReadonlyClass]
          .filter(Boolean)
          .join(" ")}
      >
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Форма 6
              </Typography>
              <Typography variant="body2" color="text.secondary">
                {summary.name} ·{" "}
                {summary.externalId
                  ? `ID ${summary.externalId}`
                  : "дані з картки особи"}
                {selectedDocument ? ` · ${selectedDocument.title}` : ""}
              </Typography>
            </Box>
            {selectedDocument ? (
              <Stack direction="row" spacing={1}>
                {questionnaireHeaderButton}
                <Button
                  variant="outlined"
                  startIcon={<ArticleOutlinedIcon />}
                  onClick={() => void saveForm6AsWord()}
                >
                  Зберегти у Word
                </Button>
              </Stack>
            ) : questionnaireHeaderButton}
          </div>
          {documentStepProgress}
        </header>

        <section
          className={
            selectedDocument
              ? "salary-documents-layout ubd-report-layout"
              : "salary-documents-layout empty"
          }
        >
          <section className="analytics-panel salary-person-documents-panel">
            <div className="panel-heading">Документи службовця</div>
            <div className="salary-document-sync">
              <span>
                {isSavingDocument ? "збереження..." : "БД синхронізована"}
              </span>
              <p>{documentMessage || "Виберіть документ зі списку."}</p>
            </div>
            <div className="salary-person-documents-list">
              {personQuestionnaireRow}
              {visiblePersonDocuments.length ? (
                visiblePersonDocuments.map((document) => (
                  <article
                    className={
                      document.id === selectedDocumentId
                        ? "salary-person-document-shell active"
                        : "salary-person-document-shell"
                    }
                    key={document.id}
                  >
                    <button
                      className="salary-person-document"
                      type="button"
                      onClick={() => openPersonDocument(document)}
                    >
                      <strong>{document.title}</strong>
                      <span>{salaryDocumentTypeLabel(document.type)}</span>
                      <em>{documentWorkflowStatusLabel(document)}</em>
                      {documentListStatusNote(document)}
                      <small>
                        {new Date(document.updatedAt).toLocaleString("uk-UA")}
                      </small>
                    </button>
                    {canMutateDocument(document) ? (
                      <button
                        aria-label="Видалити документ"
                        className="salary-person-document-delete"
                        disabled={isSavingDocument}
                        onClick={() => void deleteDocument(document)}
                        title="Видалити документ"
                        type="button"
                      >
                        <DeleteOutlineOutlinedIcon />
                      </button>
                    ) : null}
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Документів по цьому службовцю ще немає.
                </p>
              )}
            </div>
            {canEdit ? personDocumentCreateSelect : null}
          </section>

          {form6Workspace}
        </section>
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  if (mode === "form12Report") {
    return (
      <>
      <main
        className={["main-panel", "documents-ubd-page", foreignReadonlyClass]
          .filter(Boolean)
          .join(" ")}
      >
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Форма 12
              </Typography>
              <Typography variant="body2" color="text.secondary">
                {summary.name} ·{" "}
                {summary.externalId
                  ? `ID ${summary.externalId}`
                  : "дані з картки особи"}
                {selectedDocument ? ` · ${selectedDocument.title}` : ""}
              </Typography>
            </Box>
            {selectedDocument ? (
              <Stack direction="row" spacing={1}>
                {questionnaireHeaderButton}
                <Button
                  variant="outlined"
                  startIcon={<ArticleOutlinedIcon />}
                  onClick={() => void saveForm12AsWord()}
                >
                  Зберегти у Word
                </Button>
              </Stack>
            ) : questionnaireHeaderButton}
          </div>
          {documentStepProgress}
        </header>

        <section
          className={
            selectedDocument
              ? "salary-documents-layout ubd-report-layout"
              : "salary-documents-layout empty"
          }
        >
          <section className="analytics-panel salary-person-documents-panel">
            <div className="panel-heading">Документи службовця</div>
            <div className="salary-document-sync">
              <span>
                {isSavingDocument ? "збереження..." : "БД синхронізована"}
              </span>
              <p>{documentMessage || "Виберіть документ зі списку."}</p>
            </div>
            <div className="salary-person-documents-list">
              {personQuestionnaireRow}
              {visiblePersonDocuments.length ? (
                visiblePersonDocuments.map((document) => (
                  <article
                    className={
                      document.id === selectedDocumentId
                        ? "salary-person-document-shell active"
                        : "salary-person-document-shell"
                    }
                    key={document.id}
                  >
                    <button
                      className="salary-person-document"
                      type="button"
                      onClick={() => openPersonDocument(document)}
                    >
                      <strong>{document.title}</strong>
                      <span>{salaryDocumentTypeLabel(document.type)}</span>
                      <em>{documentWorkflowStatusLabel(document)}</em>
                      {documentListStatusNote(document)}
                      <small>
                        {new Date(document.updatedAt).toLocaleString("uk-UA")}
                      </small>
                    </button>
                    {canMutateDocument(document) ? (
                      <button
                        aria-label="Видалити документ"
                        className="salary-person-document-delete"
                        disabled={isSavingDocument}
                        onClick={() => void deleteDocument(document)}
                        title="Видалити документ"
                        type="button"
                      >
                        <DeleteOutlineOutlinedIcon />
                      </button>
                    ) : null}
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Документів по цьому службовцю ще немає.
                </p>
              )}
            </div>
            {canEdit ? personDocumentCreateSelect : null}
          </section>

          {form12Workspace}
        </section>
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  if (mode === "serviceCharacteristic") {
    return (
      <>
      <main
        className={["main-panel", "documents-ubd-page", foreignReadonlyClass]
          .filter(Boolean)
          .join(" ")}
      >
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Службова характеристика
              </Typography>
              <Typography variant="body2" color="text.secondary">
                {summary.name} ·{" "}
                {summary.externalId
                  ? `ID ${summary.externalId}`
                  : "дані з картки особи"}
                {selectedDocument ? ` · ${selectedDocument.title}` : ""}
              </Typography>
            </Box>
            {selectedDocument ? (
              <Stack direction="row" spacing={1}>
                {questionnaireHeaderButton}
                <Button
                  variant="outlined"
                  startIcon={<ArticleOutlinedIcon />}
                  onClick={() => void saveServiceCharacteristicAsWord()}
                >
                  Зберегти у Word
                </Button>
              </Stack>
            ) : (
              questionnaireHeaderButton
            )}
          </div>
          {documentStepProgress}
        </header>

        <section
          className={
            selectedDocument
              ? "salary-documents-layout ubd-report-layout"
              : "salary-documents-layout empty"
          }
        >
          <section className="analytics-panel salary-person-documents-panel">
            <div className="panel-heading">Документи службовця</div>
            <div className="salary-document-sync">
              <span>
                {isSavingDocument ? "збереження..." : "БД синхронізована"}
              </span>
              <p>{documentMessage || "Виберіть документ зі списку."}</p>
            </div>
            <div className="salary-person-documents-list">
              {personQuestionnaireRow}
              {visiblePersonDocuments.length ? (
                visiblePersonDocuments.map((document) => (
                  <article
                    className={
                      document.id === selectedDocumentId
                        ? "salary-person-document-shell active"
                        : "salary-person-document-shell"
                    }
                    key={document.id}
                  >
                    <button
                      className="salary-person-document"
                      type="button"
                      onClick={() => openPersonDocument(document)}
                    >
                      <strong>{document.title}</strong>
                      <span>{salaryDocumentTypeLabel(document.type)}</span>
                      <em>{documentWorkflowStatusLabel(document)}</em>
                      {documentListStatusNote(document)}
                      <small>
                        {new Date(document.updatedAt).toLocaleString("uk-UA")}
                      </small>
                    </button>
                    {canMutateDocument(document) ? (
                      <button
                        aria-label="Видалити документ"
                        className="salary-person-document-delete"
                        disabled={isSavingDocument}
                        onClick={() => void deleteDocument(document)}
                        title="Видалити документ"
                        type="button"
                      >
                        <DeleteOutlineOutlinedIcon />
                      </button>
                    ) : null}
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Документів по цьому службовцю ще немає.
                </p>
              )}
            </div>
            {canEdit ? personDocumentCreateSelect : null}
          </section>

          {serviceCharacteristicWorkspace}
        </section>
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  if (mode === "zhbdCertificate") {
    return (
      <>
      <main
        className={["main-panel", "documents-ubd-page", foreignReadonlyClass]
          .filter(Boolean)
          .join(" ")}
      >
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Довідка ЖБД
              </Typography>
              <Typography variant="body2" color="text.secondary">
                {summary.name} ·{" "}
                {summary.externalId
                  ? `ID ${summary.externalId}`
                  : "дані з картки особи"}
                {selectedDocument ? ` · ${selectedDocument.title}` : ""}
              </Typography>
            </Box>
            {selectedDocument ? (
              <Stack direction="row" spacing={1}>
                {questionnaireHeaderButton}
                <Button
                  variant="outlined"
                  startIcon={<ArticleOutlinedIcon />}
                  onClick={() => void saveZhbdCertificateAsWord()}
                >
                  Зберегти у Word
                </Button>
              </Stack>
            ) : (
              questionnaireHeaderButton
            )}
          </div>
          {documentStepProgress}
        </header>

        <section
          className={
            selectedDocument
              ? "salary-documents-layout ubd-report-layout"
              : "salary-documents-layout empty"
          }
        >
          <section className="analytics-panel salary-person-documents-panel">
            <div className="panel-heading">Документи службовця</div>
            <div className="salary-document-sync">
              <span>
                {isSavingDocument ? "збереження..." : "БД синхронізована"}
              </span>
              <p>{documentMessage || "Виберіть документ зі списку."}</p>
            </div>
            <div className="salary-person-documents-list">
              {personQuestionnaireRow}
              {visiblePersonDocuments.length ? (
                visiblePersonDocuments.map((document) => (
                  <article
                    className={
                      document.id === selectedDocumentId
                        ? "salary-person-document-shell active"
                        : "salary-person-document-shell"
                    }
                    key={document.id}
                  >
                    <button
                      className="salary-person-document"
                      type="button"
                      onClick={() => openPersonDocument(document)}
                    >
                      <strong>{document.title}</strong>
                      <span>{salaryDocumentTypeLabel(document.type)}</span>
                      <em>{documentWorkflowStatusLabel(document)}</em>
                      {documentListStatusNote(document)}
                      <small>
                        {new Date(document.updatedAt).toLocaleString("uk-UA")}
                      </small>
                    </button>
                    {canMutateDocument(document) ? (
                      <button
                        aria-label="Видалити документ"
                        className="salary-person-document-delete"
                        disabled={isSavingDocument}
                        onClick={() => void deleteDocument(document)}
                        title="Видалити документ"
                        type="button"
                      >
                        <DeleteOutlineOutlinedIcon />
                      </button>
                    ) : null}
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Документів по цьому службовцю ще немає.
                </p>
              )}
            </div>
            {canEdit ? personDocumentCreateSelect : null}
          </section>

          {zhbdCertificateWorkspace}
        </section>
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  if (mode === "ubdRestoreReport") {
    return (
      <>
      <main
        className={["main-panel", "documents-ubd-page", foreignReadonlyClass]
          .filter(Boolean)
          .join(" ")}
      >
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Рапорт на відновлення УБД
              </Typography>
              <Typography variant="body2" color="text.secondary">
                {summary.name} ·{" "}
                {summary.externalId
                  ? `ID ${summary.externalId}`
                  : "дані з картки особи"}
                {selectedDocument ? ` · ${selectedDocument.title}` : ""}
              </Typography>
            </Box>
            {selectedDocument ? (
              <Stack direction="row" spacing={1}>
                {questionnaireHeaderButton}
                <Button
                  variant="outlined"
                  startIcon={<ArticleOutlinedIcon />}
                  onClick={() => void saveUbdRestoreAsWord()}
                >
                  Зберегти у Word
                </Button>
              </Stack>
            ) : questionnaireHeaderButton}
          </div>
          {documentStepProgress}
        </header>

        <section
          className={
            selectedDocument
              ? "salary-documents-layout ubd-report-layout"
              : "salary-documents-layout empty"
          }
        >
          <section className="analytics-panel salary-person-documents-panel">
            <div className="panel-heading">Документи службовця</div>
            <div className="salary-document-sync">
              <span>
                {isSavingDocument ? "збереження..." : "БД синхронізована"}
              </span>
              <p>{documentMessage || "Виберіть документ зі списку."}</p>
            </div>
            <div className="salary-person-documents-list">
              {personQuestionnaireRow}
              {visiblePersonDocuments.length ? (
                visiblePersonDocuments.map((document) => (
                  <article
                    className={
                      document.id === selectedDocumentId
                        ? "salary-person-document-shell active"
                        : "salary-person-document-shell"
                    }
                    key={document.id}
                  >
                    <button
                      className="salary-person-document"
                      type="button"
                      onClick={() => openPersonDocument(document)}
                    >
                      <strong>{document.title}</strong>
                      <span>{salaryDocumentTypeLabel(document.type)}</span>
                      <em>{documentWorkflowStatusLabel(document)}</em>
                      {documentListStatusNote(document)}
                      <small>
                        {new Date(document.updatedAt).toLocaleString("uk-UA")}
                      </small>
                    </button>
                    {canMutateDocument(document) ? (
                      <button
                        aria-label="Видалити документ"
                        className="salary-person-document-delete"
                        disabled={isSavingDocument}
                        onClick={() => void deleteDocument(document)}
                        title="Видалити документ"
                        type="button"
                      >
                        <DeleteOutlineOutlinedIcon />
                      </button>
                    ) : null}
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Документів по цьому службовцю ще немає.
                </p>
              )}
            </div>
            {canEdit ? personDocumentCreateSelect : null}
          </section>

          {ubdRestoreWorkspace}
        </section>
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  if (mode === "salaryPowerAttorney") {
    return (
      <>
      <main className="main-panel">
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Довіреність зарплати
              </Typography>
              <Typography variant="body2" color="text.secondary">
                {summary.name} ·{" "}
                {summary.externalId
                  ? `ID ${summary.externalId}`
                  : "дані з картки особи"}
                {selectedDocument ? ` · ${selectedDocument.title}` : ""}
              </Typography>
            </Box>
            {selectedDocument ? (
              <Stack direction="row" spacing={1}>
                {questionnaireHeaderButton}
                <Button
                  variant="outlined"
                  startIcon={<ArticleOutlinedIcon />}
                  onClick={generateSalaryDoc}
                >
                  Завантажити DOC
                </Button>
                <Button
                  variant="contained"
                  startIcon={<PictureAsPdfOutlinedIcon />}
                  onClick={printSalaryDocument}
                  sx={{ color: "#1a1a14" }}
                >
                  Друк / PDF
                </Button>
              </Stack>
            ) : null}
          </div>
          {documentStepProgress}
        </header>

        <section
          className={
            selectedDocument
              ? "salary-documents-layout"
              : "salary-documents-layout empty"
          }
        >
          <section className="analytics-panel salary-person-documents-panel">
            <div className="panel-heading">Документи службовця</div>
            <div className="salary-document-sync">
              <span>
                {isSavingDocument ? "збереження..." : "БД синхронізована"}
              </span>
              <p>{documentMessage || "Виберіть документ зі списку."}</p>
            </div>
            <div className="salary-person-documents-list">
              {personQuestionnaireRow}
              {visiblePersonDocuments.length ? (
                visiblePersonDocuments.map((document) => (
                  <article
                    className={
                      document.id === selectedDocumentId
                        ? "salary-person-document-shell active"
                        : "salary-person-document-shell"
                    }
                    key={document.id}
                  >
                    <button
                      className="salary-person-document"
                      type="button"
                      onClick={() => openPersonDocument(document)}
                    >
                      <strong>{document.title}</strong>
                      <span>{salaryDocumentTypeLabel(document.type)}</span>
                      <em>{workflowStatusLabel(document.status)}</em>
                      {documentListStatusNote(document)}
                      <small>
                        {new Date(document.updatedAt).toLocaleString("uk-UA")}
                      </small>
                    </button>
                    <button
                      aria-label="Видалити документ"
                      className="salary-person-document-delete"
                      disabled={isSavingDocument}
                      onClick={() => void deleteDocument(document)}
                      title="Видалити документ"
                      type="button"
                    >
                      <DeleteOutlineOutlinedIcon />
                    </button>
                  </article>
                ))
              ) : (
                <p className="salary-empty-documents">
                  Документів по цьому службовцю ще немає.
                </p>
              )}
            </div>
            {canEdit ? personDocumentCreateSelect : null}
          </section>

          {salaryWorkspace}
        </section>
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  return (
    <main className="main-panel">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Документи та звіти
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Шаблони DOCX · джерело даних з картки особи
          </Typography>
        </Box>
        <Stack direction="row" spacing={1}>
          <Button variant="outlined" startIcon={<ArticleOutlinedIcon />}>
            Додати DOCX
          </Button>
          <Button variant="contained" sx={{ color: "#1a1a14" }}>
            Згенерувати DOCX
          </Button>
        </Stack>
      </header>

      <section className="documents-layout">
        <aside className="analytics-panel document-template-list">
          <div className="panel-heading">Бібліотека шаблонів</div>
          {["Довідка", "Наказ", "Витяг", "Рапорт", "Власний шаблон"].map(
            (template, index) => (
              <button
                className={index === 0 ? "active" : ""}
                key={template}
                type="button"
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>{template}</strong>
                  <small>
                    {index === 0
                      ? "Довідка про проходження служби"
                      : "Шаблон документа"}
                  </small>
                </span>
              </button>
            ),
          )}
        </aside>

        <section className="analytics-panel document-fields">
          <div className="panel-heading">Джерело даних</div>
          <div className="person-action-fields">
            <span>
              <strong>Особа</strong>
              {summary.name}
            </span>
            <span>
              <strong>Звання</strong>
              {summary.rank || "—"}
            </span>
            <span>
              <strong>Підрозділ / місце</strong>
              {summary.location || "—"}
            </span>
          </div>

          <div className="document-placeholder-map">
            <label>
              <code>{"{{ПІБ}}"}</code>
              <TextField
                size="small"
                fullWidth
                value={fields.pib}
                onChange={(event) =>
                  updateDefaultField("pib", event.target.value)
                }
                placeholder="Прізвище Ім'я По батькові"
              />
            </label>
            <label>
              <code>{"{{Звання}}"}</code>
              <TextField
                size="small"
                fullWidth
                value={fields.rank}
                onChange={(event) =>
                  updateDefaultField("rank", event.target.value)
                }
                placeholder="Звання"
              />
            </label>
            <label>
              <code>{"{{Підрозділ}}"}</code>
              <TextField
                size="small"
                fullWidth
                value={fields.unit}
                onChange={(event) =>
                  updateDefaultField("unit", event.target.value)
                }
                placeholder="Підрозділ / частина"
              />
            </label>
            <label>
              <code>{"{{Дата}}"}</code>
              <TextField
                size="small"
                fullWidth
                value={fields.date}
                onChange={(event) =>
                  updateDefaultField("date", event.target.value)
                }
                placeholder="ДД.ММ.РРРР"
              />
            </label>
          </div>
        </section>

        <section className="analytics-panel document-preview">
          <div className="panel-heading">Попередній перегляд документа</div>
          <div className="document-page-preview">
            <h2>ДОВІДКА</h2>
            <h3>про проходження служби та займану посаду</h3>
            <p>
              Видана {highlightOrPlaceholder(fields.pib, "{{ПІБ}}")},{" "}
              {highlightOrPlaceholder(fields.rank, "{{Звання}}")},
              військовослужбовцю, що проходить військову службу у{" "}
              {highlightOrPlaceholder(fields.unit, "{{Підрозділ}}")}.
            </p>
            <p>
              Дата видачі: {highlightOrPlaceholder(fields.date, "{{Дата}}")}.
            </p>
            <p>Довідка видана для пред’явлення за місцем вимоги.</p>
            <footer>
              <span>Командир</span>
              <span>________________</span>
            </footer>
          </div>
        </section>
      </section>
    </main>
  );
}
