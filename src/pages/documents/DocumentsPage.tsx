import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Alert,
  Box,
  Button,
  Checkbox,
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
  type BackendPersonnelRosterLatest,
} from "../../api";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
  readDataCache,
} from "../../data/idbDataCache";
import { buildDocumentRoute } from "../../app/navigation";
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
import { compressPhotoFile } from "../personnel/photoCompression";
import {
  loadPersonDocumentsForRow,
  loadPersonQuestionnaireForRow,
} from "../personnel/personAttachments";
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
import { personNameFromSyntheticDocumentId } from "./documentPersonIdentity";
import {
  buildUbdNotSubmittedRowFromDocument,
  buildUbdRosterCallSignIndex,
  exportUbdNotSubmittedExcel,
  lookupUbdRosterCallSign,
} from "./ubdNotSubmittedExcel";
import type { StaffSheetImportSnapshot } from "../anketa-data/staffSheetImport";
import {
  formatUbdBasisOrderLabel,
  findUbdBasisOrderByKey,
  pickUbdBasisOrderForTaskPeriod,
  resolveUbdBasisForTask,
  ubdBasisDateMatchesTaskPeriod,
  ubdBasisOrderOptionKey,
  ubdHasExactBasisForTaskPeriod,
} from "./ubdBasisOrders";
import { allBasisOrderOptions } from "./ubdBasisOrdersDirectory";
import {
  documentBasisFieldHighlightClass,
  documentFieldLabelClass,
  documentHasBasisDateMismatch,
  documentHasEmptyInputs,
  documentRequiredFieldIsBlank,
  isBlankForm6IdDocument,
  isBlankUbdRnokpp,
  readDocumentSkippedDueToSzch,
  resolveUbdFieldsForGapCheck,
} from "./documentFieldReadiness";
import {
  buildFighterTaskPeriodText,
  formatUbdTaskPeriodStartDate,
  getFighterTaskPlace,
  normalizeRosterMatchText,
  parseUbdTaskPeriodStartDate,
} from "../personnel/fighterStatusImport";
import { mapRosterLatestToPreviewRows } from "../excel-fill/rosterSourceSnapshot";
import { Spinner } from "@/components/ui/spinner/spinner";
import {
  copyPngDataUrlToClipboard,
  getCommanderSignatureTransparent,
  processSignatureTransparentBackground,
} from "./ubdSignatureImage";
import { createUbdWordBlob } from "./ubdWordExport";
import { createUbdBulkWordBlob } from "./ubdBulkWordExport";
import {
  documentToUbdProgressBackupEntry,
  formatUbdProgressBackupWhen,
  inferUbdWorkflowCompletedFromArtifacts,
  loadPersistedUbdBulkProgressUndo,
  loadUbdProgressBackups,
  savePersistedUbdBulkProgressUndo,
  saveSessionJournalUbdBackup,
  saveUbdProgressBackup,
  type UbdProgressBackup,
  type UbdProgressBackupEntry,
} from "./ubdProgressBackup";
import {
  createForm6Fields,
  extractForm6BasisParts,
  formatForm6BasisText,
  form6WorkflowSteps,
  mergeForm6Fields,
  stripForm6BasisLabel,
  type Form6ReportFields,
  type Form6Signatory,
} from "./form6Report";
import { capitalizeReportPosition } from "./reportPosition";
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
import {
  createLostMilitaryIdFields,
  mergeLostMilitaryIdFields,
  lostMilitaryIdWorkflowSteps,
  buildLostMilitaryIdReportText,
  buildLostMilitaryIdOrderText,
  investigatorFromPersonnelRow,
  reportSignerOf,
  reporterFooterBlock,
  type LostMilitaryIdFields,
  type LostMilitaryIdSignatory,
} from "./lostMilitaryIdReport";
import { PersonnelNamePicker } from "./PersonnelNamePicker";
import {
  createLostMilitaryIdReportWordBlob,
  createLostMilitaryIdOrderWordBlob,
  createLostMilitaryIdActWordBlob,
  createLostMilitaryIdKitZip,
} from "./lostMilitaryIdWordExport";

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
  | "temporaryMilitaryId"
  | "lostMilitaryId";

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
  /** Після першого кліку по кроках — лише явно відмічені, без авто-попередніх. */
  independentSteps?: boolean;
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
  /** БР ще не підходить (немає точної дати / чекаємо список) — рапорт неповний. */
  basisNotReady: boolean;
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
    key: "sentReport",
    title: "Відправили Рапорт",
  },
  {
    key: "sentScans",
    title: "Відправили Скани",
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

const loadLostMilitaryIdSignatoryRecords = async () => {
  const own = await api.listDocumentSignatories("lostMilitaryId");
  if (own.length) return own;
  return api.listDocumentSignatories("ubdReport");
};

const toLostMilitaryIdSignatories = (
  records: DocumentSignatorySnapshot[],
): LostMilitaryIdSignatory[] =>
  records.map((record) => ({
    blockType: record.blockType === "APPROVAL" ? "APPROVAL" : "SIGNER",
    title: record.title,
    rank: record.rank,
    fullName: record.fullName,
    signatureData: record.signatureData ?? null,
  }));

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
  "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCADOARQDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigBlPoooAKKZT6ACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoopj/cpPcDzT4+fFvTvgr8L9Z8XampuWslC2dlGcPe3TfLDbp/tO//AHz97+GvP/2UfE3xR16DxhbfFbULS58R2d5azfYbG3SKLTlnt0n+y/IvzbBKv39zf7b/AHqxntpf2j/2mHu5ys/w3+GF15VvE6Ntv/EH3Xl/uMtr9z++ku/+F69A/ZzM+sWXjvxNLOlzFr3i7UZbaZf+feB0s4v/AElq38IHs9Mp9FZrYAoooqgCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigBF6UtFFABXjX7UPxdl+EHwg1bVdMbzvEt/s0rQrdV3vLfz/JFtT+PZ80rL/cieuz+Ifj3SPhl4O1jxPr12tpo+l2zXdzL/sL/CvPzO33VX+Jq+D/ANlpNb/av/as1X4qeLE2aV4YVX07Sd3m29m8vy28Sv8Ac3Im6Vtv3pfm2p8lXHuJn1V4L8O6f+y7+zOkKMd2g6Q91eTeZ5r3d7s3O25tvmNLL8i/3tyV6F8JfC83gT4Z+FfD0wV7nTdMt7eeRP45ViXe3/Am3VyHxwhfxPqHgjwLG8SRa5rcd3qCMu/dYWe26lT/AIHKlvF/uy17En3KUtdQQ+iiipGFFFFABTKfRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUVkatq9noGm3Wo6hdw2FhaxNPPc3EixwxRqu5ndm+6tef/DT9oz4b/GbV9V0rwZ4t07XtQsFZp7e3dt+xX2703L+9Td/Gm5fmX+9QB6vRRRQAUUUUAFI33TS1w/xe+I2lfCT4e+IPGOsOfsOkWrXBiQ/PO/3UiT/ad2VF/wBp6APl/wDbL+IOleJ5dd8HanFDdeCvBNjF4k8YO0jL58u7dZaau10+aVh/wHzbdv71ep/sU/D68+HnwC0eXVQj674gZtf1B1+55tx8y7P7q7Nnyfw/NXzHdeA7zXj8M/hT4gZrnxj8SNYfx545Yvsd7OJt6W7/AHW2b2/dL/B9n219/eJtdsfA/hDVdavP3OmaRZy3UqxL92KJN3y/8BWtHtYhbnB+G3fxh8fPFWrth7LwrZxaBZsW/wCXqdUurv8A8c+xL/wFq9eXpXlv7P8A4bvPD/w2sLnV4418Qay8utart/5+rp2ldf8AgG/yv92Ja9TqHpoWFFFFIAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKyNV1Wz0Swu72/uYbO1tYmmnuJm2JEq/M7M38K1qO+K+S/jdd6l+0r8V7f4OeHbyeDwbo063fjrUbdfvfxRaer/7X3nX/AGk+8iSpQBnaToN5+3L4gGv659ps/gdp11s0zRGieJ/EUsb/APH1cbv+WW7dtT+H/rr/AKr0HxN4X0rwr+0h8FdP8P6ZaaVawaVr0X2exgWKFIFS1+TYv+0yV7lpGlWmg6ZaafYwrbWVrElvBCv3YkVdqrXi/hnf8Qf2nPEmuf8AMH8D6cnh6zZX+WW8udk92/y/3FW3i2/3t9UgPekp9M+Wn1IBRXF/E/4g6N8KfAmt+Lteufs2j6TA11O4Pzt/dRP7zu21EX+JmVa5f9nvx94q+KXwwsPFPizRbfw9d6jLLcW2n2crS+Xa7j5W9v422/xrt3f3F+7QB65Xyl8bZm+Nn7R3gz4WRot54d8PRf8ACTeKVVv3St/y5QS/3v8Alq+z/aievovxp4u0zwH4T1jxFq84ttL0mzlvbqXbu2xRLuevhjUvFOtfCL9lHxf8RNQV7D4lfFfUXSxt5ldJrV7p9lvF/wBu9vvdW+X5YkX+5TQHpv7KwPxo+M3xM+NdyJJdPmuf+EZ8MzM6sn2C2ba8qf8AXV90v/bV1/gr1f49BfFb+FPh0JNj+KtR/wBMCsyv/Z1t+/uP++tsUX/bWt74D/DS1+Dvwj8KeELZsppVjFC7/wB5/wCP/wAernPh9df8J98Z/GnipR52m6IqeFdMmZfvOr+beuv/AG18qJv+vWhgeyQpsWpqTtS0gCiiigAooooAKKKKACiiigAoplPoAKKKKACiiigAooooAKKKKACiiigAplPqOV9qNQB43+038Ypfg38MrnUdGiW88V6pKmmaBZMjP5t7IrbGdV+ZkRUeV/8AZipf2a/hA/wW+HNppd7OdR8TXjyX+tapN80t3dSu8r/P/vO3/jzbV315T8LoYP2o/wBovWPiVM0WpeBfA0r6H4Xi+WW2uLr5HuL1P73zpFsf+7FE6/f+X6O8b+O9B+HmjJqXiLVYdJsnmjtUmuG/1srvtSJF+8zt/dWgBPiX40tPhr4H13xNfLvttNtmm8rdtMr/AMESf7TttRf9p65f9nbwVfeA/hfpVrq7l/EeotLrGsPs2br25fzZfl3Pt2s+35fl+SuS+KYf4v8Axn8LfDyKKWbQdBaLxP4kbdsifaz/ANn2rfL82+VGlZP7sK/36qftm+PNW0P4f2HgbwpdGHxx48u10LT5Yptk1pE3/HxdL86t8ifKrbvlllioAi+BPxT179oT4j+JfF2l6q9t8LNJaXR9KsoViZdUnV/nvWfbu2/3V+7s8p/4mr6S+4lcj8Mfh9pXwr8D6J4V0aBLbT9Ltlt40H/jzf8AAmrV8Q65YeH9H1DVNRuYrTT7CCW6ubiZ9qRRou92Zv4V2ik9wPkn9rq+uPjX8Y/h/wDAPT336fdSxa/4kTfndZRP8kT/AOz8rP8A3t/2dq+v7DT4NPs4LW3hS2giiVI4Yk2pGq/wrXxZ+wLp998VfGPxF+NuuWz/AGjxDqLRaY9x/DEvy/J/uKiQf7sX+29fcDusKfM1U1fQTPlr9rS5k+KninwR8ErGfYviO6XVdd2jDpYQOrIn+80qbv8Act5VriPiDNpPxk/bQ8BfDewtpLnQvh1a/wBoX1vFCrWiSuibEcP/AHE+z7dm7/W/7D1ofBHx/pviHVvjb+0hqTtcaJbmXSNFVP8AoG2C/wDLLft3+a/71f8Aaldayv8AgnRoWqeKbzx/8UfEMUJ1bXtTlgW4gT7/APy1l2/O/wC63v8AL/vf7ta25Y2EfVfxa8ef8K0+G+u+IY7f7ZqFrb7LGy/5+rx/kt7f5f78rIv/AAKnfCDwCvwu+HOieHBM17PZwYurt/vTzs2+WX/gbs7Vx/jN5/G/xx8KeF4Gb+yvDS/8JJqro23dK3mwWUX+1832iVv+uSV7On3Kzew0PoooqRhRRRQAUUUUAFFFJuoAWimM6/3qhmuIraJpJXVET5nZ2+7QBZpMV55q3x6+GWkXTWeo/EbwpZ3X/PG41u1R/wDvlnqpD+0H4Nv0mfSrzUfEKxNsaXQdFvdSi3f78ETr/wCPUWuK9j07FLXnCfFa9vLhP7M8A+Lb+3b/AJeHtbeyT/vm6uIn/wDHat2vibxtfxv5Pgy2sSv3f7V1pU3f9+IpaOUL3O530Vydz/wmd5bxeU2iaPL/AMtN6y6gn/tvVuz0/Wp4dt9rK7v7+n2awr/4+0tFrAzoaKqQ2U0USqbiSUj+KTbuP/jtFBJeooooLCiiigAr5k/bB+I+r2mlaT8NPCMsv/CZ+OZfsKSwMyPYWG5EuLoOqPsb96kSv/D5u/8A5ZNX0fd3MVlbyzzyrDFEu5nf7qrXyT+yZaTfG74neNfj/rNu81pqk/8AZHhD7RE6eRpcG9N6K3/PXczbv70sq0AfSPww+HWl/CvwNo/hXRI/J0/S7ZbeL/a/2q+Uf2j/ABlpuu/tUaBDqsv2nw18LdJ/4Se+sYpf3t1qM7+VaRRJ/FLu+zqq/wDT1/tfN9uP8iV+cf7Nvw48Q/HH9rj4u+Otc82LwVpHiuVorGaB9l/dWfm2to/zfeWKJf733v4fu7GgPrz9nbwLfeE/CdzrPiVU/wCE58U3Da1rjom3ypX+5brv+by7dNsS7v7teXfDS4f43/tfeNvGTR7vD/gOL/hGNJdt/wA11uf7XKm5dv8ArfNib/r3ir179or4uQ/BH4Q+IvFky+dd2sHlafbMcfar2R/Kt4v+BSun+6u5v4ayf2TPhlH8Kvgl4d0wyede3kC6lfS7NvmzyqrO/wDvfd3f3m3N/FQwPaduxa+S/wDgo94+l8I/s+XPh6xec6r4tuotIiS3++0X35f+AMq+U3/XxX1o/wByvz2+Lkv/AA0T/wAFBtC8HRO1zoXg22ieeLzdiJKrxTyun+1ultf/AAFdaqO4H1f+zD8O4/hZ8C/Bnh7bH9ot7FZbp0/5ayv87v8A8CZt1Yn7X3xCv/AfwS1ODQ2ceKPEcsXh3RfK3b1urn5PNTav3oovNl/7ZV7kibFr5U+Kmvad8RP2tPC+g3lzCnh34ZWMvinWHuJ9kSXkq7LR/wDeiTe//bwv92hK8mJnmX7Qxg+F/wADvDnwpsxbC38P6LFrWtI7/upWR9lpat/FslvPl3/w+UlfS37NvgmD4M/s++GtMvrqNBZ6d9tvLp2bZvfdK7tu/wB7/wAdr5W8ZafefGjxB8N7HUorvz/il4obxFc2kTuvkeH7CJEt7d1f7rfvYpX2/Lv3uv8ABX1D+0Peza9H4Y+HNpu+0eL7zydReIcxaXF8963/AANdkG7+H7Ruq5q1gRb/AGd7STVfDmpePNQtJrLWfGt4+qypcrslitf9VZQ/8Bt1i/4E7/3q9lrOkubTSrLfPJHaW8SfeZtiqteV+LP2sPhJ4NmMF946067u92z7JpLtf3G7+7st1dqyYz2WivBZP2k9Y12SaPwj8IPHuvbG2pcX9hFpFvL/ALe68libb/wCn3N5+0N4lureSx0rwB4G09k/exanPda1d7v92L7PEv8A329ID3eqk1zFbI7SSqiL95navDl+CPxD8Tpdp4x+NOtyQSsrJZ+DtOg0VIv9jzf3s/8A5FWrtj+yH8M47xL7WdCufGmoIu37R4u1GfV//HLh3Rf+ArQJm34i/aT+FfhVpbXUPiD4eS7i+9aW99FcXH/fqLc//jtY6/tN6bqskSeHPBXj/wASeb/qprfw3Pa27f8AbW8+zp/49Xo/hjwT4f8ABFkLbw9oOm6Faf8APvplmlun/fCLXRfLQCPHYPHfxb1ufZp/wsstIt9v/Hx4k8TRRPu2/wDPK1iuP4v9tafbaL8a9Vs9upeLvCHh+V/vLo+gXF7LF/uSy3CL/wB9RV7FRQM8Yf4Ea1rE1q3iD4t+OtVSJcS29jc2ulxT/wDgLbxSr/wGWpR+y38Nbm4hn1jQ7vxS8X3f+Ep1a91hF/4DdSyrXsLdKNtAHI+E/hZ4O8Bo/wDwjPhPRPDzy/f/ALM06K13f98LXWeUn92n0m2gA20baWigBNtLRRQAUUUUAFFFFABTKfVO5nS2gaVnVEVdzO/8K0CZ8w/tm+LL7xNaeHvgf4Xl/wCKi8fNLDeNu/49dLT/AI+Gf+L97uWL/ceVv4K+ifBvhmx8C+FNK8P6bF5On6dbJbwJ/sqtfMf7LNvJ8aPif41+O9755tdSnbRfDMTBlRNLif5Zdjr/AB/63f8A9PEq/wB2vrqgEO201EVPu0b6yPEniGx8KeH9S1fUp1s9MsIJbq5uG/giRNztSe4z46/an125+MX7V/wl+DenKj6bpM6eKtY2Spu3pvSJf9lok82XZ/FuSvtmFFRFVF2KtfBP/BP621X4l/GH4p/F/XFkS+1ef7OiPtdEiZ90USN/eiVWi/3dlfezusKbmqmBjeL/ABHY+CfCuseIdSmW20/S7SW+upm/hiiTc7f98rXw/wD8EzPCd3rsnxA+K+ux41jxHqcqO/y7PN82WW42f3dsssv/AAHZWl/wUO/am8I6T8HNW8BeGvENtr3i3XbqDTJLHR50uHih83fKku35V3ojRbPv/vVq98A/DX7QWlfB7wz4U8O+EPDfwosbK32XWseJLltSvrh2+9KtrFtVXb77ea9CA+0L65is7WWWSZYYo13O7t8qrX5dQ/Hzwr4h0bxTcz61fXusfGDxT5Vzp2mWrXF9a+H4E2IvlKj/ADyxJFE6p/f+X7tet/ti/A2x8N/BDWtX8fePfE/jzxLqc9vp9jDc3/2OyiuJW+d4LOLbF8iJLLtff/qq7n9jv4Xad4ZfxJ4tfSLLRNPtYItD0x7aBU/dWqt9quN+35ke4adlb+5t/wBmtIuwmeS+Dvib44+MP7VOteJ/h98MZ4IfDmiw+HrNfF8raamlrlndpUXe2593+qT+GJN2zdXoHwv+FfxR+O3ifxF4+8VfE+Xwynm3Xhuzh8EWyRK1rBNsle3luPNZN0qy/Nt3NsRt23ZWB8FtU1c/APXfEOkS+T47+M3iS6/se4hf5oIpXlb7R/s/Z4vtDf8AbJFr7V8G+ErHwN4S0fw9pkfk6fpdnFZQJ/sIu1aU3cEeUWX7HHwzkmgn8R2Gq+PL2Lpd+LtYutS3t/e8qV/K/wC+Ur1vw54N0TwXpsVh4e0XT9FsU+7badapbxL/AMBVa6CisxjNlPoooAZT6KKAGU7bS0UAFFFFABRRRQAUUUjdKAFooooAKKKKACiiigAooplAD6+XP29vGtzpnwfi8BaDIf8AhK/H19FoFjbwyqkrxO266+//AAtEGi3f3rhK+oN/y18a/D6WX9oT9uLxb4qnujL4X+FcbaFo8Ufyq1/Ku26lf+98yOn/AG7xf8CAPpv4W+B7D4beAdE8MaZHElpptssP7mPytz/edtn8O52dv+BV2VMT5a4H4m/GjwT8HtGOq+MPE1loVuFd40uZf30+3/nlEvzy/wDAFagDvq+Uv23/AIj6NpXhq08J6rqNrb2Nwj61rVq86rNcaXavue3T51bzbiXZEuz/AG1/iqt/w0N8X/jfNt+D/wAPl8N+FW+ZvG/xA3WsUqbG+a3tV+d1+5tl+df76V8r+M/gsvxOtPA/iPxZ4vu/Hnjf4h+MotI0rVtRTyol0S13M1xFZrsRImlRfkf5dtxv27notfUD1f8AZC+M3ipPglp+kfDPwHqXj7xVqN02qa/4h1k/2Vo1rez/ADO/msm6d/lWVkiX/lr9+vbo/wBmTxf8W3W5+OHj6712xlj/AHngvwm8ulaEv+w7K3n3S/8AXV1/3a+h9K02z0LTrexsoI7axto1hgt4l2pEi/KqqtX5n2RUXvoB8Aaz8KPCPiH9uXwN8OvBvhrS9H8K/DrSv7f1KHTYFiT7bLKjIku37zbUtX3Nub97X6BJXx9+wrZDxx4k+LXxmnWR/wDhMNfli0yWaL52sIv+Pd0b+40X2ddv96Jq+vn+RaAPhP8A4KBXGpePvin8FPhboTvJqF/qEuqNEu1V3JsRGd9v3fKe9+7Xsf7VWsR/BT9kvxLY6Jts5WsYtAsHRvK8prp0tfN+X+55rS/8Arzz4P6zpvx7/bo8f+NbOKO50/wJp/8AwjFjdpv/AHrMy73/ALvyy/bU/wBpdn9yqv8AwUZi1Lx7L8J/hXoqyPqHijXZbjei71hiiTyHlf8A2U+27/8AtlWnVCZ0/wCyb4Xg8QXdn4htIFg0Dwho8XhPRYrbclvLcfJLqVwq/wC1P+6/7d/vfO1fWtcj8PvBNh8O/COk+HNMULZabAsCM4+d/wC87f7TN81ddUPcEFFFFIYUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUm2oZZVgjd2bYifMzNQB5v8AtE/FD/hTPwS8ZeLo3i+26Xp0r2Mcq7klum+S3i/4HK8S/wDAq8R/Z81fwd+yP+y74YvvHuvQ6bf6+raveSssst7qN1Ptff5XzSyy+V5W/wCX/erxP9tX40/8NW/EHwR8CfhBqkeo6n/a39paprEUrfYovKifYu9P9ai7pXbb/FEm3c/3Pp74Vfsh+HfCHiaXxr4tvLj4ifEi6fzZ/EmvJv8AI+feqWsH3IEX+Hb93/x2gDjbD4t/HH9pm3Wf4aeHYPhX4HvI82/i/wAY23n6lOrf8tbWyV9q/wCy8u5GVq7r4Tfsd+Bfhvqn/CSarFeeP/HUzpLP4p8Vn7bdh1/55bvliVP4dnzf7TV7+kQSld9i0AeIftK+ILmbw3pPw70a5a28S+O7r+yIGi277ez+9e3G3/Yt9/8AwJ0rzLS9L0vxT+3BZaJYSKdK+GHhWG3gs9rNsnuPmdf+/T2D7v8A7Kux+C1y3xZ+K3jL4pzKsuiWp/4RrwpMjb0ls4n33V2vzbGW4l+46/8ALKJK579ifTbvXdd+L/xH1GWR5vFXiiX7F53zN9iiX/RG/wC/DxJ/2yrToB9WV86fty/Fe8+Fn7POv/2U0n/CQ+I9vhvRfKk8l/td0rojo38LKu9/+Afw19Fsdi7q/O79tSW++Mv7Rfw48LW08sPhTRvEVroN5NFu+e/vImn3p/D+6S1i+b+FpXrMD7A/Zx8CQfDP4E+B/D1tE1ulrpkLNC3zbHdfNdf++naoP2hvHF74X8HW+l+HnVfGHiO6TRtE+b/VXEv37j/tlFvl/wBrZt/ir1VPkir53+FM0fxw+Lmr/E/zxe+GdBSXQPC8334pTv8A9KvYv9918pX+bcqf99AHif7BdrovwOuvj/o+savbW1j4V1tIp727ZYv3ESSr5r7v4dtW/wBmHxreftS/tYeMPiZIGh8N+F7FNK8PW5DKywyvL+9dG+68u12b+L5EV9uxd/tPxu/Zx+Dmt3GtfEjx34SgvJdOtnvNQm+03EUU8UCb99xEjqk+1Il++rfKiJ91a53/AIJ7eCJNC+AaeIb7TY9N1jxVqNzq9zbJF5Xlbn2+Vt/hVWV9qfw7qaA+qKKKKlbgFFFFMAooooAKKKKACiiigAooooAKKKKACiik20ALRRRQAUUUUAFFFFAGF4q8T6V4J8PX+ua5qFtpOkWETT3V7dPsiiT+89fD/hvx54r/AOCjWr6raaZcXngr4B6defZL57dvL1XxA+xH+zu3/LKJ0Zdyr/C/zbt/yfcXiHw9pninR7vStZ0y11jTbyPyrmxvoFlinX+4yN8rV84x/sCeCvCuv3GpeAfE/jT4ZLPL9on03wnrbW9jLLv3b3iZG/3dm7Zt/hpoTOB/ZY8CeGJf2y/i1feFtP03TfDHgbT7PwrpllY2aRKr/O1xLu+95qTpexM/8W+vuSvzX/YP+H3xY8Y+APEvjjwr8Xo9L/tbX7qW5tr3w/Fdfb3ZEl815Xfcnzzy/Kq19Q6jbftQ2NpFFpl78LNUZU+e61SLUYHZv9yKhgj6Hrwr9pvx5q+laFpfgXwjKqeOvHc76Rpku757OLYzXV7/ALtvF83+88X96sseKf2lfD1ok2q+E/h5r4Vv3qaTrF1ZbE/vbp0avF/gp8Qfix8SPHF98bb/AODF5rdprGnLp/hm3s9ftYVsLBW3yuqzum5riXY2/am5UT+GkUj6M+IiWf7Pn7MPiX/hGkitIvCPhe6OmK/9+C3fyt395mZV/wB5qr/sb+DYPAv7Nvw/02FvNR9KiuhMyqrMJfnXO3/ZZa+dv24/2g/ELfs/ajovib4R+I9BsdX1CwtZrq4ubWWLal1FcPF+6d/vxRSpXr/wq/aY03TPD2leHtS+GPxN8MppdjBZ/aNR8JXH2dvLiVfkeLf/AHaaBnr3xa+INj8KPh14j8VXyq8Wl2j3CwM2w3Ev3YoV/wBqWVkRf9p6+QvHnw6n+GXhz9mvUNVlD69qPxFt9V1+7dtvm3t1FcXFw/zfdXcrfJ/DVr4z/tO/DX4rfFn4eeDtR1yfSvCWnammsawuraZOiXtwuz7Fb7WT7u+Xe+77u2KvWvjl8T/gJ8R/hh4h0Pxv4r0258Ox3VvBP9kum+0Lcb1aJ7fyvmdt38UW5fv7vl30MRL+0T8QrrxRrWm/BXwbqxi8YeJUZ9Wu7Qb30TSdn764f/nk77kii3fxS7v4K9x8G+EdI8BeF9N8P6HaRWGkadAtvb28S8Iq15D8Lfh98G/2b9IvbXQdR0vRp9U2Xd/qGp6t5t9qDfwPLLK29v8Ad+78zfL81c54y/bG02+1h/Dvwd8PX/xc8X8q6aSdmmWB/v3V4/7pP9n+9t27lpAZf7eXimXW/Cvhf4N6JPv8S/EbVoNP8pGbfFYROst1K+3+H5FR/wDZd6+nPDmhWvhjQdN0izDJa2FtFaxbv7iLtWvB/wBnz9nnXvD/AIu1T4o/FLV7TxT8Utah+y77RG+w6Na/8+tlv+bb/ef+L/vpn+kKACiiigAooooAKKKKACiiigAooooAKKZT6ACiiigAooooAKKKKACiiigAooooAKyPEtytj4f1O5f7kVtK7f8AfNa9VruFbm1lidco67Wo6gfLn/BMvTVsP2NfAcnlRpLefariXYv3/wDSHRG/74RK+q91fFf/AATc1S58OeCfHXwm1q++0+IPAfiKew2fP/x6/ciZd38LvBcP/uun96vbv2gv2gtF/Z/8Ki5vi2qeJdSY2uh+H7RGe71a6/giiVdzbdzrub+HcP7y0PcDlv2jtXuviLr+hfBLQbny7rxG32rxNLCzI9noKv8A6R8yfce4b/R0/vbpf7u6veNE0Ow8P6NZaVp9qlpp9hAlra28X3IokXair/wGvJP2bPhTrvgzw7qfirxvLFd/E3xfLFqevzKE2Wz7NsVlFt/5ZW6fIvzNube+7569zoA+Pv8AgpXbfbfhf8PdPbc0V/43sIpdn9xbe6l/9CRK+nfGPizT/BvhfU9d1WQw6fp1s91O6JufaiFvlX+JuPu185/8FGNB17VP2fodV8P6bNqt14c1q11eWCJXZ0gVZYpZfk+b5Vl3N/dXc38Nc43jDV/24fFuhWGjaPrWgfBPSJYtS1rU9SgeyfxDOp3w2VujffgV1/et91tv8O1d4B6h+yhoV9qfgHUPiJ4nto/+El+IN5/b8qbN7W9m6qtlb7/4lig2f99tXZax+zV8JNauZbnUvhf4Ov555PNlmudAtXdn/vM2yvSoYUhTZGqqi/dVanoA8l/4ZV+DId3f4TeCXdvvbtAtW/8AaVei6ZpdpodlFZ2VrBZ2cS7Yre3iWNEX+6qrWrRQAi9KWiigAooooAKKKKACiiigAoopO9AC0UUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFNZdwp1FAHzf8WP2L/CXxS+IkPxBstf8UfD/wAapB9lutb8HagtlLexBUVPP3o27Yqbf935W3bE22/g9+yD4P8AhR4nl8WT3mt+N/HEq+V/wlfi6/N/qEUW3b5UT/diX5nX5F3bX+9X0LSN0oANtLRRQAyjZT6ZQA+iiigAoopN1AC0UUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQB//9k=";

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
const FALLBACK_UBD_BASIS = pickUbdBasisOrderForTaskPeriod("") ?? {
  number: "4862/ОКП/1162/дск",
  date: "09.05.2026",
};
const DEFAULT_UBD_BASIS_NUMBER = FALLBACK_UBD_BASIS.number;
const DEFAULT_UBD_BASIS_DATE = FALLBACK_UBD_BASIS.date;

const stripUbdBasisNumber = (value: string) =>
  String(value ?? "")
    .trim()
    .replace(/^№\s*/, "");

const formatUbdBasisText = (number: string, date: string) => {
  const orderNumber = stripUbdBasisNumber(number);
  const orderDate = String(date ?? "").trim();
  const reference = [orderNumber ? `№${orderNumber}` : "", orderDate ? `від ${orderDate}` : ""]
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
    basisNumber: match[1],
    basisDate: match[2].replaceAll("/", ".").replaceAll("-", "."),
  };
};

const createUbdFields = (
  row: EjournalPreviewRow | null,
  summary: ReturnType<typeof buildPersonSummary>,
  signatories: DocumentSignatorySnapshot[] = legacyUbdSignatories(),
): UbdReportFields => {
  const fullName = summary.name !== "Особа не вибрана" ? summary.name : "";
  const staffPosition =
    getPersonFullPositionTitle(row) ||
    getPersonFieldValue(row, ["посада"]) ||
    getPersonFieldValue(row, ["штатна", "посада"]) ||
    getPersonFieldValue(row, ["чим", "займається"]) ||
    getPersonFieldValue(row, ["column_5"]) ||
    "";

  const taskPeriod = buildFighterTaskPeriodText(row);
  const taskPlace = getFighterTaskPlace(row);
  const resolved = resolveUbdBasisForTask(taskPeriod, taskPlace);
  const basis = resolved ?? FALLBACK_UBD_BASIS;

  return {
    commander: "Командиру військової частини A4862",
    fullName,
    rank: summary.rank || "",
    staffPosition: capitalizeReportPosition(staffPosition),
    birthDate: summary.birthDate || "",
    rnokpp: summary.rnokpp || "",
    taskPeriod,
    taskPlace,
    basisNumber: basis.number,
    basisDate: basis.date,
    basis: formatUbdBasisText(basis.number, basis.date),
    basisNotReady: !ubdBasisDateMatchesTaskPeriod(
      taskPeriod,
      basis.date,
      taskPlace,
    ),
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
  const savedBasisNumber =
    stripUbdBasisNumber(String(saved.basisNumber ?? "")) || parsed.basisNumber;
  const savedBasisDate =
    String(saved.basisDate ?? "").trim() || parsed.basisDate;
  const merged = { ...defaults, ...saved };
  const pick = (personnel: string, document: string) =>
    String(personnel ?? "").trim() || String(document ?? "").trim();
  const taskPeriod = pick(defaults.taskPeriod, merged.taskPeriod);
  const taskPlace = pick(defaults.taskPlace, merged.taskPlace);
  const resolved = resolveUbdBasisForTask(taskPeriod, taskPlace);
  const picked = resolved
    ? { number: resolved.number, date: resolved.date }
    : pickUbdBasisOrderForTaskPeriod(taskPeriod, undefined, taskPlace);

  // № БР: локація + період, якщо є в довіднику; інакше дата «з».
  // Збережений вручну номер лишаємо, якщо він серед відповідних.
  let basisNumber = savedBasisNumber;
  let basisDate = savedBasisDate;
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
    const savedIsKnownForDate = allBasisOrderOptions().some(
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

  return {
    ...merged,
    fullName: pick(defaults.fullName, merged.fullName),
    rank: pick(defaults.rank, merged.rank),
    staffPosition: capitalizeReportPosition(
      pick(defaults.staffPosition, merged.staffPosition),
    ),
    birthDate: pick(defaults.birthDate, merged.birthDate),
    rnokpp: pick(defaults.rnokpp, merged.rnokpp),
    basisNumber,
    basisDate,
    basis: formatUbdBasisText(basisNumber, basisDate),
    // Не виводимо з розбіжності дат для старих записів — лише збережений прапор.
    basisNotReady:
      saved.basisNotReady === true ||
      (saved.basisNotReady as unknown) === "true",
    // Always prefer fresh personnel/roster values when document field is empty.
    taskPeriod,
    taskPlace,
    signatories:
      Array.isArray(saved.signatories) && saved.signatories.length
        ? saved.signatories
        : defaults.signatories,
  };
};

const findRosterRowByPersonName = (
  rows: EjournalPreviewRow[],
  fullName: string,
) => {
  const wanted = normalizeRosterMatchText(fullName);
  if (!wanted) return null;
  const wantedSurname = wanted.split(" ")[0] || "";

  return (
    rows.find((row) => {
      const rowName = normalizeRosterMatchText(
        String(row["піб"] ?? row["ПІБ"] ?? row["прізвище"] ?? ""),
      );
      return rowName && rowName === wanted;
    }) ||
    rows.find((row) => {
      const rowName = normalizeRosterMatchText(getPersonDisplayName(row));
      return rowName && rowName === wanted;
    }) ||
    rows.find((row) => {
      const rowName = normalizeRosterMatchText(
        String(row["піб"] ?? row["ПІБ"] ?? ""),
      );
      return (
        wantedSurname.length >= 3 &&
        rowName.startsWith(wantedSurname) &&
        wanted.split(" ").every((part) => !part || rowName.includes(part))
      );
    }) ||
    null
  );
};

/** Copy fighter-status / position keys from morning roster onto the open person. */
const applyRosterFieldsToPerson = (
  person: EjournalPreviewRow,
  rosterRow: EjournalPreviewRow,
): EjournalPreviewRow => {
  const next: EjournalPreviewRow = { ...person };
  for (const [key, value] of Object.entries(rosterRow)) {
    if (key.startsWith("__")) continue;
    const text = String(value ?? "").trim();
    if (!text) continue;
    next[key] = value;
    if (
      key.startsWith("fighter_status_") ||
      key.includes("посада") ||
      key.includes("піб") ||
      key.startsWith("column_")
    ) {
      next[`roster__${key.replace(/^roster__/, "")}`] = value;
    }
  }
  return next;
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
  const fromOverview = overviewById[document.personExternalId];
  // personStatus/serviceStatus у документі — snapshot на момент створення.
  // Він не є поточним кадровим статусом і не може блокувати роботу як СЗЧ.
  return compactText(fromOverview?.label || "—");
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
  { value: "lostMilitaryId", label: "Втрата військового квитка" },
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
  { value: "lostMilitaryId", label: "Втрата військового квитка" },
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
            : type === "lostMilitaryId"
              ? "Втрата військового квитка"
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
    lostMilitaryIdFields: LostMilitaryIdFields;
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
  if (type === "lostMilitaryId") {
    const workflow = { ...createEmptyWorkflow(), currentStatus: "document" };
    return {
      type,
      title: `Втрата військового квитка · ${now}`,
      status: "document",
      fields: { ...input.lostMilitaryIdFields, personStatus: input.personStatus },
      workflow: workflow as unknown as Record<string, unknown>,
      files: { ubdScans: [] },
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
    ...lostMilitaryIdWorkflowSteps,
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
          : type === "lostMilitaryId"
            ? lostMilitaryIdWorkflowSteps
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
  if (
    (type === "ubdReport" || type === "form6Report") &&
    status === "sent"
  ) {
    return "sentReport";
  }
  return status ?? "";
};

const documentWorkflowStatusLabel = (document: BackendPersonDocument) =>
  documentWorkflowSteps(document.type).find(
    (step) =>
      step.key ===
      resolveDocumentWorkflowStatus(document.type, document.status),
  )?.title || workflowStatusLabel(document.status);

/** Нормалізує карту кроків. Старий УБД `sent` → sentReport; у характеристиці `sent` — окремий крок. */
const normalizeWorkflowCompletedRecord = (
  completed: Record<string, boolean> | undefined,
  steps: Array<{ key: string }>,
) => {
  const raw = { ...(completed || {}) };
  const stepKeys = new Set(steps.map((step) => step.key));
  if (
    raw.sent === true &&
    stepKeys.has("sentReport") &&
    raw.sentReport == null
  ) {
    raw.sentReport = true;
  }
  if (stepKeys.has("sentReport") && !stepKeys.has("sent")) {
    delete raw.sent;
  }
  if (
    stepKeys.has("sent") &&
    !stepKeys.has("sentReport") &&
    raw.sentReport === true &&
    raw.sent !== false
  ) {
    raw.sent = true;
  }
  return Object.fromEntries(
    steps.map((step) => [step.key, raw[step.key] === true]),
  ) as Record<string, boolean>;
};

/**
 * Карта виконаних кроків.
 * - якщо є explicit completed / independentSteps — лише відмічені кроки (не послідовно);
 * - інакше (старі записи) — послідовно до currentStatus лише для відображення %.
 */
const resolveWorkflowCompletedMap = (
  workflow: SalaryWorkflowState,
  steps: Array<{ key: string }>,
  type?: string | null,
  statusOverride?: string | null,
) => {
  const completed = workflow.completed || {};
  const hasExplicitMap =
    (workflow as { independentSteps?: boolean }).independentSteps === true ||
    steps.some((step) => Object.prototype.hasOwnProperty.call(completed, step.key)) ||
    Object.prototype.hasOwnProperty.call(completed, "sent");

  if (hasExplicitMap) {
    return normalizeWorkflowCompletedRecord(completed, steps);
  }

  const resolvedStatus = resolveDocumentWorkflowStatus(
    type,
    statusOverride ?? workflow.currentStatus,
  );
  const index = steps.findIndex((step) => step.key === resolvedStatus);
  return Object.fromEntries(
    steps.map((step, stepIndex) => [
      step.key,
      index >= 0 && stepIndex <= index,
    ]),
  ) as Record<string, boolean>;
};

/** Найвищий відмічений крок — для фільтра журналу (не «поточний» жовтий). */
const documentHighestWorkflowStatusLabel = (
  document: BackendPersonDocument,
  workflowOverride?: SalaryWorkflowState,
) => {
  const steps = documentWorkflowSteps(document.type);
  if (!steps.length) return documentWorkflowStatusLabel(document);
  const workflow = workflowOverride ?? mergeSalaryWorkflow(document.workflow);
  const completed = resolveWorkflowCompletedMap(
    workflow,
    steps,
    document.type,
    document.status || workflow.currentStatus,
  );
  let lastTitle = "";
  for (const step of steps) {
    if (completed[step.key]) lastTitle = step.title;
  }
  return lastTitle || documentWorkflowStatusLabel(document);
};

const countWorkflowCompletedSteps = (
  workflow: SalaryWorkflowState,
  steps: Array<{ key: string }>,
  type?: string | null,
  statusOverride?: string | null,
) => {
  const map = resolveWorkflowCompletedMap(
    workflow,
    steps,
    type,
    statusOverride,
  );
  return steps.filter((step) => map[step.key]).length;
};

const buildWorkflowFromCompletedMap = (
  existing: SalaryWorkflowState,
  steps: Array<{ key: string }>,
  completedInput: Record<string, boolean>,
): SalaryWorkflowState => {
  const completed = normalizeWorkflowCompletedRecord(completedInput, steps);
  const stillDone = steps.filter((step) => completed[step.key]);
  const currentStatus =
    stillDone[stillDone.length - 1]?.key ??
    steps[0]?.key ??
    existing.currentStatus;
  return {
    ...existing,
    completed,
    independentSteps: true,
    currentStatus,
  };
};

/** Додає позначені кроки; уже виконані на документі не змінюються. */
const mergeBulkStepsIntoWorkflow = (
  existing: SalaryWorkflowState,
  steps: Array<{ key: string }>,
  bulkSteps: Record<string, boolean>,
  type?: string | null,
  statusOverride?: string | null,
): SalaryWorkflowState => {
  const currentMap = resolveWorkflowCompletedMap(
    existing,
    steps,
    type,
    statusOverride,
  );
  const merged = { ...currentMap };
  for (const step of steps) {
    if (bulkSteps[step.key] === true && !currentMap[step.key]) {
      merged[step.key] = true;
    }
  }
  return buildWorkflowFromCompletedMap(existing, steps, merged);
};

const getDocumentProgressPercent = (document: BackendPersonDocument) => {
  const workflow = mergeSalaryWorkflow(document.workflow);
  const steps = documentWorkflowSteps(document.type);
  if (!steps.length) return 0;
  const completedCount = countWorkflowCompletedSteps(
    workflow,
    steps,
    document.type,
    document.status || workflow.currentStatus,
  );
  return Math.round((completedCount / steps.length) * 100);
};

const readDocumentFieldText = (
  fields: Record<string, unknown>,
  key: string,
) => {
  const value = fields[key];
  return typeof value === "string" ? value.trim() : "";
};

const getDocumentPersonName = (document: BackendPersonDocument) => {
  const metadataName = String(document.personName ?? "").trim();
  if (metadataName) return metadataName;

  const fields = (document.fields || {}) as Record<string, unknown>;
  const fullName =
    readDocumentFieldText(fields, "fullName") ||
    readDocumentFieldText(fields, "pib") ||
    readDocumentFieldText(fields, "name") ||
    readDocumentFieldText(fields, "ПІБ") ||
    readDocumentFieldText(fields, "ФИО");
  if (fullName) return fullName;

  const assembled = [
    readDocumentFieldText(fields, "lastName"),
    readDocumentFieldText(fields, "firstName"),
    readDocumentFieldText(fields, "patronymic"),
  ]
    .filter(Boolean)
    .join(" ");
  if (assembled) return assembled;

  const syntheticName = personNameFromSyntheticDocumentId(
    document.personExternalId,
  );
  if (syntheticName) return syntheticName;

  return document.personExternalId
    ? `ID ${document.personExternalId}`
    : "Без ПІБ";
};

const readDocumentUbdTaskPeriod = (document: BackendPersonDocument) => {
  if (document.type !== "ubdReport") return "";
  const value = document.fields?.taskPeriod;
  return typeof value === "string" ? value.trim() : "";
};

const readDocumentFormPurpose = (
  document: BackendPersonDocument,
  fields?: Record<string, unknown>,
) => {
  if (document.type !== "form6Report" && document.type !== "form12Report") {
    return "";
  }
  return readDocumentFieldText(fields || (document.fields || {}), "formPurpose");
};

const readDocumentUbdTaskPeriodStartLabel = (document: BackendPersonDocument) =>
  formatUbdTaskPeriodStartDate(readDocumentUbdTaskPeriod(document));

const readDocumentUbdTaskPeriodStartTimestamp = (
  document: BackendPersonDocument,
) => {
  const date = parseUbdTaskPeriodStartDate(readDocumentUbdTaskPeriod(document));
  return date ? date.getTime() : 0;
};

type JournalSortField =
  | "createdAt"
  | "updatedAt"
  | "progress"
  | "taskPeriodStart";

type JournalReadinessFilter =
  | "ALL"
  | "INCOMPLETE"
  | "READY_TO_SEND"
  | "COMPLETE"
  | "SKIPPED_SZCH";

const ubdProgressBackupEntryFromDocument = (
  document: BackendPersonDocument,
): UbdProgressBackupEntry =>
  documentToUbdProgressBackupEntry(
    document,
    mergeSalaryWorkflow(document.workflow) as unknown as Record<string, unknown>,
  );

const workflowFromBackupEntry = (entry: UbdProgressBackupEntry): SalaryWorkflowState =>
  mergeSalaryWorkflow(entry.workflow);

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

const ubdDocumentHasEmptyInputs = (
  fields: Record<string, unknown> | null | undefined,
) => documentHasEmptyInputs("ubdReport", fields);

const ubdLiveDocument = (
  document: BackendPersonDocument,
  workflowOverride?: SalaryWorkflowState,
) =>
  workflowOverride
    ? {
        ...document,
        workflow: workflowOverride,
        status: workflowOverride.currentStatus || document.status,
      }
    : document;

/** Зелений і поза експортом «Не подавалися» — лише коли всі кроки прогресу відмічені. */
const ubdDocumentIsComplete = (
  document: BackendPersonDocument,
  workflowOverride?: SalaryWorkflowState,
) => {
  if (document.type !== "ubdReport") return false;
  return getDocumentProgressPercent(ubdLiveDocument(document, workflowOverride)) >= 100;
};

const ubdHighestCompletedStepKey = (
  document: BackendPersonDocument,
  workflowOverride?: SalaryWorkflowState,
) => {
  const live = ubdLiveDocument(document, workflowOverride);
  const workflow = workflowOverride ?? mergeSalaryWorkflow(live.workflow);
  const completed = resolveWorkflowCompletedMap(
    workflow,
    ubdWorkflowSteps,
    live.type,
    live.status || workflow.currentStatus,
  );
  let last = "";
  for (const step of ubdWorkflowSteps) {
    if (completed[step.key]) last = step.key;
  }
  return last;
};

const ubdDocumentHasBasisDateMismatch = (
  fields: Record<string, unknown> | null | undefined,
) => documentHasBasisDateMismatch("ubdReport", fields);

/** Можна відправляти рапорт: крок «Готово до відправки», без червоного/жовтого, ще не відправлений. */
const ubdDocumentIsReadyToSend = (
  document: BackendPersonDocument,
  fields: Record<string, unknown> | null | undefined,
  workflowOverride?: SalaryWorkflowState,
) => {
  if (document.type !== "ubdReport") return false;
  if (ubdDocumentHasEmptyInputs(fields)) return false;
  if (ubdDocumentHasBasisDateMismatch(fields)) return false;
  return ubdHighestCompletedStepKey(document, workflowOverride) === "ready";
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
  const [lostMilitaryIdFields, setLostMilitaryIdFields] =
    useState<LostMilitaryIdFields>(() =>
      createLostMilitaryIdFields(null, buildPersonSummary(null)),
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
  const [skippedDueToSzch, setSkippedDueToSzch] = useState(false);
  const skippedDueToSzchRef = useRef(false);
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
  const [journalStatusFilters, setJournalStatusFilters] = useState<string[]>(
    [],
  );
  const [journalReadinessFilter, setJournalReadinessFilter] =
    useState<JournalReadinessFilter>("ALL");
  const [journalStatusMenuOpen, setJournalStatusMenuOpen] = useState(false);
  const journalStatusMenuRef = useRef<HTMLDivElement | null>(null);
  const [journalMonthFilter, setJournalMonthFilter] = useState("ALL");
  const [journalNameQuery, setJournalNameQuery] = useState("");
  /** IDs зняті з експорту; усі інші з відфільтрованого списку експортуються. */
  const [journalExportDeselectedIds, setJournalExportDeselectedIds] = useState<
    Record<string, true>
  >({});
  const [journalBulkProgressSteps, setJournalBulkProgressSteps] = useState<
    Record<string, boolean>
  >({});
  const [isApplyingBulkJournalProgress, setIsApplyingBulkJournalProgress] =
    useState(false);
  const [journalBulkProgressUndo, setJournalBulkProgressUndo] = useState<
    UbdProgressBackupEntry[] | null
  >(null);
  const [ubdProgressBackups, setUbdProgressBackups] = useState<UbdProgressBackup[]>(
    () => loadUbdProgressBackups(),
  );
  const [ubdProgressBackupsOpen, setUbdProgressBackupsOpen] = useState(false);
  const refreshUbdProgressBackups = () =>
    setUbdProgressBackups(loadUbdProgressBackups());
  const showUbdExitDateColumn = journalTypeFilter === "ubdReport";
  const showFormPurposeColumn =
    journalTypeFilter === "form6Report" ||
    journalTypeFilter === "form12Report";
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
    mode === "ubdRestoreReport" ||
    mode === "lostMilitaryId";
  const [lostMilitaryIdPreviewDoc, setLostMilitaryIdPreviewDoc] = useState<
    "report" | "order" | "act"
  >("report");
  useEffect(() => {
    if (mode !== "lostMilitaryId") return;
    const step = resolveDocumentWorkflowStatus(
      "lostMilitaryId",
      workflow.currentStatus,
    );
    if (step === "order") setLostMilitaryIdPreviewDoc("order");
    else if (step === "act" || step === "received" || step === "handed") {
      setLostMilitaryIdPreviewDoc("act");
    } else {
      setLostMilitaryIdPreviewDoc("report");
    }
  }, [mode, workflow.currentStatus]);
  const lostMilitaryIdPreviewLabel =
    lostMilitaryIdPreviewDoc === "order"
      ? "Наказ"
      : lostMilitaryIdPreviewDoc === "act"
        ? "Акт"
        : "Рапорт";
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
    if (mode === "lostMilitaryId") {
      if (lostMilitaryIdPreviewDoc === "order") {
        return createLostMilitaryIdOrderWordBlob(lostMilitaryIdFields);
      }
      if (lostMilitaryIdPreviewDoc === "act") {
        return createLostMilitaryIdActWordBlob(lostMilitaryIdFields);
      }
      return createLostMilitaryIdReportWordBlob(lostMilitaryIdFields);
    }
    return Promise.reject(new Error("Немає Word-шаблону для цього документа."));
  }, [
    form12Fields,
    form6Fields,
    lostMilitaryIdFields,
    lostMilitaryIdPreviewDoc,
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
            ...lostMilitaryIdWorkflowSteps,
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

  const liveJournalFields = (
    document: BackendPersonDocument,
  ): Record<string, unknown> => {
    const stored = (document.fields || {}) as Record<string, unknown>;
    if (document.id !== selectedDocumentId) return stored;
    let live = stored;
    if (mode === document.type) {
      switch (document.type) {
        case "salaryPowerAttorney":
          live = salaryFields as unknown as Record<string, unknown>;
          break;
        case "ubdReport":
          live = ubdFields as unknown as Record<string, unknown>;
          break;
        case "form6Report":
          live = form6Fields as unknown as Record<string, unknown>;
          break;
        case "form12Report":
          live = form12Fields as unknown as Record<string, unknown>;
          break;
        case "serviceCharacteristic":
          live = serviceCharacteristicFields as unknown as Record<
            string,
            unknown
          >;
          break;
        case "zhbdCertificate":
          live = zhbdCertificateFields as unknown as Record<string, unknown>;
          break;
        case "ubdRestoreReport":
          live = ubdRestoreFields as unknown as Record<string, unknown>;
          break;
        case "temporaryMilitaryId":
          live = ticketFields as unknown as Record<string, unknown>;
          break;
        case "lostMilitaryId":
          live = lostMilitaryIdFields as unknown as Record<string, unknown>;
          break;
        default:
          break;
      }
    }
    return { ...live, skippedDueToSzch };
  };

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
      const liveWorkflow =
        document.id === selectedDocumentId
          ? mergeSalaryWorkflow({
              ...(document.workflow && typeof document.workflow === "object"
                ? document.workflow
                : {}),
              ...workflow,
            })
          : undefined;
      if (
        journalStatusFilters.length > 0 &&
        !journalStatusFilters.includes(
          documentHighestWorkflowStatusLabel(document, liveWorkflow),
        )
      ) {
        return false;
      }
      if (
        journalMonthFilter !== "ALL" &&
        documentCreatedMonthKey(document.createdAt) !== journalMonthFilter
      ) {
        return false;
      }
      if (journalReadinessFilter !== "ALL") {
        const liveFields = liveJournalFields(document);
        const skippedSzch = readDocumentSkippedDueToSzch(liveFields);
        if (journalReadinessFilter === "SKIPPED_SZCH") {
          return skippedSzch;
        }
        if (skippedSzch) return false;
        const complete =
          getDocumentProgressPercent(
            ubdLiveDocument(document, liveWorkflow),
          ) >= 100;
        if (journalReadinessFilter === "COMPLETE" && !complete) return false;
        if (document.type === "ubdReport") {
          const readyToSend = ubdDocumentIsReadyToSend(
            document,
            liveFields,
            liveWorkflow,
          );
          if (journalReadinessFilter === "READY_TO_SEND" && !readyToSend) {
            return false;
          }
          if (journalReadinessFilter === "INCOMPLETE" && readyToSend) {
            return false;
          }
        } else {
          if (journalReadinessFilter === "READY_TO_SEND") return false;
          if (journalReadinessFilter === "INCOMPLETE") {
            const incomplete =
              documentHasEmptyInputs(document.type, liveFields) ||
              documentHasBasisDateMismatch(document.type, liveFields);
            if (!incomplete) return false;
          }
        }
      }
      return true;
    });

    const direction = journalSortDirection === "desc" ? -1 : 1;
    return [...filtered].sort((left, right) => {
      let compare = 0;
      if (journalSortField === "progress") {
        compare =
          getDocumentProgressPercent(left) - getDocumentProgressPercent(right);
      } else if (journalSortField === "taskPeriodStart") {
        compare =
          readDocumentUbdTaskPeriodStartTimestamp(left) -
          readDocumentUbdTaskPeriodStartTimestamp(right);
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
    journalDocuments,
    journalMonthFilter,
    journalNameQuery,
    journalReadinessFilter,
    journalSortDirection,
    journalSortField,
    journalStatusFilters,
    journalTypeFilter,
    lostMilitaryIdFields,
    mode,
    selectedDocumentId,
    skippedDueToSzch,
    salaryFields,
    serviceCharacteristicFields,
    ticketFields,
    ubdFields,
    ubdRestoreFields,
    form6Fields,
    form12Fields,
    zhbdCertificateFields,
    workflow,
  ]);

  const journalUbdExportBlockedIds = useMemo(() => {
    const blocked = new Set<string>();
    for (const document of filteredJournalDocuments) {
      const liveFields = liveJournalFields(document);
      if (readDocumentSkippedDueToSzch(liveFields)) {
        blocked.add(document.id);
        continue;
      }
      if (document.type !== "ubdReport") continue;
      const liveWorkflow =
        document.id === selectedDocumentId
          ? mergeSalaryWorkflow({
              ...(document.workflow && typeof document.workflow === "object"
                ? document.workflow
                : {}),
              ...workflow,
            })
          : undefined;
      if (ubdDocumentIsComplete(document, liveWorkflow)) {
        blocked.add(document.id);
      }
    }
    return blocked;
  }, [
    filteredJournalDocuments,
    form12Fields,
    form6Fields,
    lostMilitaryIdFields,
    mode,
    selectedDocumentId,
    salaryFields,
    serviceCharacteristicFields,
    skippedDueToSzch,
    ticketFields,
    ubdFields,
    ubdRestoreFields,
    zhbdCertificateFields,
    workflow,
  ]);

  const journalExportDocuments = useMemo(
    () =>
      filteredJournalDocuments.filter(
        (document) =>
          !journalExportDeselectedIds[document.id] &&
          !journalUbdExportBlockedIds.has(document.id),
      ),
    [
      filteredJournalDocuments,
      journalExportDeselectedIds,
      journalUbdExportBlockedIds,
    ],
  );

  const journalExportEligibleCount =
    filteredJournalDocuments.length - journalUbdExportBlockedIds.size;
  const journalExportSelectedCount = journalExportDocuments.length;
  const journalExportAllSelected =
    journalExportEligibleCount > 0 &&
    journalExportSelectedCount === journalExportEligibleCount;
  const journalExportSomeSelected =
    journalExportSelectedCount > 0 && !journalExportAllSelected;

  const journalStatusFilterLabel =
    journalStatusFilters.length === 0
      ? "Усі статуси"
      : journalStatusFilters.length === 1
        ? journalStatusFilters[0]
        : `Статусів: ${journalStatusFilters.length}`;
  const journalStatusFilterExportLabel =
    journalStatusFilters.length === 0
      ? "Усі статуси"
      : journalStatusFilters.join(", ");

  useEffect(() => {
    const visible = new Set(
      filteredJournalDocuments.map((document) => document.id),
    );
    setJournalExportDeselectedIds((prev) => {
      let changed = false;
      const next: Record<string, true> = {};
      for (const id of Object.keys(prev)) {
        if (visible.has(id)) next[id] = true;
        else changed = true;
      }
      if (!changed && Object.keys(next).length === Object.keys(prev).length) {
        return prev;
      }
      return next;
    });
  }, [filteredJournalDocuments]);

  useEffect(() => {
    if (!journalStatusMenuOpen) return;
    const onPointerDown = (event: MouseEvent) => {
      const root = journalStatusMenuRef.current;
      if (!root || root.contains(event.target as Node)) return;
      setJournalStatusMenuOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setJournalStatusMenuOpen(false);
    };
    window.addEventListener("mousedown", onPointerDown);
    window.addEventListener("keydown", onKeyDown);
    return () => {
      window.removeEventListener("mousedown", onPointerDown);
      window.removeEventListener("keydown", onKeyDown);
    };
  }, [journalStatusMenuOpen]);

  const toggleJournalStatusFilter = (title: string) => {
    setJournalStatusFilters((current) => {
      if (current.length === 0) return [title];
      if (current.includes(title)) {
        return current.filter((item) => item !== title);
      }
      return [...current, title];
    });
  };

  const toggleJournalExportDocument = (documentId: string, selected: boolean) => {
    setJournalExportDeselectedIds((prev) => {
      const next = { ...prev };
      if (selected) delete next[documentId];
      else next[documentId] = true;
      return next;
    });
  };

  const toggleJournalExportAll = (selected: boolean) => {
    if (selected) {
      setJournalExportDeselectedIds({});
      return;
    }
    setJournalExportDeselectedIds(
      Object.fromEntries(
        filteredJournalDocuments
          .filter((document) => !journalUbdExportBlockedIds.has(document.id))
          .map((document) => [document.id, true as const]),
      ),
    );
  };

  const exportDocumentJournal = async () => {
    if (!filteredJournalDocuments.length) {
      setDocumentMessage("Немає рядків для експорту за поточними фільтрами.");
      return;
    }
    if (!journalExportDocuments.length) {
      setDocumentMessage("Немає вибраних рядків для експорту.");
      return;
    }

    setIsExportingDocumentJournal(true);
    try {
      const periodFilterLabel =
        journalMonthFilter === "ALL"
          ? "Усі місяці"
          : formatJournalMonthLabel(journalMonthFilter);

      if (journalTypeFilter === "ubdReport") {
        setDocumentMessage("Готую таблицю УБД «Не подавалися»…");
        const [staffImport, cachedRoster] = await Promise.all([
          readDataCache<StaffSheetImportSnapshot>(CacheKeys.staffSheetImport),
          readDataCache<BackendPersonnelRosterLatest>(CacheKeys.rosterLatest),
        ]);
        const rosterLatest =
          cachedRoster ??
          (await api.getLatestPersonnelRoster().catch(() => null));
        const rosterCallSignIndex = buildUbdRosterCallSignIndex([
          ...(staffImport?.rows ?? []),
          ...mapRosterLatestToPreviewRows(rosterLatest),
        ]);

        const uniqueIds = [
          ...new Set(
            journalExportDocuments
              .map((document) => document.personExternalId)
              .filter(Boolean),
          ),
        ];
        const profileById = new Map<
          string,
          Awaited<ReturnType<typeof api.getPersonnelProfile>> | null
        >();
        const concurrency = 4;
        for (let index = 0; index < uniqueIds.length; index += concurrency) {
          const chunk = uniqueIds.slice(index, index + concurrency);
          const results = await Promise.all(
            chunk.map(async (personId) => {
              try {
                const profile = await api.getPersonnelProfile(personId);
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

        const rows = journalExportDocuments.map((document) => {
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
          const personName = getDocumentPersonName(journalDocument);
          const rosterMatch = lookupUbdRosterCallSign(
            rosterCallSignIndex,
            journalDocument.personExternalId,
            personName,
          );
          return buildUbdNotSubmittedRowFromDocument({
            document: journalDocument,
            profile: profileById.get(journalDocument.personExternalId) ?? null,
            personnelStatus: readDocumentPersonStatus(
              journalDocument,
              personStatusById,
            ),
            rosterRow: rosterMatch?.row ?? null,
            rosterCallSign: rosterMatch?.callSign ?? "",
          });
        });

        const { fileName, rowCount } = await exportUbdNotSubmittedExcel({
          rows,
          periodFilterLabel,
        });
        const skippedComplete = [...journalUbdExportBlockedIds].filter(
          (id) => {
            const document = filteredJournalDocuments.find(
              (item) => item.id === id,
            );
            if (!document) return false;
            return !readDocumentSkippedDueToSzch(liveJournalFields(document));
          },
        ).length;
        const skippedSzch = journalUbdExportBlockedIds.size - skippedComplete;
        const skipParts = [
          skippedComplete
            ? `з повним прогресом: ${skippedComplete}`
            : "",
          skippedSzch ? `СЗЧ: ${skippedSzch}` : "",
        ].filter(Boolean);
        setDocumentMessage(
          skipParts.length
            ? `Експортовано УБД «Не подавалися»: ${rowCount} рядків · ${fileName}. Пропущено ${skipParts.join(" · ")}.`
            : `Експортовано УБД «Не подавалися»: ${rowCount} рядків · ${fileName}`,
        );
        return;
      }

      const { fileName, rowCount } = await exportDocumentsJournalExcel({
        rows: journalExportDocuments.map((document) => {
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
            formPurpose: readDocumentFormPurpose(
              journalDocument,
              liveJournalFields(journalDocument),
            ),
            createdAt: journalDocument.createdAt,
            updatedAt: journalDocument.updatedAt,
            taskPeriodEnd: readDocumentUbdTaskPeriodStartLabel(journalDocument),
          };
        }),
        includeUbdExitDate: showUbdExitDateColumn,
        includeFormPurpose: journalExportDocuments.some(
          (document) =>
            document.type === "form6Report" ||
            document.type === "form12Report",
        ),
        compactFormExport: journalExportDocuments.every(
          (document) =>
            document.type === "form6Report" ||
            document.type === "form12Report",
        ),
        typeFilterLabel:
          JOURNAL_DOCUMENT_TYPE_FILTERS.find(
            (item) => item.value === journalTypeFilter,
          )?.label ?? "Усі документи",
        statusFilterLabel: journalStatusFilterExportLabel,
        periodFilterLabel,
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

  const exportUbdBulkWordFromJournal = async () => {
    if (journalTypeFilter !== "ubdReport") {
      setDocumentMessage("Спочатку відфільтруйте журнал за типом «Рапорт на УБД».");
      return;
    }
    if (!journalExportDocuments.length) {
      setDocumentMessage("Немає вибраних рядків для спільного рапорту УБД.");
      return;
    }

    setIsExportingDocumentJournal(true);
    try {
      setDocumentMessage("Формую спільний Word-рапорт УБД…");

      const completeDocuments = journalExportDocuments.filter((document) => {
        if (document.type !== "ubdReport") return false;
        const liveFields =
          document.id === selectedDocumentId && mode === "ubdReport"
            ? (ubdFields as unknown as Record<string, unknown>)
            : ((document.fields || {}) as Record<string, unknown>);
        return !ubdDocumentHasEmptyInputs(liveFields);
      });

      if (!completeDocuments.length) {
        setDocumentMessage(
          "Серед вибраних немає «зелених» (заповнених) рапортів УБД для спільного Word.",
        );
        return;
      }

      const people = completeDocuments.map((document) => {
        const liveFields =
          document.id === selectedDocumentId && mode === "ubdReport"
            ? (ubdFields as unknown as Record<string, unknown>)
            : ((document.fields || {}) as Record<string, unknown>);
        const resolved = resolveUbdFieldsForGapCheck(liveFields);
        const fullName =
          String(resolved.fullName ?? "").trim() ||
          getDocumentPersonName(document);
        return {
          rank: String(resolved.rank ?? "").trim(),
          fullName,
          staffPosition: capitalizeReportPosition(
            String(resolved.staffPosition ?? ""),
          ),
          birthDate: String(resolved.birthDate ?? "").trim(),
          rnokpp: String(resolved.rnokpp ?? "").trim(),
          taskPeriod: String(resolved.taskPeriod ?? "").trim(),
          taskPlace: String(resolved.taskPlace ?? "").trim(),
          basisNumber: String(resolved.basisNumber ?? "").trim(),
          basisDate: String(resolved.basisDate ?? "").trim(),
          basis: String(resolved.basis ?? "").trim(),
        };
      });

      let signatories: DocumentSignatorySnapshot[] =
        mode === "ubdReport" && ubdFields.signatories.length
          ? ubdFields.signatories
          : [];
      if (!signatories.length) {
        try {
          const configured = await api.listDocumentSignatories("ubdReport");
          signatories = configured.length
            ? snapshotSignatories(configured)
            : legacyUbdSignatories();
        } catch {
          signatories = legacyUbdSignatories();
        }
      }

      const firstCommander = String(
        completeDocuments[0]?.fields?.commander ?? "",
      ).trim();
      const commander =
        (mode === "ubdReport" && ubdFields.commander.trim()) ||
        firstCommander ||
        "Командиру військової частини A4862";

      const blob = await createUbdBulkWordBlob({
        commander,
        people,
        signatories: signatories.map((signatory) =>
          signatory.signatureData === UBD_COMMANDER_SIGNATURE_SRC
            ? { ...signatory, signatureData: commanderSignatureSrc }
            : signatory,
        ),
      });

      const skipped = journalExportDocuments.length - completeDocuments.length;
      const fileName = `УБД рапорт спільний · ${people.length} осіб · ${dayjs().format("DD.MM.YYYY")}.docx`;
      downloadBlob(fileName, blob);
      setDocumentMessage(
        skipped > 0
          ? `Word УБД: ${people.length} осіб (пропущено незаповнених: ${skipped}) · ${fileName}`
          : `Word УБД: ${people.length} осіб · ${fileName}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося сформувати спільний Word УБД: ${error.message}`
          : "Не вдалося сформувати спільний Word УБД.",
      );
    } finally {
      setIsExportingDocumentJournal(false);
    }
  };

  const toggleJournalBulkProgressStep = (stepKey: string) => {
    setJournalBulkProgressSteps((current) => ({
      ...current,
      [stepKey]: !current[stepKey],
    }));
  };

  const applyBulkJournalProgress = async () => {
    const targets = journalExportDocuments.filter(
      (document) => document.type === "ubdReport",
    );
    if (!targets.length) {
      setDocumentMessage("Немає обраних рапортів УБД для масового прогресу.");
      return;
    }
    const markedSteps = ubdWorkflowSteps.filter(
      (step) => journalBulkProgressSteps[step.key],
    );
    if (!markedSteps.length) {
      setDocumentMessage("Оберіть хоча б один крок прогресу.");
      return;
    }

    const backupEntries = targets.map(ubdProgressBackupEntryFromDocument);
    saveUbdProgressBackup({
      reason: "before-bulk-apply",
      label: `Перед масовим прогресом · ${backupEntries.length} рапортів`,
      entries: backupEntries,
    });
    refreshUbdProgressBackups();

    setIsApplyingBulkJournalProgress(true);
    try {
      let updated = 0;
      let skipped = 0;
      let fail = 0;
      const undoEntries: UbdProgressBackupEntry[] = [];
      const concurrency = 4;
      for (let index = 0; index < targets.length; index += concurrency) {
        const chunk = targets.slice(index, index + concurrency);
        const results = await Promise.all(
          chunk.map(async (document) => {
            try {
              const existing = mergeSalaryWorkflow(document.workflow);
              const currentMap = resolveWorkflowCompletedMap(
                existing,
                ubdWorkflowSteps,
                document.type,
                document.status || existing.currentStatus,
              );
              const hasNewSteps = ubdWorkflowSteps.some(
                (step) =>
                  journalBulkProgressSteps[step.key] && !currentMap[step.key],
              );
              if (!hasNewSteps) {
                return "skipped" as const;
              }
              undoEntries.push(ubdProgressBackupEntryFromDocument(document));
              const nextWorkflow = mergeBulkStepsIntoWorkflow(
                existing,
                ubdWorkflowSteps,
                journalBulkProgressSteps,
                document.type,
                document.status || existing.currentStatus,
              );
              const saved = await api.updatePersonDocument(
                document.personExternalId,
                document.id,
                {
                  title: document.title,
                  status: nextWorkflow.currentStatus,
                  workflow: nextWorkflow as unknown as Record<string, unknown>,
                },
                { suppressErrorToast: true },
              );
              applyUpdatedDocument(saved);
              if (document.id === selectedDocumentId) {
                setWorkflow(nextWorkflow);
              }
              return "updated" as const;
            } catch {
              return "fail" as const;
            }
          }),
        );
        updated += results.filter((result) => result === "updated").length;
        skipped += results.filter((result) => result === "skipped").length;
        fail += results.filter((result) => result === "fail").length;
      }
      if (updated > 0 && undoEntries.length) {
        setJournalBulkProgressUndo(undoEntries);
        savePersistedUbdBulkProgressUndo(undoEntries);
      }
      const parts = [
        updated > 0 ? `оновлено ${updated}` : "",
        skipped > 0 ? `без змін ${skipped}` : "",
        fail > 0 ? `помилок ${fail}` : "",
      ].filter(Boolean);
      setDocumentMessage(
        parts.length
          ? `Масовий прогрес: ${parts.join(", ")}.${
              updated > 0 ? " Можна скасувати останню зміну." : ""
            }`
          : "Масовий прогрес не застосовано.",
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося застосувати прогрес: ${error.message}`
          : "Не вдалося застосувати прогрес.",
      );
    } finally {
      setIsApplyingBulkJournalProgress(false);
    }
  };

  useEffect(() => {
    const persistedUndo = loadPersistedUbdBulkProgressUndo();
    if (persistedUndo) {
      setJournalBulkProgressUndo(persistedUndo);
    }
  }, []);

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
            : requestedDocumentType === "lost-military-id"
              ? "lostMilitaryId"
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
                      : savedMode === "lostMilitaryId"
                        ? "lostMilitaryId"
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
      setLostMilitaryIdFields(createLostMilitaryIdFields(person, nextSummary));
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

  useEffect(() => {
    const skipped = readDocumentSkippedDueToSzch(
      (selectedDocument?.fields || {}) as Record<string, unknown>,
    );
    skippedDueToSzchRef.current = skipped;
    setSkippedDueToSzch(skipped);
  }, [selectedDocument?.fields?.skippedDueToSzch, selectedDocumentId]);
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
          staffPosition: capitalizeReportPosition(
            staffLooksGeneric ? nextStaff : current.staffPosition,
          ),
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

  // УБД: підтягнути період/місце/посаду зі «Статус бійців» + Загальний список.
  useEffect(() => {
    if (!isPersonDocumentMode) return;
    if (mode !== "ubdReport") return;
    const lookupName = summary.name?.trim() || "";
    if (!lookupName || lookupName === "Особа не вибрана") return;

    let cancelled = false;
    void (async () => {
      try {
        const latest = await fetchWithCache({
          key: CacheKeys.rosterLatest,
          fetcher: () => api.getLatestPersonnelRoster(),
          isChanged: jsonChanged,
        });
        if (cancelled) return;
        const rows = mapRosterLatestToPreviewRows(latest);
        const match = findRosterRowByPersonName(rows, lookupName);
        if (!match) return;

        if (selectedPerson) {
          setSelectedPerson(applyRosterFieldsToPerson(selectedPerson, match));
        }

        const enriched = createUbdFields(
          applyRosterFieldsToPerson(selectedPerson ?? match, match),
          buildPersonSummary(
            applyRosterFieldsToPerson(selectedPerson ?? match, match),
          ),
        );
        setUbdFields((current) => {
          const next = { ...current };
          let changed = false;
          if (!current.taskPeriod.trim() && enriched.taskPeriod.trim()) {
            next.taskPeriod = enriched.taskPeriod;
            const picked = pickUbdBasisOrderForTaskPeriod(enriched.taskPeriod);
            if (picked) {
              next.basisNumber = picked.number;
              next.basisDate = picked.date;
              next.basis = formatUbdBasisText(picked.number, picked.date);
            }
            next.basisNotReady = !ubdBasisDateMatchesTaskPeriod(
              next.taskPeriod,
              next.basisDate,
            );
            changed = true;
          }
          if (
            (!current.taskPlace.trim() || current.taskPlace.trim() === "-") &&
            enriched.taskPlace.trim()
          ) {
            next.taskPlace = enriched.taskPlace;
            changed = true;
          }
          if (!current.staffPosition.trim() && enriched.staffPosition.trim()) {
            next.staffPosition = capitalizeReportPosition(enriched.staffPosition);
            changed = true;
          }
          if (!current.rnokpp.trim() && enriched.rnokpp.trim()) {
            next.rnokpp = enriched.rnokpp;
            changed = true;
          }
          return changed ? next : current;
        });
      } catch (error) {
        console.warn("[УБД] не вдалося підтягнути Статус бійців з ранкового:", error);
      }
    })();

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- enrich once per person/document open
  }, [isPersonDocumentMode, mode, summary.name, selectedDocumentId]);

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
          : requestedDocumentType === "lost-military-id"
            ? "lostMilitaryId"
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
        targetDocumentType !== "temporaryMilitaryId" &&
        targetDocumentType !== "lostMilitaryId") ||
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
        const relatedDocumentMetadata = selectedPerson
          ? await loadPersonDocumentsForRow(selectedPerson, {
              anketaFullName: summary.name,
            })
          : [];
        const relatedDocumentIds = new Set([
          personExternalId,
          ...relatedDocumentMetadata.map(
            (document) => document.personExternalId,
          ),
        ]);
        const documentsPromise = Promise.all(
          [...relatedDocumentIds].map((id) =>
            api.listPersonDocuments(id, { full: true }).catch(() => []),
          ),
        ).then((groups) => {
          const unique = new Map<string, BackendPersonDocument>();
          for (const document of groups.flat()) unique.set(document.id, document);
          return [...unique.values()].sort(
            (left, right) =>
              new Date(right.updatedAt).getTime() -
              new Date(left.updatedAt).getTime(),
          );
        });
        const [documents, configuredRecords, personPhoto, questionnaires] =
          await Promise.all([
            documentsPromise,
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
              : targetDocumentType === "lostMilitaryId"
                ? loadLostMilitaryIdSignatoryRecords()
                : Promise.resolve([]),
            api.getPersonPhoto(personExternalId).catch(() => null),
            api.listPersonQuestionnaires().catch(() => []),
          ]);
        let questionnaire =
          questionnaires.find(
            (item) => item.personExternalId === personExternalId,
          ) ?? null;
        if (!questionnaire && selectedPerson) {
          const resolved = await loadPersonQuestionnaireForRow(
            selectedPerson,
            { anketaFullName: summary.name },
          );
          questionnaire = resolved.questionnaire
            ? {
                personExternalId: resolved.resolvedExternalId,
                fileName: resolved.questionnaire.fileName,
              }
            : null;
        }
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
        const lostMilitaryIdDefaults = createLostMilitaryIdFields(
          selectedPerson,
          summary,
          toLostMilitaryIdSignatories(
            configured.length ? configured : legacyUbdSignatories(),
          ),
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
            const latest = await api.listPersonDocuments(personExternalId, {
              full: true,
            });
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
            lostMilitaryIdFields: lostMilitaryIdDefaults,
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
            setLostMilitaryIdFields(lostMilitaryIdDefaults);
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
            setLostMilitaryIdFields(lostMilitaryIdDefaults);
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
        setLostMilitaryIdFields({
          ...mergeLostMilitaryIdFields(lostMilitaryIdDefaults, active.fields),
          signatories: lostMilitaryIdDefaults.signatories,
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
        const normalizeName = (value: string) =>
          value.toLocaleLowerCase("uk-UA").replace(/\s+/g, " ").trim();
        const overviewRows = overview?.rows ?? [];
        const overviewNameById = new Map(
          overviewRows
            .filter((row) => row.externalId && row.name)
            .map((row) => [row.externalId, row.name]),
        );
        const hydratedDocuments = documents.map((document) => ({
          ...document,
          personName:
            String(document.personName ?? "").trim() ||
            overviewNameById.get(document.personExternalId) ||
            null,
        }));
        const statusesById = Object.fromEntries(
          overviewRows
            .filter((row) => row.externalId)
            .map((row) => [row.externalId, overviewStatusSnapshot(row)]),
        );
        const statusByUniqueName = new Map<string, PersonStatusSnapshot | null>();
        for (const row of overviewRows) {
          const key = normalizeName(row.name);
          if (!key) continue;
          statusByUniqueName.set(
            key,
            statusByUniqueName.has(key) ? null : overviewStatusSnapshot(row),
          );
        }
        for (const document of hydratedDocuments) {
          if (statusesById[document.personExternalId]) continue;
          const status = statusByUniqueName.get(
            normalizeName(getDocumentPersonName(document)),
          );
          if (status) statusesById[document.personExternalId] = status;
        }
        setAllPersonDocuments(hydratedDocuments);
        setPersonStatusById(statusesById);
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
          force: true,
          fetcher: () => api.getPersonnelOverview(),
          isChanged: jsonChanged,
        }).catch(() => null),
      ]);
      applyJournal(documents, overview);
      const sessionBackup = saveSessionJournalUbdBackup(
        documents,
        ubdProgressBackupEntryFromDocument,
      );
      if (sessionBackup) {
        setUbdProgressBackups(loadUbdProgressBackups());
      }
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
  const workflowCompletedMap = resolveWorkflowCompletedMap(
    workflow,
    activeWorkflowSteps,
    selectedDocument?.type || mode,
    workflow.currentStatus,
  );
  const completedCount = activeWorkflowSteps.filter(
    (step) => workflowCompletedMap[step.key],
  ).length;
  const progressPercent = activeWorkflowSteps.length
    ? Math.round((completedCount / activeWorkflowSteps.length) * 100)
    : 0;

  const updateDefaultField = (key: DefaultDocumentFieldKey, value: string) => {
    setFields((current) => ({ ...current, [key]: value }));
  };

  const openPersonDocument = async (documentSummary: BackendPersonDocument) => {
    let document = documentSummary;
    if (!Object.prototype.hasOwnProperty.call(documentSummary, "files")) {
      setDocumentMessage("Завантажую повні дані документа…");
      try {
        const fullDocuments = await api.listPersonDocuments(
          documentSummary.personExternalId,
          { full: true },
        );
        const loaded = fullDocuments.find(
          (item) => item.id === documentSummary.id,
        );
        if (!loaded) {
          throw new Error("Документ не знайдено в БД");
        }
        document = {
          ...documentSummary,
          ...loaded,
          personName: loaded.personName || documentSummary.personName,
        };
        setAllPersonDocuments((current) =>
          current.map((item) => (item.id === document.id ? document : item)),
        );
        setPersonDocuments((current) =>
          current.some((item) => item.id === document.id)
            ? current.map((item) => (item.id === document.id ? document : item))
            : [document, ...current],
        );
      } catch (error) {
        setDocumentMessage(
          error instanceof Error
            ? `Не вдалося завантажити документ: ${error.message}`
            : "Не вдалося завантажити повні дані документа.",
        );
        return;
      }
    }

    const nextPersonId = String(document.personExternalId || "").trim();
    const documentPersonName = getDocumentPersonName(document);
    const normalizePersonName = (value: string) =>
      value.toLocaleLowerCase("uk-UA").replace(/\s+/g, " ").trim();
    let personForDocument = selectedPerson;
    if (
      documentPersonName &&
      !documentPersonName.startsWith("ID ") &&
      normalizePersonName(getPersonDisplayName(selectedPerson)) !==
        normalizePersonName(documentPersonName)
    ) {
      const profile = await api
        .getPersonnelProfile(nextPersonId, documentPersonName)
        .catch(() => null);
      const rowValues = (value: unknown) => {
        if (!value || typeof value !== "object" || Array.isArray(value)) {
          return {};
        }
        const record = value as Record<string, unknown>;
        const values = record.values;
        return values && typeof values === "object" && !Array.isArray(values)
          ? (values as Record<string, unknown>)
          : {};
      };
      const historical = rowValues(profile?.ejournal?.historicalOosRow);
      const current = rowValues(profile?.ejournal?.oosRow);
      const roster = rowValues(profile?.roster?.row);
      if (
        Object.keys(historical).length ||
        Object.keys(current).length ||
        Object.keys(roster).length
      ) {
        personForDocument = {
          ...historical,
          ...current,
          ...roster,
          __dbRowId:
            String(
              (profile?.ejournal?.oosRow as Record<string, unknown> | null)
                ?.id ?? "",
            ) || undefined,
          піб:
            String(
              roster.піб ??
                current[
                  "прізвище_за_наявності_імя_по_батькові_за_наявності"
                ] ??
                historical[
                  "прізвище_за_наявності_імя_по_батькові_за_наявності"
                ] ??
                documentPersonName,
            ).trim(),
        };
        setSelectedPerson(personForDocument);
      }
    }
    const summaryForDocument = buildPersonSummary(personForDocument);

    if (isPersonDocumentMode) {
      const currentParams = new URLSearchParams(window.location.search);
      const nextRoute = buildDocumentRoute({
        personExternalId:
          nextPersonId || requestedPersonId || personExternalId,
        rowId: currentParams.get("rowId") || undefined,
        documentId: document.id,
        type: document.type,
      });
      window.history.replaceState(
        { ...(window.history.state as object | null), page: "documents" },
        "",
        nextRoute,
      );
    }
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
    void loadPersonQuestionnaireForRow(personForDocument, {
      anketaFullName: documentPersonName,
    })
      .then(({ questionnaire: item, resolvedExternalId }) => {
        setPersonQuestionnaire(
          item
            ? {
                personExternalId:
                  item.personExternalId || resolvedExternalId,
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
        personForDocument,
        summaryForDocument,
        configured.length ? configured : legacyUbdSignatories(),
      );
      const merged = mergeUbdFields(defaults, document.fields);
      setUbdFields({ ...merged, signatories: defaults.signatories });
      setDocumentFiles(mergeDocumentFiles(document.files));
    } else if (document.type === "form6Report") {
      setMode("form6Report");
      const configured = snapshotSignatories(await loadForm6SignatoryRecords());
      const defaults = createForm6Fields(
        personForDocument,
        summaryForDocument,
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
        personForDocument,
        summaryForDocument,
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
        personForDocument,
        summaryForDocument,
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
        personForDocument,
        summaryForDocument,
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
            getPersonFullPositionTitle(personForDocument) ||
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
        personForDocument,
        summaryForDocument,
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
        summaryForDocument,
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
    } else if (document.type === "lostMilitaryId") {
      setMode("lostMilitaryId");
      const configured = snapshotSignatories(
        await loadLostMilitaryIdSignatoryRecords(),
      );
      const defaults = createLostMilitaryIdFields(
        personForDocument,
        summaryForDocument,
        toLostMilitaryIdSignatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      setLostMilitaryIdFields({
        ...mergeLostMilitaryIdFields(defaults, document.fields),
        signatories: defaults.signatories,
      });
      setDocumentFiles(mergeDocumentFiles(document.files));
    } else {
      const defaults = createSalaryFields(personForDocument, summaryForDocument);
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
      let personForFields = selectedPerson;
      try {
        const latest = await api.getLatestPersonnelRoster();
        const rows = mapRosterLatestToPreviewRows(latest);
        const match = findRosterRowByPersonName(rows, summary.name);
        if (match) {
          personForFields = applyRosterFieldsToPerson(
            selectedPerson ?? match,
            match,
          );
          setSelectedPerson(personForFields);
        }
      } catch {
        /* keep selectedPerson */
      }
      const personSummary = buildPersonSummary(personForFields);
      const configured = snapshotSignatories(
        await api.listDocumentSignatories("ubdReport"),
      );
      const defaults = createUbdFields(
        personForFields,
        personSummary,
        configured.length ? configured : legacyUbdSignatories(),
      );
      const fieldPayload = {
        ...defaults,
        personStatus: createPersonStatusSnapshot(personForFields, personSummary)
          .label,
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
      setWorkflow(nextWorkflow);
      setDocumentMessage(
        `Створено УБД рапорт · ${new Date(created.createdAt).toLocaleString("uk-UA")}`,
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
      const created = await api.createPersonDocument(
        personExternalId,
        {
          type: "form12Report",
          title: `Форма 12 · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
          status: "document",
          fields: fieldPayload as unknown as Record<string, unknown>,
          workflow: nextWorkflow as unknown as Record<string, unknown>,
          files: { ubdScans: [] },
        },
        { suppressErrorToast: true },
      );
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

  const createLostMilitaryIdDocument = async () => {
    if (!personExternalId) return;
    const nextWorkflow = {
      ...createEmptyWorkflow(),
      currentStatus: "document",
    };
    setIsSavingDocument(true);
    setDocumentMessage("Створюю документ про втрату військового квитка...");
    try {
      const configured = snapshotSignatories(
        await loadLostMilitaryIdSignatoryRecords(),
      );
      const defaults = createLostMilitaryIdFields(
        selectedPerson,
        summary,
        toLostMilitaryIdSignatories(
          configured.length ? configured : legacyUbdSignatories(),
        ),
      );
      const fieldPayload = {
        ...defaults,
        personStatus: createPersonStatusSnapshot(selectedPerson, summary).label,
      };
      const created = await api.createPersonDocument(personExternalId, {
        type: "lostMilitaryId",
        title: `Втрата військового квитка · ${dayjs().format("DD.MM.YYYY HH:mm")}`,
        status: "document",
        fields: fieldPayload as unknown as Record<string, unknown>,
        workflow: nextWorkflow as unknown as Record<string, unknown>,
        files: { ubdScans: [] },
      });
      setPersonDocuments((current) => [created, ...current]);
      setAllPersonDocuments((current) => [created, ...current]);
      setSelectedDocumentId(created.id);
      setLostMilitaryIdFields(mergeLostMilitaryIdFields(defaults, created.fields));
      setDocumentFiles(mergeDocumentFiles(created.files));
      setWorkflow(mergeSalaryWorkflow(created.workflow));
      setMode("lostMilitaryId");
      setDocumentMessage(
        `Створено документ: ${created.title} · ${new Date(created.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося створити документ: ${error.message}`
          : "Не вдалося створити документ про втрату військового квитка.",
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
      } else if (
        selectedDocument.type === "lostMilitaryId" ||
        mode === "lostMilitaryId"
      ) {
        await saveLostMilitaryIdDocument(lostMilitaryIdFields, nextFiles);
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
    } else if (
      selectedDocument?.type === "lostMilitaryId" ||
      mode === "lostMilitaryId"
    ) {
      void saveLostMilitaryIdDocument(lostMilitaryIdFields, nextFiles);
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

  const withPersistedDocumentFlags = <T extends object>(fields: T) => ({
    ...fields,
    personStatus: nextPersonStatusLabel(),
    skippedDueToSzch: skippedDueToSzchRef.current,
  });

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

  const restoreUbdProgressEntries = async (
    entries: UbdProgressBackupEntry[],
    successMessage: (restored: number, total: number, fail: number) => string,
  ) => {
    if (!entries.length) {
      setDocumentMessage("Немає збереженого прогресу для відновлення.");
      return;
    }

    setIsApplyingBulkJournalProgress(true);
    try {
      let restored = 0;
      let fail = 0;
      const concurrency = 4;
      for (let index = 0; index < entries.length; index += concurrency) {
        const chunk = entries.slice(index, index + concurrency);
        const results = await Promise.all(
          chunk.map(async (entry) => {
            try {
              const workflow = workflowFromBackupEntry(entry);
              const saved = await api.updatePersonDocument(
                entry.personExternalId,
                entry.documentId,
                {
                  title: entry.title,
                  status: entry.status,
                  workflow: workflow as unknown as Record<string, unknown>,
                },
                { suppressErrorToast: true },
              );
              applyUpdatedDocument(saved);
              if (entry.documentId === selectedDocumentId) {
                setWorkflow(workflow);
              }
              return true;
            } catch {
              return false;
            }
          }),
        );
        restored += results.filter(Boolean).length;
        fail += results.filter((result) => !result).length;
      }
      setDocumentMessage(successMessage(restored, entries.length, fail));
      return { restored, fail, total: entries.length };
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося відновити прогрес: ${error.message}`
          : "Не вдалося відновити прогрес.",
      );
      return { restored: 0, fail: entries.length, total: entries.length };
    } finally {
      setIsApplyingBulkJournalProgress(false);
    }
  };

  const undoBulkJournalProgress = async () => {
    if (!journalBulkProgressUndo?.length) {
      setDocumentMessage("Немає останньої масової зміни для скасування.");
      return;
    }
    const entries = journalBulkProgressUndo;
    const result = await restoreUbdProgressEntries(entries, (restored, total, fail) =>
      fail > 0
        ? `Скасовано прогрес: ${restored} з ${total} (помилок: ${fail}).`
        : `Скасовано останню масову зміну прогресу для ${restored} рапортів УБД.`,
    );
    if (result?.fail === 0) {
      setJournalBulkProgressUndo(null);
      savePersistedUbdBulkProgressUndo(null);
    }
  };

  const restoreUbdProgressBackup = async (backup: UbdProgressBackup) => {
    const result = await restoreUbdProgressEntries(
      backup.entries,
      (restored, total, fail) =>
        fail > 0
          ? `Відновлено з копії «${backup.label}»: ${restored} з ${total} (помилок: ${fail}).`
          : `Відновлено прогрес з копії «${backup.label}» для ${restored} рапортів УБД.`,
    );
    if (result && result.fail === 0 && backup.reason === "before-bulk-apply") {
      setJournalBulkProgressUndo(null);
      savePersistedUbdBulkProgressUndo(null);
    }
  };

  const saveManualUbdProgressBackup = () => {
    const ubdDocuments = allPersonDocuments.filter(
      (document) => document.type === "ubdReport",
    );
    if (!ubdDocuments.length) {
      setDocumentMessage("Немає рапортів УБД для резервної копії.");
      return;
    }
    const backup = saveUbdProgressBackup({
      reason: "manual",
      label: `Ручна копія · ${ubdDocuments.length} рапортів`,
      entries: ubdDocuments.map(ubdProgressBackupEntryFromDocument),
    });
    refreshUbdProgressBackups();
    setDocumentMessage(
      `Збережено резервну копію прогресу: ${backup.entries.length} рапортів · ${formatUbdProgressBackupWhen(backup.savedAt)}.`,
    );
  };

  const restoreUbdProgressFromArtifacts = async () => {
    const targets = journalExportDocuments.filter(
      (document) => document.type === "ubdReport",
    );
    if (!targets.length) {
      setDocumentMessage("Оберіть рапорти УБД для часткового відновлення.");
      return;
    }

    setIsApplyingBulkJournalProgress(true);
    try {
      let updated = 0;
      let skipped = 0;
      let fail = 0;
      const concurrency = 4;
      for (let index = 0; index < targets.length; index += concurrency) {
        const chunk = targets.slice(index, index + concurrency);
        const results = await Promise.all(
          chunk.map(async (document) => {
            try {
              const existing = mergeSalaryWorkflow(document.workflow);
              const currentMap = resolveWorkflowCompletedMap(
                existing,
                ubdWorkflowSteps,
                document.type,
                document.status || existing.currentStatus,
              );
              const inferred = inferUbdWorkflowCompletedFromArtifacts(document);
              const merged = Object.fromEntries(
                ubdWorkflowSteps.map((step) => [
                  step.key,
                  currentMap[step.key] || inferred[step.key] === true,
                ]),
              ) as Record<string, boolean>;
              const hasNewSteps = ubdWorkflowSteps.some(
                (step) => merged[step.key] && !currentMap[step.key],
              );
              if (!hasNewSteps) {
                return "skipped" as const;
              }
              const nextWorkflow = buildWorkflowFromCompletedMap(
                existing,
                ubdWorkflowSteps,
                merged,
              );
              const saved = await api.updatePersonDocument(
                document.personExternalId,
                document.id,
                {
                  title: document.title,
                  status: nextWorkflow.currentStatus,
                  workflow: nextWorkflow as unknown as Record<string, unknown>,
                },
                { suppressErrorToast: true },
              );
              applyUpdatedDocument(saved);
              if (document.id === selectedDocumentId) {
                setWorkflow(nextWorkflow);
              }
              return "updated" as const;
            } catch {
              return "fail" as const;
            }
          }),
        );
        updated += results.filter((result) => result === "updated").length;
        skipped += results.filter((result) => result === "skipped").length;
        fail += results.filter((result) => result === "fail").length;
      }
      setDocumentMessage(
        updated > 0
          ? `Частково відновлено прогрес (з файлів): ${updated}, без змін ${skipped}${
              fail > 0 ? `, помилок ${fail}` : ""
            }.`
          : "Не знайдено додаткових кроків для відновлення з файлів.",
      );
    } catch (error) {
      setDocumentMessage(
        error instanceof Error
          ? `Не вдалося відновити прогрес з файлів: ${error.message}`
          : "Не вдалося відновити прогрес з файлів.",
      );
    } finally {
      setIsApplyingBulkJournalProgress(false);
    }
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
    if (!documentSavePersonId || !selectedDocumentId) {
      saveWorkflowForPerson(workflowKey, nextWorkflow);
      return;
    }

    setIsSavingDocument(true);
    try {
      const fieldPayload = withPersistedDocumentFlags(nextFields);
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
        overwriteFields: true,
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
    if (!documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = withPersistedDocumentFlags(nextFields);
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
        overwriteFields: true,
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
    if (!documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = withPersistedDocumentFlags(nextFields);
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
        overwriteFields: true,
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
    if (!documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = withPersistedDocumentFlags(nextFields);
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
    if (!documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = withPersistedDocumentFlags(nextFields);
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
    if (!documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = withPersistedDocumentFlags(nextFields);
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
        { suppressErrorToast: true },
      );
      applyUpdatedDocument(updated);
      setDocumentMessage(
        `Збережено в БД · ${new Date(updated.updatedAt).toLocaleString("uk-UA")}`,
      );
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Не вдалося зберегти Форму 12.";
      setDocumentMessage(
        /import row not found/i.test(message)
          ? "Форма 12 змінена локально. У БД немає рядка ЕЖООС для цієї особи — імпортуйте журнал або збережіть документ після синхронізації."
          : `Не вдалося зберегти Форму 12: ${message}`,
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
    if (!documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = withPersistedDocumentFlags(nextFields);
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
    if (!documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = withPersistedDocumentFlags(nextFields);
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

  const saveLostMilitaryIdDocument = async (
    nextFields: LostMilitaryIdFields,
    nextFiles = documentFiles,
    nextWorkflow = workflow,
  ) => {
    if (!documentSavePersonId || !selectedDocumentId) return;

    setIsSavingDocument(true);
    try {
      const fieldPayload = withPersistedDocumentFlags(nextFields);
      const updated = await api.updatePersonDocument(
        documentSavePersonId,
        selectedDocumentId,
        {
          title: selectedDocument?.title || "Втрата військового квитка",
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
          ? `Не вдалося зберегти документ: ${error.message}`
          : "Не вдалося зберегти документ про втрату військового квитка.",
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
      const next = {
        ...current,
        [key]: key === "staffPosition" ? capitalizeReportPosition(value) : value,
      } as UbdReportFields;
      if (key === "taskPeriod" || key === "taskPlace") {
        const resolved = resolveUbdBasisForTask(next.taskPeriod, next.taskPlace);
        if (resolved) {
          next.basisNumber = resolved.number;
          next.basisDate = resolved.date;
          next.basis = formatUbdBasisText(resolved.number, resolved.date);
        }
        next.basisNotReady = !ubdBasisDateMatchesTaskPeriod(
          next.taskPeriod,
          next.basisDate,
          next.taskPlace,
        );
      }
      scheduleDocumentFieldSave(() => {
        void saveUbdDocument(next);
      });
      return next;
    });
  };

  const updateUbdBasisOrder = (optionKey: string) => {
    const option = findUbdBasisOrderByKey(optionKey);
    if (!option) return;
    setUbdFields((current) => {
      const next = {
        ...current,
        basisNumber: option.number,
        basisDate: option.date,
        basis: formatUbdBasisText(option.number, option.date),
        basisNotReady: !ubdBasisDateMatchesTaskPeriod(
          current.taskPeriod,
          option.date,
          current.taskPlace,
        ),
      };
      scheduleDocumentFieldSave(() => {
        void saveUbdDocument(next);
      });
      return next;
    });
  };

  const updateUbdBasisNotReady = (notReady: boolean) => {
    setUbdFields((current) => {
      const next = { ...current, basisNotReady: notReady };
      scheduleDocumentFieldSave(() => {
        void saveUbdDocument(next);
      });
      return next;
    });
  };

  const updateForm6Field = (
    key: Exclude<keyof Form6ReportFields, "signatories" | "basisManual">,
    value: string,
  ) => {
    setDocumentMessage("Є незбережені зміни у Формі 6.");
    setForm6Fields((current) => {
      const next = {
        ...current,
        [key]: key === "staffPosition" ? capitalizeReportPosition(value) : value,
      };
      if (
        (key === "taskPeriod" || key === "taskPlace") &&
        !current.basisManual
      ) {
        const resolved = resolveUbdBasisForTask(next.taskPeriod, next.taskPlace);
        if (resolved) {
          next.basisNumber = resolved.number;
          next.basisDate = resolved.date;
          next.basis = formatForm6BasisText(resolved.number, resolved.date);
        }
      }
      if (key === "basis") {
        next.basis = /^\s*Підстава:/i.test(value)
          ? stripForm6BasisLabel(value)
          : value;
        const extracted = extractForm6BasisParts(next.basis);
        if (extracted.basisNumber) next.basisNumber = extracted.basisNumber;
        if (extracted.basisDate) next.basisDate = extracted.basisDate;
      } else if (
        (key === "basisNumber" || key === "basisDate") &&
        !current.basisManual
      ) {
        next.basis = formatForm6BasisText(next.basisNumber, next.basisDate);
      } else if (key === "basisNumber" && current.basisManual) {
        const typed = stripForm6BasisLabel(value);
        if (/розпорядження/i.test(typed)) {
          next.basis = typed;
          const extracted = extractForm6BasisParts(typed);
          if (extracted.basisNumber) next.basisNumber = extracted.basisNumber;
          if (extracted.basisDate) next.basisDate = extracted.basisDate;
        }
      }
      return next;
    });
  };

  const updateForm6BasisManual = (manual: boolean) => {
    setDocumentMessage("Є незбережені зміни у Формі 6.");
    setForm6Fields((current) => {
      const next = { ...current, basisManual: manual };
      if (!manual) {
        const resolved = resolveUbdBasisForTask(
          current.taskPeriod,
          current.taskPlace,
        );
        if (resolved) {
          next.basisNumber = resolved.number;
          next.basisDate = resolved.date;
          next.basis = formatForm6BasisText(resolved.number, resolved.date);
        }
      }
      return next;
    });
  };

  const updateForm6BasisOrder = (optionKey: string) => {
    const option = findUbdBasisOrderByKey(optionKey);
    if (!option) return;
    setDocumentMessage("Є незбережені зміни у Формі 6.");
    setForm6Fields((current) => {
      const next = {
        ...current,
        basisNumber: option.number,
        basisDate: option.date,
        basis: formatForm6BasisText(option.number, option.date),
      };
      return next;
    });
  };


  const updateServiceCharacteristicField = (
    key: Exclude<keyof ServiceCharacteristicFields, "signatories">,
    value: string,
  ) => {
    setServiceCharacteristicFields((current) => {
      const next = {
        ...current,
        [key]: key === "staffPosition" ? capitalizeReportPosition(value) : value,
      };
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
      const next = {
        ...current,
        [key]: key === "staffPosition" ? capitalizeReportPosition(value) : value,
      };
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
      const next = {
        ...current,
        [key]: key === "staffPosition" ? capitalizeReportPosition(value) : value,
      };
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
      const next = {
        ...current,
        [key]: key === "staffPosition" ? capitalizeReportPosition(value) : value,
      };
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

  const updateLostMilitaryIdField = <K extends keyof LostMilitaryIdFields>(
    key: K,
    value: LostMilitaryIdFields[K],
  ) => {
    setLostMilitaryIdFields((current) => {
      const next = {
        ...current,
        [key]:
          (key === "staffPosition" || key === "investigatorPosition") &&
          typeof value === "string"
            ? capitalizeReportPosition(value)
            : value,
      };
      scheduleDocumentFieldSave(() => {
        void saveLostMilitaryIdDocument(next);
      });
      return next;
    });
  };

  const applyLostMilitaryIdPatch = (
    patch: Partial<LostMilitaryIdFields>,
  ) => {
    setLostMilitaryIdFields((current) => {
      const next = {
        ...current,
        ...patch,
        ...(typeof patch.investigatorPosition === "string"
          ? {
              investigatorPosition: capitalizeReportPosition(
                patch.investigatorPosition,
              ),
            }
          : {}),
      };
      scheduleDocumentFieldSave(() => {
        void saveLostMilitaryIdDocument(next);
      });
      return next;
    });
  };

  const addTicketPhoto = async (files: FileList | null) => {
    const file = files?.[0];
    if (!file || !selectedDocument) return;

    setIsSavingDocument(true);
    try {
      const dataUrl = await compressPhotoFile(file);
      const ticketPhoto: UbdScanFile = {
        id: `photo-${Date.now()}`,
        name: file.name.replace(/\.[^.]+$/, "") + ".jpg",
        type: "image/jpeg",
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
            fileName: ticketPhoto.name,
            mimeType: "image/jpeg",
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
      personQuestionnaire?.personExternalId ||
        selectedDocument?.personExternalId ||
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
      personQuestionnaire?.personExternalId ||
        selectedDocument?.personExternalId ||
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
      const url = await api.createPersonQuestionnairePreviewUrl(
        targetId,
        personQuestionnaire.fileName || questionnaireFileName,
      );
      window.open(url, "_blank", "noopener,noreferrer");
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
      | TemporaryMilitaryIdFields
      | LostMilitaryIdFields,
    nextWorkflow: SalaryWorkflowState,
    nextFiles = documentFiles,
  ) => {
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

    if (
      selectedDocument?.type === "lostMilitaryId" ||
      mode === "lostMilitaryId"
    ) {
      void saveLostMilitaryIdDocument(
        nextFields as LostMilitaryIdFields,
        nextFiles,
        nextWorkflow,
      );
      return;
    }

    void saveSalaryDocument(nextFields as SalaryDocumentFields, nextWorkflow);
  };

  const updateWorkflow = (next: SalaryWorkflowState) => {
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
            : mode === "lostMilitaryId"
            ? lostMilitaryIdFields
            : salaryFields,
      next,
    );
  };

  const setWorkflowStep = (
    stepKey: string,
    action: "toggle" | "on" = "toggle",
  ) => {
    let targetKey = stepKey;
    if (!activeWorkflowSteps.some((step) => step.key === targetKey)) {
      targetKey = resolveDocumentWorkflowStatus(
        selectedDocument?.type,
        stepKey,
      );
    }
    if (!activeWorkflowSteps.some((step) => step.key === targetKey)) return;

    // База — те, що зараз видно як виконане (і явні галочки, і старий
    // послідовний прогрес до currentStatus). Інакше клік по пізньому кроку
    // обнуляє перші, бо в `completed` їх ще не було.
    const baseMap = {
      ...resolveWorkflowCompletedMap(
        workflow,
        activeWorkflowSteps,
        selectedDocument?.type || mode,
        selectedDocument?.status || workflow.currentStatus,
      ),
    };

    const nextDone = action === "on" ? true : !baseMap[targetKey];
    const completed = {
      ...baseMap,
      [targetKey]: nextDone,
    };

    const stillDone = activeWorkflowSteps.filter((step) => completed[step.key]);
    const currentStatus = nextDone
      ? targetKey
      : stillDone[stillDone.length - 1]?.key ||
        activeWorkflowSteps[0]?.key ||
        targetKey;

    updateWorkflow({
      ...workflow,
      completed,
      independentSteps: true,
      currentStatus,
    });
    if (targetKey === "fighterSign") {
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
    setWorkflowStep("document", "on");
  };

  const printSalaryDocument = () => {
    const printWindow = window.open("", "_blank", "noopener,noreferrer");
    if (!printWindow) return;
    printWindow.document.write(salaryDocumentHtml(salaryFields));
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
    setWorkflowStep("print", "on");
  };

  const saveUbdAsWord = async () => {
    if (ubdFields.basisNotReady) {
      setDocumentMessage(
        "БР ще не підходить — Word недоступний, доки не оберете розпорядження з датою початку періоду завдань або не знімете позначку.",
      );
      return;
    }
    setDocumentMessage("Формую Word-документ...");
    const blob = await createUbdWordBlob(ubdWordFields);
    const fileName = `${safeFilePart(ubdFields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(
      `Word-файл «${fileName}» збережено.`,
    );
    setWorkflowStep("document", "on");
    void saveUbdDocument(ubdFields);
  };

  const saveForm6AsWord = async () => {
    setDocumentMessage("Формую Word-документ Форми 6...");
    const blob = await createForm6WordBlob(form6Fields);
    const fileName = `${safeFilePart(form6Fields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(`Word-файл «${fileName}» збережено.`);
    setWorkflowStep("document", "on");
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
    setWorkflowStep("document", "on");
    void saveServiceCharacteristicDocument(serviceCharacteristicFields);
  };

  const saveZhbdCertificateAsWord = async () => {
    setDocumentMessage("Формую Word-документ довідки ЖБД...");
    const blob = await createZhbdCertificateWordBlob(zhbdCertificateFields);
    const fileName = `${safeFilePart(zhbdCertificateFields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(`Word-файл «${fileName}» збережено.`);
    setWorkflowStep("document", "on");
    void saveZhbdCertificateDocument(zhbdCertificateFields);
  };

  const saveForm12AsWord = async () => {
    setDocumentMessage("Формую Word-документ Форми 12...");
    const blob = await createForm12WordBlob(form12Fields);
    const fileName = `${safeFilePart(form12Fields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(`Word-файл «${fileName}» збережено.`);
    setWorkflowStep("document", "on");
    void saveForm12Document(form12Fields);
  };

  const saveUbdRestoreAsWord = async () => {
    setDocumentMessage("Формую Word-документ рапорта на відновлення УБД...");
    const blob = await createUbdRestoreWordBlob(ubdRestoreFields);
    const fileName = `${safeFilePart(ubdRestoreFields.folderName)}.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(`Word-файл «${fileName}» збережено.`);
    setWorkflowStep("document", "on");
    void saveUbdRestoreDocument(ubdRestoreFields);
  };

  const saveLostMilitaryIdAsWord = async () => {
    setDocumentMessage("Формую рапорт про втрату військового квитка...");
    const blob = await createLostMilitaryIdReportWordBlob(lostMilitaryIdFields);
    const fileName = `${safeFilePart(lostMilitaryIdFields.folderName)} · Рапорт.docx`;
    downloadBlob(fileName, blob);
    setDocumentMessage(`Word-файл «${fileName}» збережено.`);
    setWorkflowStep("document", "on");
    void saveLostMilitaryIdDocument(lostMilitaryIdFields);
  };

  const saveLostMilitaryIdKit = async () => {
    setDocumentMessage("Формую комплект: рапорт, наказ, акт...");
    const blob = await createLostMilitaryIdKitZip(lostMilitaryIdFields);
    const fileName = `${safeFilePart(lostMilitaryIdFields.folderName)}.zip`;
    downloadBlob(fileName, blob);
    setDocumentMessage(
      `Комплект «${fileName}» збережено (рапорт, наказ, акт — 3 файли в архіві).`,
    );
    setWorkflowStep("document", "on");
    void saveLostMilitaryIdDocument(lostMilitaryIdFields);
  };

  const printUbdDocument = () => {
    if (ubdFields.basisNotReady) {
      setDocumentMessage(
        "БР ще не підходить — друк/PDF недоступні, доки не оберете розпорядження з датою початку періоду завдань або не знімете позначку.",
      );
      return;
    }
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
          const isDone = Boolean(workflowCompletedMap[step.key]);
          const isCurrent =
            step.key ===
            resolveDocumentWorkflowStatus(
              selectedDocument?.type || mode,
              workflow.currentStatus,
            );
          return (
            <button
              className={[
                "salary-step-node",
                isDone ? "done" : "",
                isCurrent ? "current" : "",
              ]
                .filter(Boolean)
                .join(" ")}
              key={step.key}
              type="button"
              onClick={() => setWorkflowStep(step.key)}
              title={
                isDone
                  ? `${step.title} · клікніть, щоб зняти`
                  : `${step.title} · клікніть, щоб позначити`
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
    if (type === "lostMilitaryId") return createLostMilitaryIdDocument();
  };

  const personDocumentCreateSelect = (
    <TextField
      select
      size="small"
      label="Створити документ"
      value="__none"
      disabled={isSavingDocument || !personExternalId}
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

  const updateSkippedDueToSzch = (next: boolean) => {
    skippedDueToSzchRef.current = next;
    setSkippedDueToSzch(next);
    if (selectedDocumentId) {
      const patchDocument = (document: BackendPersonDocument) =>
        document.id !== selectedDocumentId
          ? document
          : {
              ...document,
              fields: {
                ...(document.fields && typeof document.fields === "object"
                  ? document.fields
                  : {}),
                skippedDueToSzch: next,
              },
            };
      setPersonDocuments((current) => current.map(patchDocument));
      setAllPersonDocuments((current) => current.map(patchDocument));
    }
    const activeFields =
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
                    : mode === "lostMilitaryId"
                      ? lostMilitaryIdFields
                      : salaryFields;
    scheduleDocumentFieldSave(() => {
      saveActiveDocumentWorkflow(activeFields, workflow);
    });
  };

  const personStatusLooksSzch = /СЗЧ|САМОВІЛ/i.test(
    selectedDocument
      ? readDocumentPersonStatus(selectedDocument, personStatusById)
      : "",
  );

  const documentStatusNotePanel = (
    value: string,
    onChange: (value: string) => void,
  ) => (
    <div className="document-status-note-panel">
      <div className="panel-heading">Коментар до статусу</div>
      <div className="document-szch-skip">
        <Checkbox
          checked={skippedDueToSzch}
          onCheckedChange={(checked) =>
            updateSkippedDueToSzch(checked === true)
          }
          label="Не потрібно робити: СЗЧ"
        />
      </div>
      {personStatusLooksSzch && !skippedDueToSzch ? (
        <span className="document-field-hint">
          У статусі службовця є СЗЧ — можна зняти документ з роботи.
        </span>
      ) : null}
      {skippedDueToSzch ? (
        <span className="document-field-hint">
          Документ сірий у журналі і не входить в експорт.
        </span>
      ) : null}
      <TextField
        size="small"
        fullWidth
        multiline
        rows={3}
        value={value}
        onChange={(event) => onChange(event.target.value)}
        placeholder="Проблема з документами, уточнення по статусу або що треба доробити"
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
    if (mode === "lostMilitaryId") return lostMilitaryIdFields.statusNote.trim();
    if (mode === "salaryPowerAttorney") return salaryFields.statusNote.trim();
    return readDocumentStatusNote(document);
  };

  const documentListStatusNote = (document: BackendPersonDocument) => {
    const skipped = readDocumentSkippedDueToSzch(liveJournalFields(document));
    const note = liveDocumentStatusNote(document);
    if (!skipped && !note) return null;
    const text = skipped
      ? note
        ? `СЗЧ · ${note}`
        : "Не потрібно: СЗЧ"
      : note;
    return (
      <small className="document-status-note-preview" title={text}>
        {text}
      </small>
    );
  };

  const personDocumentShellClass = (document: BackendPersonDocument) =>
    [
      "salary-person-document-shell",
      document.id === selectedDocumentId ? "active" : "",
      readDocumentSkippedDueToSzch(liveJournalFields(document))
        ? "is-skipped-szch"
        : "",
    ]
      .filter(Boolean)
      .join(" ");

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
      <div className="panel-heading">
        {mode === "lostMilitaryId"
          ? `Попередній перегляд Word · ${lostMilitaryIdPreviewLabel}`
          : "Попередній перегляд Word"}
      </div>
      {mode === "lostMilitaryId" ? (
        <>
          <Stack direction="row" spacing={1} sx={{ mb: 1 }}>
            {(
              [
                ["report", "Рапорт"],
                ["order", "Наказ"],
                ["act", "Акт"],
              ] as const
            ).map(([key, label]) => (
              <Button
                key={key}
                size="small"
                variant={lostMilitaryIdPreviewDoc === key ? "contained" : "outlined"}
                onClick={() => setLostMilitaryIdPreviewDoc(key)}
              >
                {label}
              </Button>
            ))}
          </Stack>
          <Typography variant="caption" color="text.secondary" sx={{ display: "block", mb: 1 }}>
            «Сформувати комплект» завантажує три окремі .docx у ZIP. Новий формат акту — з
            «ЗАТВЕРДЖУЮ» зверху; відкрийте файл «… · Акт.docx».
          </Typography>
        </>
      ) : null}
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
                ].map(([key, label]) => {
                  const isRnokppInvalid =
                    key === "rnokpp" && isBlankUbdRnokpp(ubdFields.rnokpp);
                  const isInvalid = documentRequiredFieldIsBlank(
                    "ubdReport",
                    key,
                    ubdFields as unknown as Record<string, unknown>,
                  );
                  const field = (
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
                  );
                  return (
                    <label
                      key={key}
                      className={documentFieldLabelClass(
                        isInvalid && "document-field-invalid",
                      )}
                    >
                      <code>{label}</code>
                      {isRnokppInvalid ? (
                        <div className="document-field-with-hint">
                          {field}
                          <span className="document-field-hint document-field-hint-error">
                            Потрібно рівно 10 цифр РНОКПП
                            {String(ubdFields.rnokpp ?? "").replace(/\D/g, "")
                              .length
                              ? ` · зараз ${String(ubdFields.rnokpp ?? "").replace(/\D/g, "").length}`
                              : ""}
                          </span>
                        </div>
                      ) : (
                        field
                      )}
                    </label>
                  );
                })}
                <label
                  className={documentBasisFieldHighlightClass(
                    "ubdReport",
                    ubdFields as unknown as Record<string, unknown>,
                    "basisNumber",
                  )}
                >
                  <code>№ розпорядження</code>
                  <TextField
                    select
                    size="small"
                    fullWidth
                    value={ubdBasisOrderOptionKey({
                      number: ubdFields.basisNumber || FALLBACK_UBD_BASIS.number,
                      date: ubdFields.basisDate || FALLBACK_UBD_BASIS.date,
                    })}
                    onChange={(event) =>
                      updateUbdBasisOrder(event.target.value)
                    }
                  >
                    {!allBasisOrderOptions().some(
                      (item) =>
                        item.number === ubdFields.basisNumber &&
                        item.date === ubdFields.basisDate,
                    ) && ubdFields.basisNumber ? (
                      <MenuItem
                        value={ubdBasisOrderOptionKey({
                          number: ubdFields.basisNumber,
                          date: ubdFields.basisDate,
                        })}
                      >
                        {formatUbdBasisOrderLabel({
                          number: ubdFields.basisNumber,
                          date: ubdFields.basisDate,
                        })}{" "}
                        · збережене
                      </MenuItem>
                    ) : null}
                    {allBasisOrderOptions().map((option) => (
                      <MenuItem
                        key={ubdBasisOrderOptionKey(option)}
                        value={ubdBasisOrderOptionKey(option)}
                      >
                        {formatUbdBasisOrderLabel(option)}
                      </MenuItem>
                    ))}
                  </TextField>
                </label>
                <label
                  className={documentBasisFieldHighlightClass(
                    "ubdReport",
                    ubdFields as unknown as Record<string, unknown>,
                    "basisDate",
                  )}
                >
                  <code>Дата розпорядження</code>
                  <TextField
                    size="small"
                    fullWidth
                    value={ubdFields.basisDate}
                    readOnly
                  />
                </label>
                <div className="wide ubd-basis-not-ready">
                  <Checkbox
                    checked={ubdFields.basisNotReady}
                    onCheckedChange={(checked) =>
                      updateUbdBasisNotReady(checked === true)
                    }
                    label="БР ще не підходить"
                  />
                  {ubdFields.basisNotReady ||
                  !ubdBasisDateMatchesTaskPeriod(
                    ubdFields.taskPeriod,
                    ubdFields.basisDate,
                    ubdFields.taskPlace,
                  ) ? (
                    <Alert severity="warning" className="ubd-basis-not-ready-alert">
                      {ubdFields.basisNotReady
                        ? "Рапорт позначено як незавершений: Word і PDF заблоковані, доки не оберете потрібне розпорядження або не знімете позначку."
                        : `Дата БР (${ubdFields.basisDate || "—"}) не збігається з початком періоду завдань${
                            ubdHasExactBasisForTaskPeriod(
                              ubdFields.taskPeriod,
                              ubdFields.taskPlace,
                            )
                              ? ""
                              : " — у списках ще немає БР на цю дату"
                          }. Поставте «БР ще не підходить», якщо рапорт поки не можна сформувати.`}
                    </Alert>
                  ) : null}
                </div>
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
          <div className="document-form-save-row">
            <span>
              Зміни зберігаються лише після натискання кнопки.
            </span>
            <Button
              variant="contained"
              disabled={isSavingDocument}
              onClick={() => void saveForm6Document(form6Fields)}
            >
              {isSavingDocument ? "Збереження…" : "Зберегти дані"}
            </Button>
          </div>
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
              ] as Array<
                [
                  Exclude<
                    keyof Form6ReportFields,
                    "signatories" | "basisManual"
                  >,
                  string,
                ]
              >
            ).map(([key, label]) => (
              <label
                key={key}
                className={documentFieldLabelClass(
                  (key === "staffPosition" ||
                    key === "address" ||
                    key === "commander") &&
                    "wide",
                  documentRequiredFieldIsBlank(
                    "form6Report",
                    key,
                    form6Fields as unknown as Record<string, unknown>,
                  ) && "document-field-invalid",
                )}
              >
                <code>{label}</code>
                <div className="document-field-with-hint">
                  <TextField
                    size="small"
                    fullWidth
                    multiline={
                      key === "staffPosition" ||
                      key === "address" ||
                      key === "commander"
                    }
                    rows={
                      key === "staffPosition" ||
                      key === "address" ||
                      key === "commander"
                        ? 2
                        : undefined
                    }
                    value={form6Fields[key]}
                    onChange={(event) =>
                      updateForm6Field(key, event.target.value)
                    }
                  />
                  {key === "idDocument" &&
                  isBlankForm6IdDocument(form6Fields.idDocument) ? (
                    <span className="document-field-hint document-field-hint-error">
                      Потрібно серію та номер паспорта, не лише назву документа
                    </span>
                  ) : null}
                  {key === "rnokpp" && isBlankUbdRnokpp(form6Fields.rnokpp) ? (
                    <span className="document-field-hint document-field-hint-error">
                      Потрібно рівно 10 цифр РНОКПП
                      {String(form6Fields.rnokpp ?? "").replace(/\D/g, "")
                        .length
                        ? ` · зараз ${String(form6Fields.rnokpp ?? "").replace(/\D/g, "").length}`
                        : ""}
                    </span>
                  ) : null}
                </div>
              </label>
            ))}
            <div className="wide ubd-basis-not-ready">
              <Checkbox
                checked={form6Fields.basisManual}
                onCheckedChange={(checked) =>
                  updateForm6BasisManual(checked === true)
                }
                label="Ввести БР вручну"
              />
            </div>
            <label
              className={documentBasisFieldHighlightClass(
                "form6Report",
                form6Fields as unknown as Record<string, unknown>,
                "basisNumber",
              )}
            >
              <code>№ розпорядження</code>
              {form6Fields.basisManual ? (
                <TextField
                  size="small"
                  fullWidth
                  value={form6Fields.basisNumber}
                  onChange={(event) =>
                    updateForm6Field("basisNumber", event.target.value)
                  }
                  placeholder="4862/ОКП/…/дск"
                />
              ) : (
                <TextField
                  select
                  size="small"
                  fullWidth
                  value={ubdBasisOrderOptionKey({
                    number: form6Fields.basisNumber || FALLBACK_UBD_BASIS.number,
                    date: form6Fields.basisDate || FALLBACK_UBD_BASIS.date,
                  })}
                  onChange={(event) =>
                    updateForm6BasisOrder(event.target.value)
                  }
                >
                  {!allBasisOrderOptions().some(
                    (item) =>
                      item.number === form6Fields.basisNumber &&
                      item.date === form6Fields.basisDate,
                  ) && form6Fields.basisNumber ? (
                    <MenuItem
                      value={ubdBasisOrderOptionKey({
                        number: form6Fields.basisNumber,
                        date: form6Fields.basisDate,
                      })}
                    >
                      {formatUbdBasisOrderLabel({
                        number: form6Fields.basisNumber,
                        date: form6Fields.basisDate,
                      })}{" "}
                      · збережене
                    </MenuItem>
                  ) : null}
                  {allBasisOrderOptions().map((option) => (
                    <MenuItem
                      key={`form6-${ubdBasisOrderOptionKey(option)}`}
                      value={ubdBasisOrderOptionKey(option)}
                    >
                      {formatUbdBasisOrderLabel(option)}
                    </MenuItem>
                  ))}
                </TextField>
              )}
            </label>
            <label
              className={documentBasisFieldHighlightClass(
                "form6Report",
                form6Fields as unknown as Record<string, unknown>,
                "basisDate",
              )}
            >
              <code>Дата розпорядження</code>
              <TextField
                size="small"
                fullWidth
                value={form6Fields.basisDate}
                readOnly={!form6Fields.basisManual}
                onChange={
                  form6Fields.basisManual
                    ? (event) =>
                        updateForm6Field("basisDate", event.target.value)
                    : undefined
                }
                placeholder="дд.мм.рррр"
              />
            </label>
            {!ubdBasisDateMatchesTaskPeriod(
              form6Fields.taskPeriod,
              form6Fields.basisDate,
              form6Fields.taskPlace,
            ) ? (
              <Alert severity="warning" className="ubd-basis-not-ready-alert">
                {form6Fields.basisManual
                  ? `Дата БР (${form6Fields.basisDate || "—"}) не збігається з початком періоду завдань.`
                  : `Дата БР (${form6Fields.basisDate || "—"}) не збігається з початком періоду завдань${
                      ubdHasExactBasisForTaskPeriod(
                        form6Fields.taskPeriod,
                        form6Fields.taskPlace,
                      )
                        ? ""
                        : " — у списках ще немає БР на цю дату"
                    }. Поставте «Ввести БР вручну», якщо потрібне інше розпорядження.`}
              </Alert>
            ) : null}
            <label className="wide">
              <code>Підстава</code>
              <div className="document-field-with-hint">
                <TextField
                  size="small"
                  fullWidth
                  multiline
                  rows={2}
                  value={form6Fields.basis}
                  readOnly={!form6Fields.basisManual}
                  onChange={
                    form6Fields.basisManual
                      ? (event) =>
                          updateForm6Field("basis", event.target.value)
                      : undefined
                  }
                  placeholder="Бойове розпорядження командира … №… від …"
                />
                {form6Fields.basisManual ? (
                  <span className="document-field-hint">
                    У Word вже є «Підстава:» — введіть лише текст БР після
                    цього слова.
                  </span>
                ) : null}
              </div>
            </label>
            {(
              [
                ["date", "Дата"],
                ["formPurpose", "Для чого форма"],
                ["folderName", "Папка"],
              ] as Array<
                [
                  Exclude<
                    keyof Form6ReportFields,
                    "signatories" | "basisManual"
                  >,
                  string,
                ]
              >
            ).map(([key, label]) => (
              <label
                key={key}
                className={documentFieldLabelClass(
                  key === "formPurpose" && "wide",
                )}
              >
                <code>{label}</code>
                <TextField
                  size="small"
                  fullWidth
                  multiline={key === "formPurpose"}
                  rows={key === "formPurpose" ? 2 : undefined}
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
                ["formPurpose", "Для чого форма"],
                ["date", "Дата"],
                ["folderName", "Папка"],
              ] as Array<
                [Exclude<keyof Form12ReportFields, "signatories" | "signatureData" | "signatureFileName" | "statusNote">, string]
              >
            ).map(([key, label]) => (
              <label
                key={key}
                className={documentFieldLabelClass(
                  (key === "staffPosition" ||
                    key === "commander" ||
                    key === "formPurpose") && "wide",
                  documentRequiredFieldIsBlank(
                    "form12Report",
                    key,
                    form12Fields as unknown as Record<string, unknown>,
                  ) && "document-field-invalid",
                )}
              >
                <code>{label}</code>
                <TextField
                  size="small"
                  fullWidth
                  multiline={
                    key === "staffPosition" ||
                    key === "commander" ||
                    key === "formPurpose"
                  }
                  rows={
                    key === "staffPosition" ||
                    key === "commander" ||
                    key === "formPurpose"
                      ? 2
                      : undefined
                  }
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
                className={documentFieldLabelClass(
                  (key === "staffPosition" ||
                    key.endsWith("Paragraph") ||
                    key === "conclusion") &&
                    "wide",
                  documentRequiredFieldIsBlank(
                    "serviceCharacteristic",
                    key,
                    serviceCharacteristicFields as unknown as Record<
                      string,
                      unknown
                    >,
                  ) && "document-field-invalid",
                )}
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
                className={documentFieldLabelClass(
                  (key === "staffPosition" ||
                    key === "bodyParagraph" ||
                    key === "basisOrders" ||
                    key === "basisJournal" ||
                    key === "periodNote") &&
                    "wide",
                  documentRequiredFieldIsBlank(
                    "zhbdCertificate",
                    key,
                    zhbdCertificateFields as unknown as Record<string, unknown>,
                  ) && "document-field-invalid",
                )}
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
                className={documentFieldLabelClass(
                  (key === "staffPosition" ||
                    key === "signerTitle" ||
                    key === "circumstances" ||
                    key === "requestText" ||
                    key === "commander" ||
                    key === "coveringCommander") &&
                    "wide",
                  documentRequiredFieldIsBlank(
                    "ubdRestoreReport",
                    key,
                    ubdRestoreFields as unknown as Record<string, unknown>,
                  ) && "document-field-invalid",
                )}
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
            <label
              className={documentFieldLabelClass(
                documentRequiredFieldIsBlank(
                  "temporaryMilitaryId",
                  "rank",
                  ticketFields as unknown as Record<string, unknown>,
                ) && "document-field-invalid",
              )}
            >
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
            <label
              className={documentFieldLabelClass(
                documentRequiredFieldIsBlank(
                  "temporaryMilitaryId",
                  "fullName",
                  ticketFields as unknown as Record<string, unknown>,
                ) && "document-field-invalid",
              )}
            >
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
            <label
              className={documentFieldLabelClass(
                documentRequiredFieldIsBlank(
                  "temporaryMilitaryId",
                  "birthDate",
                  ticketFields as unknown as Record<string, unknown>,
                ) && "document-field-invalid",
              )}
            >
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

  const lostMilitaryIdWorkspace = (
    <>
      {selectedDocument ? (
        <section className="analytics-panel document-fields">
          <div className="panel-heading">Втрата військового квитка</div>
          <div className="document-placeholder-map salary-document-map ubd-document-map">
            {(
              [
                ["fullName", "Військовослужбовець"],
                ["rank", "Звання"],
                ["staffPosition", "Посада"],
                ["unitLabel", "Підрозділ"],
                ["militaryUnit", "В/ч"],
                ["addressee", "Кому"],
                ["lossDate", "Дата втрати"],
                ["fromLocation", "Звідки"],
                ["toLocation", "Куди"],
                ["customCircumstances", "Інші обставини"],
                ["searchResult", "Результат пошуку"],
                ...(lostMilitaryIdFields.signatories.length
                  ? []
                  : ([
                      ["reporterFullName", "Хто подає рапорт"],
                      ["reporterRank", "Звання того, хто подає"],
                      ["reporterTitle", "Посада того, хто подає"],
                    ] as Array<[keyof LostMilitaryIdFields, string]>)),
              ] as Array<[keyof LostMilitaryIdFields, string]>
            ).map(([key, label]) => (
              <label
                key={key}
                className={documentFieldLabelClass(
                  (key === "staffPosition" ||
                    key === "addressee" ||
                    key === "reporterTitle" ||
                    key === "investigatorPosition" ||
                    key === "customCircumstances") &&
                    "wide",
                  documentRequiredFieldIsBlank(
                    "lostMilitaryId",
                    key,
                    lostMilitaryIdFields as unknown as Record<string, unknown>,
                  ) && "document-field-invalid",
                )}
              >
                <code>{label}</code>
                <TextField
                  size="small"
                  fullWidth
                  multiline={
                    key === "staffPosition" ||
                    key === "addressee" ||
                    key === "reporterTitle" ||
                    key === "investigatorPosition" ||
                    key === "customCircumstances"
                  }
                  rows={
                    key === "staffPosition" ||
                    key === "reporterTitle" ||
                    key === "investigatorPosition"
                      ? 2
                      : undefined
                  }
                  value={String(lostMilitaryIdFields[key] ?? "")}
                  onChange={(event) =>
                    updateLostMilitaryIdField(
                      key,
                      event.target.value as LostMilitaryIdFields[typeof key],
                    )
                  }
                />
              </label>
            ))}
            <div className="wide document-investigator-manual">
              <Checkbox
                checked={lostMilitaryIdFields.investigatorManual}
                onCheckedChange={(checked) =>
                  applyLostMilitaryIdPatch({
                    investigatorManual: Boolean(checked),
                  })
                }
                label="Ввести розслідувача вручну"
              />
            </div>
            <label className="wide">
              <code>Хто проводить розслідування</code>
              {lostMilitaryIdFields.investigatorManual ? (
                <TextField
                  size="small"
                  fullWidth
                  value={lostMilitaryIdFields.investigatorFullName}
                  onChange={(event) =>
                    applyLostMilitaryIdPatch({
                      investigatorFullName: event.target.value,
                      investigatorPersonId: "",
                    })
                  }
                />
              ) : (
                <PersonnelNamePicker
                  value={lostMilitaryIdFields.investigatorFullName}
                  onPick={(row) =>
                    applyLostMilitaryIdPatch(investigatorFromPersonnelRow(row))
                  }
                />
              )}
            </label>
            <label>
              <code>Звання розслідувача</code>
              <TextField
                size="small"
                fullWidth
                disabled={!lostMilitaryIdFields.investigatorManual}
                value={lostMilitaryIdFields.investigatorRank}
                onChange={(event) =>
                  updateLostMilitaryIdField("investigatorRank", event.target.value)
                }
              />
            </label>
            <label className="wide">
              <code>Посада розслідувача</code>
              <TextField
                size="small"
                fullWidth
                multiline
                rows={2}
                disabled={!lostMilitaryIdFields.investigatorManual}
                value={capitalizeReportPosition(
                  lostMilitaryIdFields.investigatorPosition,
                )}
                onChange={(event) =>
                  updateLostMilitaryIdField(
                    "investigatorPosition",
                    event.target.value,
                  )
                }
              />
            </label>
            {(
              [
                ["reportDate", "Дата рапорту"],
                ["orderNumber", "Номер наказу"],
                ["orderDate", "Дата наказу"],
                ["folderName", "Папка"],
              ] as Array<[keyof LostMilitaryIdFields, string]>
            ).map(([key, label]) => (
              <label key={key}>
                <code>{label}</code>
                <TextField
                  size="small"
                  fullWidth
                  value={String(lostMilitaryIdFields[key] ?? "")}
                  onChange={(event) =>
                    updateLostMilitaryIdField(
                      key,
                      event.target.value as LostMilitaryIdFields[typeof key],
                    )
                  }
                />
              </label>
            ))}
            {lostMilitaryIdFields.signatories.length ? (
              <div className="wide document-default-signatories">
                <code>Підпис рапорту (з шаблону)</code>
                {lostMilitaryIdFields.signatories
                  .filter((signatory) => signatory.blockType === "SIGNER")
                  .map((signatory, index) => {
                    const footer = reporterFooterBlock(lostMilitaryIdFields);
                    const signer = reportSignerOf(lostMilitaryIdFields);
                    return (
                      <article
                        key={`${signatory.fullName}-${index}`}
                      >
                        <strong>
                          {signer?.fullName || signatory.fullName}
                        </strong>
                        {footer.titleLines.map((line) => (
                          <span key={line}>{line}</span>
                        ))}
                        <small>
                          {footer.rank} · {footer.name}
                        </small>
                      </article>
                    );
                  })}
              </div>
            ) : null}
          </div>
          <Stack direction="row" spacing={2} sx={{ mt: 1, flexWrap: "wrap" }}>
            <label>
              <Checkbox
                checked={lostMilitaryIdFields.isExactDate}
                onCheckedChange={(checked) =>
                  updateLostMilitaryIdField("isExactDate", Boolean(checked))
                }
              />
              <span>Дата точна (інакше «орієнтовно»)</span>
            </label>
            <label>
              <Checkbox
                checked={lostMilitaryIdFields.circumstanceKind === "movement"}
                onCheckedChange={(checked) =>
                  updateLostMilitaryIdField(
                    "circumstanceKind",
                    checked ? "movement" : "custom",
                  )
                }
              />
              <span>Під час переміщення</span>
            </label>
            <label>
              <Checkbox
                checked={lostMilitaryIdFields.searchConducted}
                onCheckedChange={(checked) =>
                  updateLostMilitaryIdField("searchConducted", Boolean(checked))
                }
              />
              <span>Пошук проводився</span>
            </label>
            <label>
              <Checkbox
                checked={lostMilitaryIdFields.editManually}
                onCheckedChange={(checked) =>
                  updateLostMilitaryIdField("editManually", Boolean(checked))
                }
              />
              <span>Редагувати вручну</span>
            </label>
          </Stack>
          {lostMilitaryIdFields.editManually ? (
            <div className="document-placeholder-map salary-document-map ubd-document-map">
              <label className="wide">
                <code>ПІБ (орудний)</code>
                <TextField
                  size="small"
                  fullWidth
                  value={lostMilitaryIdFields.fullNameInstrumentalManual}
                  placeholder="ЛИСЕНКОМ Михайлом Юрійовичем"
                  onChange={(event) =>
                    updateLostMilitaryIdField(
                      "fullNameInstrumentalManual",
                      event.target.value,
                    )
                  }
                />
              </label>
              <label className="wide">
                <code>Розслідувач (давальний)</code>
                <TextField
                  size="small"
                  fullWidth
                  value={lostMilitaryIdFields.investigatorDativeManual}
                  placeholder="ГУНЬКУ Олександру Олександровичу"
                  onChange={(event) =>
                    updateLostMilitaryIdField(
                      "investigatorDativeManual",
                      event.target.value,
                    )
                  }
                />
              </label>
              <label className="wide">
                <code>Текст рапорту</code>
                <TextField
                  size="small"
                  fullWidth
                  multiline
                  rows={8}
                  value={
                    lostMilitaryIdFields.reportTextOverride ||
                    buildLostMilitaryIdReportText({
                      ...lostMilitaryIdFields,
                      reportTextOverride: "",
                    })
                  }
                  onChange={(event) =>
                    updateLostMilitaryIdField(
                      "reportTextOverride",
                      event.target.value,
                    )
                  }
                />
              </label>
              <label className="wide">
                <code>Текст наказу</code>
                <TextField
                  size="small"
                  fullWidth
                  multiline
                  rows={6}
                  value={
                    lostMilitaryIdFields.orderTextOverride ||
                    buildLostMilitaryIdOrderText({
                      ...lostMilitaryIdFields,
                      orderTextOverride: "",
                    })
                  }
                  onChange={(event) =>
                    updateLostMilitaryIdField(
                      "orderTextOverride",
                      event.target.value,
                    )
                  }
                />
              </label>
            </div>
          ) : (
            <section className="analytics-panel" style={{ marginTop: 12 }}>
              <div className="panel-heading">Текст рапорту</div>
              <Typography variant="body2" sx={{ whiteSpace: "pre-wrap" }}>
                {buildLostMilitaryIdReportText(lostMilitaryIdFields)}
              </Typography>
            </section>
          )}
          {documentStatusNotePanel(lostMilitaryIdFields.statusNote, (value) =>
            updateLostMilitaryIdField("statusNote", value),
          )}
          {personQuestionnaireRow}
        </section>
      ) : null}
      {wordPreviewPanel}
    </>
  );

  const salaryWorkspace = (
    <>
      {selectedDocument ? (
        <section className="analytics-panel document-fields">
          <div className="panel-heading">Дані для рапорта</div>
          <div className="document-placeholder-map salary-document-map">
            <label
              className={documentFieldLabelClass(
                documentRequiredFieldIsBlank(
                  "salaryPowerAttorney",
                  "fullName",
                  salaryFields as unknown as Record<string, unknown>,
                ) && "document-field-invalid",
              )}
            >
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
            <label
              className={documentFieldLabelClass(
                documentRequiredFieldIsBlank(
                  "salaryPowerAttorney",
                  "rnokpp",
                  salaryFields as unknown as Record<string, unknown>,
                ) && "document-field-invalid",
              )}
            >
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
            <label
              className={documentFieldLabelClass(
                documentRequiredFieldIsBlank(
                  "salaryPowerAttorney",
                  "iban",
                  salaryFields as unknown as Record<string, unknown>,
                ) && "document-field-invalid",
              )}
            >
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
            <label
              className={documentFieldLabelClass(
                documentRequiredFieldIsBlank(
                  "salaryPowerAttorney",
                  "bankName",
                  salaryFields as unknown as Record<string, unknown>,
                ) && "document-field-invalid",
              )}
            >
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
    (mode === "lostMilitaryId" &&
      selectedDocument?.type === "lostMilitaryId") ||
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
                !journalExportSelectedCount
              }
              onClick={() => void exportDocumentJournal()}
              sx={{ color: "#1a1a14" }}
            >
              {isExportingDocumentJournal
                ? "Експорт..."
                : journalTypeFilter === "ubdReport"
                  ? `Excel УБД · ${journalExportSelectedCount}`
                  : `Excel · ${journalExportSelectedCount}`}
            </Button>
            {journalTypeFilter === "ubdReport" ? (
              <Button
                variant="outlined"
                startIcon={<ArticleOutlinedIcon />}
                disabled={
                  isLoadingDocumentJournal ||
                  isExportingDocumentJournal ||
                  !journalExportSelectedCount
                }
                onClick={() => void exportUbdBulkWordFromJournal()}
              >
                Word УБД · {journalExportSelectedCount}
              </Button>
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

        <section className="analytics-panel documents-journal-panel">
          <div className="panel-heading">
            Усі документи · {filteredJournalDocuments.length}
            {filteredJournalDocuments.length !== journalDocuments.length
              ? ` / ${journalDocuments.length}`
              : ""}
            {filteredJournalDocuments.length
              ? ` · експорт ${journalExportSelectedCount}`
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
                const allowed = new Set(nextSteps.map((step) => step.title));
                setJournalStatusFilters((current) =>
                  current.filter((title) => allowed.has(title)),
                );
              }}
            >
              {JOURNAL_DOCUMENT_TYPE_FILTERS.map((item) => (
                <MenuItem key={item.value} value={item.value}>
                  {item.label}
                </MenuItem>
              ))}
            </TextField>
            <div
              className="sci-field documents-journal-status-filter"
              ref={journalStatusMenuRef}
            >
              <span>Статус</span>
              <button
                type="button"
                className="documents-journal-status-trigger"
                aria-expanded={journalStatusMenuOpen}
                aria-haspopup="listbox"
                onClick={() => setJournalStatusMenuOpen((open) => !open)}
              >
                <span>{journalStatusFilterLabel}</span>
                <span aria-hidden="true">▾</span>
              </button>
              {journalStatusMenuOpen ? (
                <div
                  className="documents-journal-status-menu"
                  role="listbox"
                  aria-multiselectable="true"
                >
                  <Checkbox
                    checked={journalStatusFilters.length === 0}
                    onCheckedChange={(checked) => {
                      if (checked === true) setJournalStatusFilters([]);
                    }}
                    label="Усі статуси"
                  />
                  {journalStatusOptions.map((title) => (
                    <Checkbox
                      key={title}
                      checked={journalStatusFilters.includes(title)}
                      onCheckedChange={() => toggleJournalStatusFilter(title)}
                      label={title}
                    />
                  ))}
                </div>
              ) : null}
            </div>
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
            <TextField
              select
              size="small"
              label="Готовність"
              value={journalReadinessFilter}
              onChange={(event) =>
                setJournalReadinessFilter(
                  event.target.value as JournalReadinessFilter,
                )
              }
            >
              <MenuItem value="ALL">Усі</MenuItem>
              <MenuItem value="READY_TO_SEND">Готово до відправки</MenuItem>
              <MenuItem value="INCOMPLETE">Не готові до відправки</MenuItem>
              <MenuItem value="COMPLETE">Повний прогрес</MenuItem>
              <MenuItem value="SKIPPED_SZCH">Не потрібні: СЗЧ</MenuItem>
            </TextField>
          </div>
          {journalTypeFilter === "ubdReport" ? (
            <div className="documents-journal-bulk-progress documents-journal-bulk-progress-recovery">
              <div className="documents-journal-bulk-progress-heading">
                <button
                  type="button"
                  className="documents-journal-bulk-progress-toggle"
                  aria-expanded={ubdProgressBackupsOpen}
                  onClick={() =>
                    setUbdProgressBackupsOpen((current) => !current)
                  }
                >
                  <span
                    className={`documents-journal-bulk-progress-chevron${
                      ubdProgressBackupsOpen ? " is-open" : ""
                    }`}
                    aria-hidden
                  />
                  <code>
                    Резервні копії прогресу УБД · збережено в браузері ·{" "}
                    {ubdProgressBackups.length}
                  </code>
                </button>
                <Stack direction="row" spacing={1}>
                  <Button
                    variant="outlined"
                    disabled={isApplyingBulkJournalProgress}
                    onClick={saveManualUbdProgressBackup}
                  >
                    Зберегти копію зараз
                  </Button>
                  {journalExportSelectedCount > 0 ? (
                    <Button
                      variant="outlined"
                      disabled={isApplyingBulkJournalProgress}
                      onClick={() => void restoreUbdProgressFromArtifacts()}
                    >
                      Додати з файлів
                    </Button>
                  ) : null}
                </Stack>
              </div>
              {ubdProgressBackupsOpen ? (
                ubdProgressBackups.length ? (
                  <div className="documents-journal-bulk-progress-backups">
                    {ubdProgressBackups.map((backup) => (
                      <div
                        className="documents-journal-bulk-progress-backup-row"
                        key={backup.id}
                      >
                        <span>
                          {formatUbdProgressBackupWhen(backup.savedAt)} ·{" "}
                          {backup.label} · {backup.entries.length} рапортів
                        </span>
                        <Button
                          variant="outlined"
                          color="warning"
                          size="small"
                          disabled={isApplyingBulkJournalProgress}
                          onClick={() => void restoreUbdProgressBackup(backup)}
                        >
                          Відновити
                        </Button>
                      </div>
                    ))}
                  </div>
                ) : (
                  <code className="documents-journal-bulk-progress-empty">
                    Копій ще немає. Перед масовим прогресом копія створюється
                    автоматично; також можна зберегти вручну.
                  </code>
                )
              ) : null}
            </div>
          ) : null}
          {journalTypeFilter === "ubdReport" && journalBulkProgressUndo?.length ? (
            <div className="documents-journal-bulk-progress documents-journal-bulk-progress-undo">
              <div className="documents-journal-bulk-progress-heading">
                <code>
                  Остання масова зміна прогресу · {journalBulkProgressUndo.length}{" "}
                  рапортів · можна повернути попередній стан
                </code>
                <Button
                  variant="outlined"
                  color="warning"
                  disabled={isApplyingBulkJournalProgress}
                  onClick={() => void undoBulkJournalProgress()}
                >
                  {isApplyingBulkJournalProgress
                    ? "Скасовую..."
                    : "Скасувати останнє"}
                </Button>
              </div>
            </div>
          ) : null}
          {journalTypeFilter === "ubdReport" && journalExportSelectedCount > 0 ? (
            <div className="documents-journal-bulk-progress">
              <div className="documents-journal-bulk-progress-heading">
                <code>
                  Прогрес для обраних · {journalExportSelectedCount} · додасть
                  обрані кроки, якщо їх ще немає
                </code>
                <Stack direction="row" spacing={1}>
                  <Button
                    variant="outlined"
                    disabled={isApplyingBulkJournalProgress}
                    onClick={() => setJournalBulkProgressSteps({})}
                  >
                    Очистити
                  </Button>
                  <Button
                    variant="contained"
                    disabled={isApplyingBulkJournalProgress}
                    onClick={() => void applyBulkJournalProgress()}
                    sx={{ color: "#1a1a14" }}
                  >
                    {isApplyingBulkJournalProgress
                      ? "Застосовую..."
                      : "Застосувати"}
                  </Button>
                </Stack>
              </div>
              <div className="salary-step-bar documents-journal-bulk-progress-steps">
                {ubdWorkflowSteps.map((step, index) => (
                  <button
                    className={[
                      "salary-step-node",
                      journalBulkProgressSteps[step.key] ? "done" : "",
                    ]
                      .filter(Boolean)
                      .join(" ")}
                    key={step.key}
                    type="button"
                    onClick={() => toggleJournalBulkProgressStep(step.key)}
                    title={
                      journalBulkProgressSteps[step.key]
                        ? `${step.title} · зняти`
                        : `${step.title} · додати всім обраним, якщо ще немає`
                    }
                  >
                    <span>{index + 1}</span>
                    <i />
                    <strong>{step.title}</strong>
                  </button>
                ))}
              </div>
            </div>
          ) : null}
          <div
            className={[
              "documents-journal-table",
              showUbdExitDateColumn ? "documents-journal-table-with-exit-date" : "",
              showFormPurposeColumn
                ? "documents-journal-table-with-form-purpose"
                : "",
            ]
              .filter(Boolean)
              .join(" ")}
          >
            <div
              className={[
                "documents-journal-row",
                "header",
                showUbdExitDateColumn ? "documents-journal-row-with-exit-date" : "",
                showFormPurposeColumn
                  ? "documents-journal-row-with-form-purpose"
                  : "",
              ]
                .filter(Boolean)
                .join(" ")}
            >
              <span
                className="documents-journal-export-check"
                title="Експортувати всі / зняти всі"
                onClick={(event) => event.stopPropagation()}
              >
                <Checkbox
                  checked={
                    journalExportAllSelected
                      ? true
                      : journalExportSomeSelected
                        ? "indeterminate"
                        : false
                  }
                  disabled={!filteredJournalDocuments.length}
                  onCheckedChange={(checked) =>
                    toggleJournalExportAll(checked === true)
                  }
                />
              </span>
              <span>Службовець</span>
              <span>Статус службовця</span>
              <span>Документ</span>
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
              {showFormPurposeColumn ? <span>Для чого форма</span> : null}
              {showUbdExitDateColumn ? (
                <button
                  type="button"
                  className="documents-journal-sort"
                  aria-label={
                    journalSortField === "taskPeriodStart" &&
                    journalSortDirection === "desc"
                      ? "Сортувати за датою виходу: спочатку старіші"
                      : "Сортувати за датою виходу: спочатку новіші"
                  }
                  title="Сортувати за датою виходу (перша дата періоду)"
                  onClick={() => toggleJournalSort("taskPeriodStart")}
                >
                  Вихід від
                  {journalSortField === "taskPeriodStart"
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
                const liveFields = liveJournalFields(journalDocument);
                const skippedSzch = readDocumentSkippedDueToSzch(liveFields);
                const statusNote = [
                  skippedSzch ? "Не потрібно: СЗЧ" : "",
                  liveDocumentStatusNote(journalDocument),
                ]
                  .filter(Boolean)
                  .join(" · ");
                const isComplete = !skippedSzch && progress >= 100;
                const exportBlocked = journalUbdExportBlockedIds.has(
                  document.id,
                );
                const hasMissingFields =
                  !skippedSzch &&
                  !isComplete &&
                  documentHasEmptyInputs(journalDocument.type, liveFields);
                const hasBasisDateMismatch =
                  !skippedSzch &&
                  !isComplete &&
                  !hasMissingFields &&
                  documentHasBasisDateMismatch(
                    journalDocument.type,
                    liveFields,
                  );
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
                      showFormPurposeColumn
                        ? "documents-journal-row-with-form-purpose"
                        : "",
                      document.id === selectedDocumentId ? "active" : "",
                      isComplete ? "is-complete" : "",
                      hasMissingFields ? "is-incomplete" : "",
                      hasBasisDateMismatch ? "is-basis-mismatch" : "",
                      skippedSzch ? "is-skipped-szch" : "",
                    ]
                      .filter(Boolean)
                      .join(" ")}
                    key={document.id}
                    role="button"
                    tabIndex={0}
                    title={
                      skippedSzch
                        ? "Не потрібно робити: СЗЧ — не входить в експорт"
                        : isComplete
                        ? journalDocument.type === "ubdReport"
                          ? "Весь прогрес заповнений — готовий, в експорт «Не подавалися» не входить"
                          : "Весь прогрес заповнений"
                        : hasMissingFields
                          ? "Не заповнені обов’язкові поля"
                          : hasBasisDateMismatch
                            ? "Дата БР не збігається з початком періоду завдань"
                            : undefined
                    }
                    onClick={openDocument}
                    onKeyDown={(event) => {
                      if (event.key !== "Enter" && event.key !== " ") return;
                      event.preventDefault();
                      openDocument();
                    }}
                  >
                    <span
                      className="documents-journal-export-check"
                      title={
                        skippedSzch
                          ? "Не входить в експорт: документ не потрібен через СЗЧ"
                          : exportBlocked
                            ? "Не входить в експорт «Не подавалися»: прогрес заповнений повністю"
                            : undefined
                      }
                      onClick={(event) => event.stopPropagation()}
                      onKeyDown={(event) => event.stopPropagation()}
                    >
                      <Checkbox
                        checked={
                          !exportBlocked &&
                          !journalExportDeselectedIds[document.id]
                        }
                        disabled={exportBlocked}
                        onCheckedChange={(checked) =>
                          toggleJournalExportDocument(
                            document.id,
                            checked === true,
                          )
                        }
                      />
                    </span>
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
                    <span className="documents-journal-progress">
                      <i style={{ width: `${progress}%` }} />
                      <b>{progress}%</b>
                    </span>
                    <span>
                      {documentHighestWorkflowStatusLabel(
                        journalDocument,
                        document.id === selectedDocumentId
                          ? mergeSalaryWorkflow(journalDocument.workflow)
                          : undefined,
                      )}
                    </span>
                    <span
                      className="documents-journal-note"
                      title={statusNote}
                    >
                      {statusNote || "—"}
                    </span>
                    <span>{getDocumentFileSummary(document)}</span>
                    {showFormPurposeColumn ? (
                      <span
                        className="documents-journal-note"
                        title={
                          readDocumentFormPurpose(document, liveFields) ||
                          undefined
                        }
                      >
                        {readDocumentFormPurpose(document, liveFields) || "—"}
                      </span>
                    ) : null}
                    {showUbdExitDateColumn ? (
                      <span title={readDocumentUbdTaskPeriod(document) || undefined}>
                        {readDocumentUbdTaskPeriodStartLabel(document) || "—"}
                      </span>
                    ) : null}
                    <span>
                      {new Date(document.createdAt).toLocaleString("uk-UA")}
                    </span>
                    <span>
                      {new Date(document.updatedAt).toLocaleString("uk-UA")}
                    </span>
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
                  </div>
                );
              })
            ) : (
              <div className="documents-journal-empty">
                {journalDocuments.length
                  ? "Немає документів за вибраними фільтрами."
                  : "Документи ще не створювались."}
              </div>
            )}
            </div>
          </div>
        </section>
        {mode === "ubdReport" && selectedDocument?.type === "ubdReport" ? (
          <section className="documents-journal-ubd" id="documents-journal-ubd">
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
                  </Typography>
                </Box>
                <Stack direction="row" spacing={1}>
                  {questionnaireHeaderButton}
                  <Button
                    variant="outlined"
                    startIcon={<ArticleOutlinedIcon />}
                    disabled={ubdFields.basisNotReady}
                    onClick={() => void saveUbdAsWord()}
                  >
                    Зберегти у Word
                  </Button>
                  <Button
                    variant="contained"
                    startIcon={<PictureAsPdfOutlinedIcon />}
                    disabled={ubdFields.basisNotReady}
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
          <section className="documents-journal-ubd" id="documents-journal-ubd">
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
          <section className="documents-journal-ubd" id="documents-journal-ubd">
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
          <section className="documents-journal-ubd" id="documents-journal-ubd">
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
          <section className="documents-journal-ubd" id="documents-journal-ubd">
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
          <section className="documents-journal-ubd" id="documents-journal-ubd">
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
          <section className="documents-journal-ubd" id="documents-journal-ubd">
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
        {mode === "lostMilitaryId" &&
        selectedDocument?.type === "lostMilitaryId" ? (
          <section className="documents-journal-ubd" id="documents-journal-ubd">
            <header className="topbar analytics-topbar salary-document-topbar">
              <div className="salary-document-main-row">
                <Box className="salary-document-title">
                  <Typography component="h1" variant="h4">
                    Втрата військового квитка
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
                    onClick={() => void saveLostMilitaryIdAsWord()}
                  >
                    Рапорт у Word
                  </Button>
                  <Button
                    variant="contained"
                    startIcon={<FileDownloadOutlinedIcon />}
                    onClick={() => void saveLostMilitaryIdKit()}
                    sx={{ color: "#1a1a14" }}
                  >
                    Сформувати комплект
                  </Button>
                </Stack>
              </div>
              {documentStepProgress}
            </header>
            <section className="salary-documents-layout ubd-report-layout documents-journal-ubd-layout">
              {lostMilitaryIdWorkspace}
            </section>
          </section>
        ) : null}
        {mode === "salaryPowerAttorney" &&
        selectedDocument?.type === "salaryPowerAttorney" ? (
          <section className="documents-journal-ubd" id="documents-journal-ubd">
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
      <main className="main-panel documents-ubd-page">
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
                    className={personDocumentShellClass(document)}
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
            {personDocumentCreateSelect}
          </section>

          {ticketWorkspace}
        </section>
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  if (mode === "lostMilitaryId") {
    return (
      <>
      <main className="main-panel documents-ubd-page">
        <header className="topbar analytics-topbar salary-document-topbar">
          <div className="salary-document-main-row">
            <Box className="salary-document-title">
              <Typography component="h1" variant="h4">
                Втрата військового квитка
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
                  onClick={() => void saveLostMilitaryIdAsWord()}
                >
                  Рапорт у Word
                </Button>
                <Button
                  variant="contained"
                  startIcon={<FileDownloadOutlinedIcon />}
                  onClick={() => void saveLostMilitaryIdKit()}
                  sx={{ color: "#1a1a14" }}
                >
                  Сформувати комплект
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
                    className={personDocumentShellClass(document)}
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
            {personDocumentCreateSelect}
          </section>

          {lostMilitaryIdWorkspace}
        </section>
      </main>
      {questionnairePreviewDialog}
      </>
    );
  }

  if (mode === "ubdReport") {
    return (
      <>
      <main className="main-panel documents-ubd-page">
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
                  disabled={ubdFields.basisNotReady}
                  onClick={() => void saveUbdAsWord()}
                >
                  Зберегти у Word
                </Button>
                <Button
                  variant="contained"
                  startIcon={<PictureAsPdfOutlinedIcon />}
                  disabled={ubdFields.basisNotReady}
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
                    className={personDocumentShellClass(document)}
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
            {personDocumentCreateSelect}
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
      <main className="main-panel documents-ubd-page">
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
                    className={personDocumentShellClass(document)}
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
            {personDocumentCreateSelect}
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
      <main className="main-panel documents-ubd-page">
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
                    className={personDocumentShellClass(document)}
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
            {personDocumentCreateSelect}
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
      <main className="main-panel documents-ubd-page">
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
                    className={personDocumentShellClass(document)}
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
            {personDocumentCreateSelect}
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
      <main className="main-panel documents-ubd-page">
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
                    className={personDocumentShellClass(document)}
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
            {personDocumentCreateSelect}
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
      <main className="main-panel documents-ubd-page">
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
                    className={personDocumentShellClass(document)}
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
            {personDocumentCreateSelect}
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
                    className={personDocumentShellClass(document)}
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
            {personDocumentCreateSelect}
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
