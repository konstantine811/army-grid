import {
  useCallback,
  useDeferredValue,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import {
  Alert,
  Box,
  Button,
  Chip,
  LinearProgress,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { AddPhotoAlternateOutlinedIcon } from "@/components/sci/icons";
import { ArticleOutlinedIcon } from "@/components/sci/icons";
import { ArrowLeftOutlinedIcon } from "@/components/sci/icons";
import { CalendarMonthOutlinedIcon } from "@/components/sci/icons";
import { DeleteOutlineOutlinedIcon } from "@/components/sci/icons";
import { FileUploadOutlinedIcon } from "@/components/sci/icons";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import { FormatListBulletedOutlinedIcon } from "@/components/sci/icons";
import { LoginOutlinedIcon } from "@/components/sci/icons";
import { LocationOnOutlinedIcon } from "@/components/sci/icons";
import { LogoutOutlinedIcon } from "@/components/sci/icons";
import { PersonOutlinedIcon } from "@/components/sci/icons";
import { PersonSearchOutlinedIcon } from "@/components/sci/icons";
import { PictureAsPdfOutlinedIcon } from "@/components/sci/icons";
import { SearchOutlinedIcon } from "@/components/sci/icons";
import { WarningAmberOutlinedIcon } from "@/components/sci/icons";
import { PushPinOutlinedIcon } from "@/components/sci/icons";
import { InfoOutlinedIcon } from "@/components/sci/icons";
import {
  api,
  type BackendEjournalImport,
  type BackendPersonDocument,
  type BackendPersonQuestionnaire,
  type BackendPersonnelRosterLatest,
} from "../../api";
import { useAuth } from "../../auth/AuthProvider";
import {
  CacheKeys,
  deleteDataCache,
  fetchWithCache,
  jsonChanged,
  readDataCache,
} from "../../data/idbDataCache";
import { STAFF_SHEET_SYNCED_EVENT } from "../../data/staffSheetAutoSync";
import { loadSharedEjournalImports, loadSharedRosterLatest } from "../../data/sharedAppData";
import {
  loadPersonnelDataset,
  personnelDatasetToPreview,
  type PersonnelDataset,
} from "../../data/personnelDataset";
import {
  extractFighterStatusFieldRows,
  getFighterStatusFieldTone,
  normalizeRosterMatchText,
} from "./fighterStatusImport";
import {
  readWorkbookSnapshot,
  valueToDisplay,
} from "../../excelRoundTrip";
import type { DbPreviewState, EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { parseDbColumns } from "../ejournal/ejournalUtils";
import { PhotoCropDialog, type CropRect } from "./PhotoCropDialog";
import { FloatingQuestionnairePreview } from "./FloatingQuestionnairePreview";
import { PersonnelVirtualList } from "./PersonnelVirtualList";
import { QuestionnaireDiskSearchDialog } from "./QuestionnaireDiskSearchDialog";
import {
  PERSON_CARD_FIELDS,
  PERSON_SECTION_LABELS,
  buildPersonListSummary,
  buildPersonSummary,
  buildQuestionnaireExportFileName,
  collectPersonCallSignFieldValues,
  createDefaultActionForm,
  dataUrlToFile,
  dataUrlToObjectUrl,
  downloadQuestionnairePdf,
  extractPersonCallSign,
  extractPhones,
  findEjournalPersonnelSheet,
  formatPersonFieldValue,
  formatUaPhoneDisplay,
  buildOrphanAttachmentMigrationPairs,
  getPersonDisplayName,
  getPersonFieldValue,
  inferRosterFieldLabel,
  isLikelyPersonnelRow,
  resolvePersonDisplayNameFromRoster,
  resolvePersonRosterStatus,
  cleanPersonDisplayName,
  looksLikePersonBirthDate,
  migratePersonAttachmentsBetweenIds,
  normalizePersonBirthKey,
  normalizeUaPhone,
  pickFullPositionFromPersonRow,
  resolveMorningGeneralListColumnLabel,
  resolvePersonIdentityKey,
  resolvePersonBirthDate,
  formatPersonBirthDateWithAge,
  computeFullYearsFromBirthDate,
  isPositionIndexField,
  loadAllEjournalSheetRows,
  sheetRowsCacheKey,
  personActions,
  renameQuestionnaireFile,
  revokeQuestionnairePreviewUrl,
  resolvePersonFieldKey,
  type PersonAction,
  type PersonActionForm,
  type PersonFieldDef,
  type PersonnelRecord,
} from "./personnelUtils";
import { downloadBlob, sanitizeFileName } from "../../shared/browserExport";
import { notifyPersonnelAttachmentChanged } from "../../shared/personnelAttachmentSync";
import type { QuestionnairePdfSource } from "./questionnaireShare";
import {
  extractPhonesFromDocuments,
  migrateStoredPersonPhones,
  readStoredPersonPhones,
  uniqueNormalizedPhones,
  upsertPersonPhonesDocument,
  writeStoredPersonPhones,
} from "./personPhonesStore";
import {
  buildQuestionnairePresencePeople,
  collectPersonAttachmentLookupIds,
  loadPersonDocumentsForRow,
  loadPersonQuestionnaireForRow,
  questionnaireFileMatchesPerson,
} from "./personAttachments";
import { migrateStoredPersonSignatures } from "./personSignatureStore";
import {
  formatAnketaBulkMergeReport,
  mergeCachedAnketaToPersonnel,
} from "../anketa-data/anketaPersonMerge";
import {
  mergeAnketaCreatedRowsIntoPreview,
} from "../anketa-data/anketaPersonnelRosterCreate";
import {
  formatVkTpvDovidkyMergeReport,
  mergeVkTpvDovidkyRecords,
} from "./vkTpvDovidkyImport";
import { runParseVkTpvDovidkyHeavy } from "../anketa-data/runStaffSheetHeavyJobs";
import {
  importStaffSheetFromFile,
} from "../anketa-data/staffSheetImport";
import {
  applyPersonnelMergeDelta,
  buildRosterOnlyPreviewState,
  getRosterPersonName,
  isPersonnelFromArchive,
  isPersonnelInStaffRoster,
  ROSTER_FIELD_PREFIX,
  type PersonnelMergeDelta,
} from "./personnelRosterMerge";
import { runHeavyJob } from "../../workers/runHeavyJob";
import {
  compressPhotoDataUrl,
  createPhotoThumbnailDataUrl,
} from "./photoCompression";
import { personnelSearchMatchesQuery } from "./personnelSearch";

function PersonCardName({ name }: { name: string }) {
  const ref = useRef<HTMLHeadingElement>(null);

  useLayoutEffect(() => {
    const el = ref.current;
    const parent = el?.parentElement;
    if (!el || !parent) return;

    const fit = () => {
      el.style.fontSize = "";
      const available = parent.clientWidth;
      if (available <= 0) return;
      const width = el.scrollWidth;
      if (width <= available) return;
      const current = parseFloat(getComputedStyle(el).fontSize);
      if (!current) return;
      el.style.fontSize = `${Math.max(13, (current * available) / width)}px`;
    };

    fit();
    const observer = new ResizeObserver(fit);
    observer.observe(parent);
    return () => observer.disconnect();
  }, [name]);

  return (
    <h2
      ref={ref}
      className="sci-text sci-text-h4 person-card-name"
      title={name}
      style={{
        ["--name-len" as string]: Math.max(name.trim().length, 8),
      }}
    >
      {name}
    </h2>
  );
}

const PERSONNEL_FOCUS_KEY = "army-grid:focus-personnel";
const ATTACHMENT_HEAL_SESSION_KEY = "army-grid:attachments-healed-v2";
const MAX_QUESTIONNAIRE_FILE_BYTES = 350 * 1024 * 1024;

const formatFileSize = (bytes: number) => {
  if (!Number.isFinite(bytes)) return "";
  const mb = bytes / 1024 / 1024;
  return `${mb.toFixed(mb >= 10 ? 0 : 1)} MB`;
};

const normalizeRosterText = normalizeRosterMatchText;

const normalizePersonnelSearchText = (value: unknown) =>
  valueToDisplay(value as Parameters<typeof valueToDisplay>[0])
    .replace(/[ʼ’']/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const getRawCallSignSearchValues = (row: EjournalPreviewRow) =>
  collectPersonCallSignFieldValues(row);

const personnelMergeDeltaCache = new WeakMap<
  EjournalPreviewRow[],
  WeakMap<EjournalPreviewRow[], PersonnelMergeDelta>
>();
const rosterPreviewRowsCache = new WeakMap<
  BackendPersonnelRosterLatest,
  EjournalPreviewRow[]
>();

const mergeRosterRowsIntoPreviewState = async (
  preview: DbPreviewState,
  rosterRows: EjournalPreviewRow[],
  anketaCreatedRows: EjournalPreviewRow[] = [],
) => {
  let byRoster = personnelMergeDeltaCache.get(preview.rows);
  if (!byRoster) {
    byRoster = new WeakMap();
    personnelMergeDeltaCache.set(preview.rows, byRoster);
  }
  let delta = byRoster.get(rosterRows);
  if (!delta) {
    delta = await runHeavyJob({
        type: "mergePersonnel",
        preview,
        rosterRows,
      });
    byRoster.set(rosterRows, delta);
  }
  return {
    ...preview,
    rows: mergeAnketaCreatedRowsIntoPreview(
      applyPersonnelMergeDelta(preview.rows, delta),
      anketaCreatedRows,
    ),
  };
};

const readPersonnelFocusTarget = () => {
  const params = new URLSearchParams(window.location.search);
  const fromQuery = {
    rowId: params.get("rowId")?.trim() || "",
    externalId: params.get("externalId")?.trim() || "",
  };
  if (fromQuery.rowId || fromQuery.externalId) return fromQuery;

  try {
    const raw = window.localStorage.getItem(PERSONNEL_FOCUS_KEY);
    if (!raw) return { rowId: "", externalId: "" };
    const parsed = JSON.parse(raw) as {
      rowId?: string;
      externalId?: string;
    };
    return {
      rowId: parsed.rowId?.trim() || "",
      externalId: parsed.externalId?.trim() || "",
    };
  } catch {
    return { rowId: "", externalId: "" };
  }
};

const clearPersonnelFocusTarget = () => {
  try {
    window.localStorage.removeItem(PERSONNEL_FOCUS_KEY);
  } catch {
    // ignore
  }
  const url = new URL(window.location.href);
  if (url.searchParams.has("rowId") || url.searchParams.has("externalId")) {
    url.searchParams.delete("rowId");
    url.searchParams.delete("externalId");
    window.history.replaceState(
      { page: "personnel" },
      "",
      `${url.pathname}${url.search}${url.hash}`,
    );
  }
};

export function PersonnelPage({
  onOpenDocuments,
}: {
  onOpenDocuments: (
    row: EjournalPreviewRow,
    mode?:
      | "default"
      | "salaryPowerAttorney"
      | "ubdReport"
      | "form6Report"
      | "form12Report"
      | "serviceCharacteristic"
      | "zhbdCertificate"
      | "ubdRestoreReport"
      | "temporaryMilitaryId"
      | "lostMilitaryId",
    meta?: { fullPosition?: string },
  ) => void;
}) {
  const { canEditArea } = useAuth();
  const canEdit = canEditArea("personnel");
  const [dbPreview, setDbPreview] = useState<DbPreviewState | null>(null);
  const [rosterLabels, setRosterLabels] = useState<Record<string, string>>({});
  const [selectedRowId, setSelectedRowId] = useState("");
  const [mobilePane, setMobilePane] = useState<"list" | "card" | "side">("list");
  const personnelFocusLockRef = useRef<{
    rowId: string;
    externalId: string;
  } | null>(null);
  const personnelLoadGenerationRef = useRef(0);
  const personnelLoadControllerRef = useRef<AbortController | null>(null);
  const requestedPhotoIdsRef = useRef(new Set<string>());
  const requestedFullPhotoIdsRef = useRef(new Set<string>());
  const availablePhotoIdsRef = useRef<Set<string> | null>(null);
  const photoIndexPromiseRef = useRef<Promise<Set<string>> | null>(null);
  const holdStatusUntilRef = useRef(0);
  const [query, setQuery] = useState("");
  const [editValues, setEditValues] = useState<Record<string, string>>({});
  const [activePersonAction, setActivePersonAction] =
    useState<PersonAction | null>(null);
  const [personActionForm, setPersonActionForm] = useState<PersonActionForm>(
    () => createDefaultActionForm(),
  );
  const [photoByExternalId, setPhotoByExternalId] = useState<Record<string, string>>({});
  const [phonesByExternalId, setPhonesByExternalId] = useState<
    Record<string, string[]>
  >(() => readStoredPersonPhones());
  const [phoneDocByExternalId, setPhoneDocByExternalId] = useState<
    Record<string, BackendPersonDocument>
  >({});
  const [phoneDraft, setPhoneDraft] = useState("");
  const [isSavingPhone, setIsSavingPhone] = useState(false);
  const [questionnaireByExternalId, setQuestionnaireByExternalId] = useState<
    Record<string, true>
  >({});
  const [questionnairePresenceStatus, setQuestionnairePresenceStatus] =
    useState<"loading" | "ready" | "error">("loading");
  const [staffFilter, setStaffFilter] = useState<
    "all" | "in" | "archive"
  >("in");
  const [photoCropFile, setPhotoCropFile] = useState<File | null>(null);
  const [isPhotoCropOpen, setIsPhotoCropOpen] = useState(false);
  const [questionnaire, setQuestionnaire] =
    useState<BackendPersonQuestionnaire | null>(null);
  const [personRelatedDocuments, setPersonRelatedDocuments] = useState<
    BackendPersonDocument[]
  >([]);
  const [pendingQuestionnaireFile, setPendingQuestionnaireFile] =
    useState<File | null>(null);
  const [questionnairePreviewUrl, setQuestionnairePreviewUrl] = useState("");
  const [questionnairePreviewTitle, setQuestionnairePreviewTitle] = useState("");
  const [isQuestionnairePreviewOpen, setIsQuestionnairePreviewOpen] =
    useState(false);
  /** Floating preview only when opened from disk-search results. */
  const [isDiskFloatingPreview, setIsDiskFloatingPreview] = useState(false);
  const [diskPreviewFile, setDiskPreviewFile] = useState<File | null>(null);
  const [isDiskFloatingCrop, setIsDiskFloatingCrop] = useState(false);
  const [isDiskSearchOpen, setIsDiskSearchOpen] = useState(false);
  const [isUploadingQuestionnaire, setIsUploadingQuestionnaire] =
    useState(false);
  const [isMergingAnketaData, setIsMergingAnketaData] = useState(false);
  const [isMergingVkTpvDovidky, setIsMergingVkTpvDovidky] = useState(false);
  const [isPhotoLightboxOpen, setIsPhotoLightboxOpen] = useState(false);
  const [message, setMessage] = useState(`API: ${api.baseUrl}`);
  const [isLoading, setIsLoading] = useState(false);
  const personnelRows = useMemo<PersonnelRecord[]>(() => {
    const rows = (dbPreview?.rows ?? [])
      .filter(isLikelyPersonnelRow)
      .map((row) => ({ row, summary: buildPersonListSummary(row) }));
    const nameCounts = new Map<string, number>();
    for (const record of rows) {
      const key = normalizeRosterText(record.summary.name);
      if (!key) continue;
      nameCounts.set(key, (nameCounts.get(key) ?? 0) + 1);
    }
    return rows.map((record) => {
      const key = normalizeRosterText(record.summary.name);
      if (!key || (nameCounts.get(key) ?? 0) < 2) return record;
      return {
        ...record,
        summary: {
          ...record.summary,
          birthDate: resolvePersonBirthDate(record.row),
        },
      };
    });
  }, [dbPreview]);
  const deferredQuery = useDeferredValue(query);
  const filteredPersonnel = useMemo(() => {
    const normalizedQuery = normalizePersonnelSearchText(deferredQuery);

    return personnelRows.filter((record) => {
      const inStaff = isPersonnelInStaffRoster(record.row);
      if (staffFilter === "in" && !inStaff) return false;
      if (
        staffFilter === "archive" &&
        !isPersonnelFromArchive(record.row)
      ) {
        return false;
      }
      if (!normalizedQuery) return true;

      const searchableText = normalizePersonnelSearchText(
        [
        record.summary.name,
        resolvePersonDisplayNameFromRoster(record.row),
        getRosterPersonName(record.row),
        record.summary.callSign,
        getPersonFieldValue(record.row, ["позивний"]),
        getPersonFieldValue(record.row, ["позив"]),
        ...getRawCallSignSearchValues(record.row),
        record.summary.rank,
        record.summary.externalId,
        getPersonFieldValue(record.row, ["індекс", "посади"]),
        getPersonFieldValue(record.row, ["місце_дислокації"]),
        getPersonFieldValue(record.row, ["рнокпп_за_наявності"]),
        getPersonFieldValue(record.row, ["додаткова_інформація"]),
        getPersonFieldValue(record.row, ["військового", "квитка"]),
        ...(phonesByExternalId[record.summary.externalId] ?? []),
      ]
          .filter(Boolean)
          .join(" "),
      );

      return personnelSearchMatchesQuery(searchableText, normalizedQuery);
    });
  }, [
    personnelRows,
    phonesByExternalId,
    deferredQuery,
    staffFilter,
  ]);

  const staffCounts = useMemo(() => {
    const inStaff = personnelRows.reduce(
      (count, record) =>
        isPersonnelInStaffRoster(record.row) ? count + 1 : count,
      0,
    );
    const archive = personnelRows.reduce(
      (count, record) =>
        isPersonnelFromArchive(record.row) ? count + 1 : count,
      0,
    );
    return {
      all: personnelRows.length,
      in: inStaff,
      archive,
    };
  }, [personnelRows]);
  const questionnaireCounts = useMemo(() => {
    let withQuestionnaire = 0;
    for (const record of filteredPersonnel) {
      const externalId = record.summary.externalId;
      if (externalId && questionnaireByExternalId[externalId]) {
        withQuestionnaire += 1;
      }
    }
    return {
      withQuestionnaire,
      withoutQuestionnaire: filteredPersonnel.length - withQuestionnaire,
    };
  }, [filteredPersonnel, questionnaireByExternalId]);
  const selectedRecord = useMemo(
    () =>
      personnelRows.find((record) => record.row.__dbRowId === selectedRowId) ??
      filteredPersonnel[0] ??
      null,
    [filteredPersonnel, personnelRows, selectedRowId],
  );
  const selectedRow = selectedRecord?.row ?? null;
  const selectedSummary = useMemo(
    () => buildPersonSummary(selectedRow),
    [selectedRow],
  );
  const selectedPhoto =
    (selectedSummary.externalId && photoByExternalId[selectedSummary.externalId]) ||
    "";
  const selectedCallSign = useMemo(
    () =>
      selectedSummary.callSign?.trim() ||
      extractPersonCallSign(
        questionnaire?.fileName ?? undefined,
        selectedSummary.name,
        selectedSummary.additionalInfo,
      ),
    [
      questionnaire?.fileName,
      selectedSummary.additionalInfo,
      selectedSummary.callSign,
      selectedSummary.name,
    ],
  );
  const questionnaireExportFileName = useMemo(() => {
    const stored = String(questionnaire?.fileName ?? "").trim();
    if (
      stored &&
      questionnaireFileMatchesPerson(stored, [selectedSummary.name])
    ) {
      return sanitizeFileName(stored);
    }
    return sanitizeFileName(
      buildQuestionnaireExportFileName(selectedSummary.name, selectedCallSign),
    );
  }, [questionnaire?.fileName, selectedCallSign, selectedSummary.name]);
  const currentQuestionnaireShareSource = useMemo((): QuestionnairePdfSource | null => {
    if (pendingQuestionnaireFile) return { file: pendingQuestionnaireFile };
    if (diskPreviewFile) return { file: diskPreviewFile };
    if (questionnaire?.fileData) return { fileData: questionnaire.fileData };
    return null;
  }, [diskPreviewFile, pendingQuestionnaireFile, questionnaire?.fileData]);
  const editableFields = useMemo(
    () =>
      PERSON_CARD_FIELDS.map((field) => ({
        ...field,
        key: resolvePersonFieldKey(selectedRow, field.parts),
      })).filter(
        (field): field is PersonFieldDef & { key: string } => Boolean(field.key),
      ),
    [selectedRow],
  );
  const editableFieldsBySection = useMemo(() => {
    const birthDate = String(selectedSummary.birthDate ?? "").trim();
    const sections: PersonFieldDef["section"][] = [
      "identity",
      "service",
      "orders",
      "contacts",
    ];
    return sections
      .map((section) => ({
        section,
        label: PERSON_SECTION_LABELS[section],
        fields: editableFields.filter((field) => {
          if (field.section !== section) return false;
          // Місце перебування вже в шапці картки.
          if (field.parts.includes("місце_перебування")) return false;
          // ID з датою народження не дублюємо — дата вже в «Дата народження».
          if (field.parts.includes("id")) {
            const idValue = formatPersonFieldValue(
              selectedRow?.[field.key],
              field,
            ).trim();
            if (looksLikePersonBirthDate(idValue)) return false;
            if (
              birthDate &&
              normalizePersonBirthKey(idValue) ===
                normalizePersonBirthKey(birthDate)
            ) {
              return false;
            }
          }
          return true;
        }),
      }))
      .filter((group) => group.fields.length > 0);
  }, [editableFields, selectedRow, selectedSummary.birthDate]);

  const birthDateWithAge = useMemo(
    () => formatPersonBirthDateWithAge(selectedSummary.birthDate),
    [selectedSummary.birthDate],
  );
  const rosterFieldRows = useMemo(() => {
    const birthDate = String(selectedSummary.birthDate ?? "").trim();
    const cardName = cleanPersonDisplayName(selectedSummary.name);
    return Object.entries(selectedRow ?? {})
      .filter(
        ([key, value]) =>
          key.startsWith(ROSTER_FIELD_PREFIX) &&
          !key.includes("fighter_status_") &&
          valueToDisplay(value as Parameters<typeof valueToDisplay>[0]).trim(),
      )
      .map(([key, value]) => {
        const sourceKey = key.slice(ROSTER_FIELD_PREFIX.length);
        const displayed = valueToDisplay(
          value as Parameters<typeof valueToDisplay>[0],
        ).trim();
        const label = inferRosterFieldLabel(sourceKey, displayed, rosterLabels);
        const labelNorm = label
          .trim()
          .toLocaleLowerCase("uk-UA")
          .replace(/_/g, " ");
        const isPibField =
          labelNorm === "піб" ||
          labelNorm === "прізвище" ||
          labelNorm.includes("піб") ||
          /(^|_)(піб|прізвище|column_14)(_|$)/i.test(sourceKey);
        return {
          key,
          sourceKey,
          label,
          value: isPibField ? cleanPersonDisplayName(displayed) : displayed,
          isPibField,
        };
      })
      .filter((field) => {
        const labelNorm = field.label
          .trim()
          .toLocaleLowerCase("uk-UA")
          .replace(/_/g, " ");
        const keyNorm = field.sourceKey.toLocaleLowerCase("uk-UA");
        const isYearField =
          labelNorm === "рік" ||
          labelNorm === "рік народження" ||
          keyNorm === "рік" ||
          keyNorm === "rik" ||
          /(^|_)(рік|year|column_17)(_|$)/i.test(field.sourceKey);
        const isBirthDateField =
          labelNorm === "дата народження" ||
          labelNorm.includes("дата народ") ||
          (keyNorm.includes("народ") &&
            (keyNorm.includes("дата") || keyNorm.includes("день"))) ||
          /(^|_)(column_16)(_|$)/i.test(field.sourceKey);
        const isFullYearsField =
          labelNorm === "повних років" ||
          labelNorm.includes("повних років") ||
          /(^|_)(column_18)(_|$)/i.test(field.sourceKey);
        const isStayPlaceField =
          labelNorm === "місце перебування" ||
          labelNorm.includes("перебуван") ||
          labelNorm === "дислокація" ||
          labelNorm.includes("дислокац") ||
          /(^|_)(column_31|column_40)(_|$)/i.test(field.sourceKey);
        const isPositionField =
          labelNorm === "посада" ||
          labelNorm === "повна посада" ||
          (labelNorm.includes("посада") &&
            !labelNorm.includes("індекс") &&
            !labelNorm.includes("прийняття")) ||
          /(^|_)(column_5|column_7)(_|$)/i.test(field.sourceKey);
        const isRosterStatusField =
          labelNorm === "статус" ||
          /(^|_)(column_21|column_37)(_|$)/i.test(field.sourceKey);

        // ПІБ зі штатки часто має чужу дату в дужках — після очистки це той самий рядок.
        if (
          field.isPibField &&
          cardName &&
          cleanPersonDisplayName(field.value) === cardName
        ) {
          return false;
        }
        // Дублі зі Штатки ховаємо, якщо дата народження вже в шапці картки
        // (у т.ч. підставлені зі Штатки через resolvePersonBirthDate).
        if (birthDate && (isYearField || isBirthDateField || isFullYearsField)) {
          return false;
        }
        // «Рік» з повною датою — не показуємо окремо, вона піде в шапку.
        if (!birthDate && isYearField && looksLikePersonBirthDate(field.value)) {
          return false;
        }
        // Місце перебування / посада / статус уже в шапці.
        if (isStayPlaceField || isPositionField || isRosterStatusField) {
          return false;
        }
        return true;
      });
  }, [rosterLabels, selectedRow, selectedSummary.birthDate, selectedSummary.name]);
  const rosterStatus = useMemo(
    () => resolvePersonRosterStatus(selectedRow, rosterLabels),
    [rosterLabels, selectedRow],
  );
  const fighterStatusFieldRows = useMemo(
    () => extractFighterStatusFieldRows(selectedRow, rosterLabels),
    [rosterLabels, selectedRow],
  );
  const additionalInfoKey = useMemo(
    () => resolvePersonFieldKey(selectedRow, ["додаткова_інформація"]),
    [selectedRow],
  );
  /** Live-parsed from «Додаткова інформація», plus phones saved separately from imports. */
  const savedPhones =
    (selectedSummary.externalId &&
      phonesByExternalId[selectedSummary.externalId]) ||
    [];
  const parsedPhones = useMemo(() => {
    const draft =
      (additionalInfoKey && editValues[additionalInfoKey]) ||
      selectedSummary.additionalInfo;
    return uniqueNormalizedPhones([
      ...extractPhones(draft),
      ...((selectedSummary.externalId &&
        phonesByExternalId[selectedSummary.externalId]) ||
        []),
    ]);
  }, [
    additionalInfoKey,
    editValues,
    phonesByExternalId,
    selectedSummary.additionalInfo,
    selectedSummary.externalId,
  ]);

  useEffect(() => {
    setEditValues(
      Object.fromEntries(
        editableFields.map((field) => {
          const raw = selectedRow?.[field.key];
          let text = formatPersonFieldValue(raw, field);
          // Дата народження: мердж ООС + Штатка — одне поле без дублів у шапці.
          if (
            field.parts.includes("дата_народження") &&
            !String(text ?? "").trim()
          ) {
            text = resolvePersonBirthDate(selectedRow);
          }
          return [field.key, text];
        }),
      ),
    );
    setPhoneDraft("");
  }, [editableFields, selectedRow]);

  useEffect(() => {
    if (selectedRecord?.row.__dbRowId && !selectedRowId) {
      setSelectedRowId(selectedRecord.row.__dbRowId);
    }
  }, [selectedRecord, selectedRowId]);

  useEffect(() => {
    if (!selectedRowId || !selectedRow) return;

    const externalId = selectedSummary.externalId;
    const fullName = selectedSummary.name;
    if (!externalId && !fullName) return;

    let isCancelled = false;

    void (async () => {
      try {
        const profile = externalId
          ? await api.getPersonnelProfile(externalId, fullName)
          : null;

        if (isCancelled) return;

        if (externalId) {
          const { document, phones } = extractPhonesFromDocuments(
            profile?.documents,
          );
          if (document) {
            setPhoneDocByExternalId((current) => ({
              ...current,
              [externalId]: document,
            }));
          }
          const localPhones = uniqueNormalizedPhones([
            ...(readStoredPersonPhones()[externalId] ?? []),
            ...phones,
          ]);
          if (phones.length) {
            setPhonesByExternalId((current) => {
              const merged = uniqueNormalizedPhones([
                ...(current[externalId] ?? []),
                ...phones,
              ]);
              if (
                merged.length === (current[externalId] ?? []).length &&
                merged.every(
                  (phone, index) => phone === current[externalId]?.[index],
                )
              ) {
                return current;
              }
              const next = { ...current, [externalId]: merged };
              writeStoredPersonPhones(next);
              return next;
            });
          }
          if (localPhones.length && !document) {
            void upsertPersonPhonesDocument(externalId, localPhones, null)
              .then((saved) => {
                if (isCancelled || !saved) return;
                setPhoneDocByExternalId((current) => ({
                  ...current,
                  [externalId]: saved,
                }));
              })
              .catch(() => {
                // Local numbers still remain if the backend rejects this document type.
              });
          }
        }

        const label = fullName || externalId || selectedRowId;
        console.group(`[Особовий склад] ${label}`);

        console.log("Рядок ООС (staging):", selectedRow);
        console.log("Профіль з БД:", profile);

        if (!profile) {
          console.groupEnd();
          return;
        }

        if (profile?.exitPeriods) {
          console.log("Періоди виходу / відсутності:", profile.exitPeriods);

          if (profile.exitPeriods.hasAny) {
            console.log("Знайдено періоди виходу:", {
              absences: profile.exitPeriods.absences.length,
              openAbsences: profile.exitPeriods.openAbsences.length,
              locationPeriods: profile.exitPeriods.locationPeriods.length,
              servicePeriods: profile.exitPeriods.servicePeriods.length,
              rosterEvents: profile.exitPeriods.rosterEvents.length,
              absentSheetRows: profile.exitPeriods.absentSheetRows.length,
              fighterStatusExitDate: profile.exitPeriods.fighterStatus?.exitDate ?? null,
              oosExitFields: profile.exitPeriods.oosExitFields,
            });
          } else {
            console.debug(
              "[Особовий склад] періодів виходу в БД немає для цієї особи.",
            );
          }
        }

        if (!profile?.person) {
          console.debug(
            "[Особовий склад] запис Person у БД відсутній — профіль зʼявиться після імпорту ЕЖООС.",
          );
        }

        console.groupEnd();
      } catch (error) {
        if (isCancelled) return;
        console.error("[Особовий склад] не вдалося завантажити профіль з БД:", error);
      }
    })();

    return () => {
      isCancelled = true;
    };
  }, [selectedRow, selectedRowId, selectedSummary.externalId, selectedSummary.name]);

  useEffect(() => {
    const externalId = selectedSummary.externalId;
    if (!externalId) return;
    if (requestedFullPhotoIdsRef.current.has(externalId)) return;
    requestedFullPhotoIdsRef.current.add(externalId);

    let isCancelled = false;
    void api
      .getPersonPhoto(externalId)
      .then(async (photo) => {
        if (isCancelled || !photo?.photoData) return;
        const compressed = await compressPhotoDataUrl(photo.photoData).catch(
          () => photo.photoData,
        );
        if (isCancelled) return;
        setPhotoByExternalId((photos) => ({
          ...photos,
          [externalId]: compressed,
        }));
        if (compressed !== photo.photoData) {
          void createPhotoThumbnailDataUrl(compressed).then((thumbnailData) =>
            api.upsertPersonPhoto(externalId, {
              photoData: compressed,
              thumbnailData,
              fileName: "photo.jpg",
              mimeType: "image/jpeg",
            }),
          );
        }
      })
      .catch(() => {
        // No saved photo yet — keep silent, avatar stays placeholder.
      });

    return () => {
      isCancelled = true;
    };
  }, [photoByExternalId, selectedSummary.externalId]);

  useEffect(() => {
    setIsPhotoLightboxOpen(false);
  }, [selectedRowId]);

  useEffect(() => {
    if (!isPhotoLightboxOpen) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setIsPhotoLightboxOpen(false);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [isPhotoLightboxOpen]);

  useEffect(() => {
    const externalId = selectedSummary.externalId;
    if (!externalId) {
      setQuestionnaire(null);
      setPersonRelatedDocuments([]);
      return;
    }

    let isCancelled = false;
    const nameIsAmbiguous =
      personnelRows.filter(
        (item) =>
          normalizeRosterText(item.summary.name) ===
          normalizeRosterText(selectedSummary.name),
      ).length > 1;
    setQuestionnaire(null);
    void loadPersonQuestionnaireForRow(selectedRow, undefined, {
      nameIsAmbiguous,
    })
      .then(async ({ questionnaire: next, resolvedExternalId }) => {
        if (isCancelled) return;
        if (
          canEdit &&
          next?.fileData &&
          resolvedExternalId &&
          resolvedExternalId !== externalId
        ) {
          try {
            const copied = await api.upsertPersonQuestionnaire(externalId, {
              fileData: next.fileData,
              fileName:
                next.fileName?.trim() ||
                sanitizeFileName(
                  buildQuestionnaireExportFileName(selectedSummary.name),
                ),
              mimeType: next.mimeType ?? "application/pdf",
            });
            if (isCancelled) return;
            setQuestionnaireByExternalId((current) => ({
              ...current,
              [externalId]: true,
            }));
            setQuestionnaire(copied ?? next);
            return;
          } catch {
            // Show the PDF found under the previous identity even if copy fails.
          }
        }
        setQuestionnaire(next);
      })
      .catch(() => {
        if (!isCancelled) setQuestionnaire(null);
      });
    void loadPersonDocumentsForRow(
      selectedRow,
      { anketaFullName: selectedSummary.name },
      { nameIsAmbiguous },
    )
      .then((documents) => {
        if (!isCancelled) setPersonRelatedDocuments(documents);
      })
      .catch(() => {
        if (!isCancelled) setPersonRelatedDocuments([]);
      });

    return () => {
      isCancelled = true;
    };
  }, [
    canEdit,
    personnelRows,
    selectedRow,
    selectedSummary.externalId,
    selectedSummary.name,
  ]);

  useEffect(() => {
    return () => {
      if (questionnairePreviewUrl) {
        revokeQuestionnairePreviewUrl(questionnairePreviewUrl);
      }
    };
  }, [questionnairePreviewUrl]);

  const exportCurrentQuestionnaire = async () => {
    const externalId = selectedSummary.externalId;
    if (
      externalId &&
      questionnaire?.fileData &&
      !pendingQuestionnaireFile &&
      !diskPreviewFile
    ) {
      try {
        const blob = await api.fetchPersonQuestionnaireFile(
          externalId,
          questionnaireExportFileName,
          true,
        );
        downloadBlob(blob, questionnaireExportFileName);
        setMessage(`Експортовано: ${questionnaireExportFileName}`);
        return;
      } catch (error) {
        setMessage(
          error instanceof Error
            ? error.message
            : "Не вдалося експортувати анкету.",
        );
        return;
      }
    }

    downloadCurrentQuestionnaire();
    if (pendingQuestionnaireFile || diskPreviewFile || questionnaire?.fileData) {
      setMessage(`Експортовано: ${questionnaireExportFileName}`);
    }
  };

  const revealCurrentQuestionnaireInFinder = async () => {
    const externalId = selectedSummary.externalId;
    if (!externalId || !questionnaire) return;

    setMessage("Шукаю оригінал анкети на диску…");
    try {
      const person = {
        externalId,
        fullName: selectedSummary.name,
        callSign: selectedSummary.callSign,
      };
      let result = await api.searchQuestionnairesOnDisk({
        people: [person],
        refreshIndex: false,
      });
      let matches = result.people[0]?.matches ?? [];
      if (!matches.length) {
        result = await api.searchQuestionnairesOnDisk({
          people: [person],
          refreshIndex: true,
        });
        matches = result.people[0]?.matches ?? [];
      }

      const storedName = String(questionnaire.fileName ?? "")
        .normalize("NFC")
        .toLocaleLowerCase("uk-UA");
      const match =
        matches.find(
          (item) =>
            item.fileName
              .normalize("NFC")
              .toLocaleLowerCase("uk-UA") === storedName,
        ) ?? matches[0];
      if (!match) {
        throw new Error(
          "Оригінальний PDF не знайдено у папці анкет. У БД збережена лише копія.",
        );
      }

      await api.revealDiskQuestionnaireInFinder(match.relativePath);
      setMessage(`Відкрито у Finder: ${match.fileName}`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося показати анкету у Finder",
      );
    }
  };

  const downloadCurrentQuestionnaire = () => {
    if (pendingQuestionnaireFile) {
      downloadQuestionnairePdf(questionnaireExportFileName, {
        file: pendingQuestionnaireFile,
      });
      return;
    }
    if (diskPreviewFile) {
      downloadQuestionnairePdf(questionnaireExportFileName, {
        file: diskPreviewFile,
      });
      return;
    }
    if (questionnaire?.fileData) {
      downloadQuestionnairePdf(questionnaireExportFileName, {
        fileData: questionnaire.fileData,
      });
    }
  };

  const loadPersonnelPhotoIndex = useCallback(async () => {
    if (availablePhotoIdsRef.current) return availablePhotoIdsRef.current;
    if (!photoIndexPromiseRef.current) {
      photoIndexPromiseRef.current = api
        .listPersonPhotos()
        .then(
          (photos) =>
            new Set(
              photos
                .filter(
                  (photo) =>
                    photo.personExternalId &&
                    (photo.hasFile ||
                      photo.hasThumbnail ||
                      Boolean(photo.photoData)),
                )
                .map((photo) => photo.personExternalId),
            ),
        )
        .then((ids) => {
          availablePhotoIdsRef.current = ids;
          return ids;
        })
        .finally(() => {
          photoIndexPromiseRef.current = null;
        });
    }
    return photoIndexPromiseRef.current;
  }, []);

  const loadVisiblePersonnelPhotos = useCallback((externalIds: string[]) => {
    void loadPersonnelPhotoIndex().then((availableIds) => {
      const nextIds = [...new Set(externalIds)]
        .filter((externalId) => Boolean(externalId) && availableIds.has(externalId))
        .filter((externalId) => {
          if (requestedPhotoIdsRef.current.has(externalId)) return false;
          requestedPhotoIdsRef.current.add(externalId);
          return true;
        });
      if (!nextIds.length) return;

      void Promise.all(
        nextIds.map(async (externalId) => {
          try {
            const thumbnail = await api.getPersonPhotoThumbnail(externalId);
            if (!thumbnail) return;
            setPhotoByExternalId((current) =>
              current[externalId]
                ? current
                : { ...current, [externalId]: thumbnail },
            );
          } catch {
            // A missing photo is expected; keep the placeholder.
          }
        }),
      );
    });
  }, [loadPersonnelPhotoIndex]);

  const loadPersonnelQuestionnaireIds = async (
    rows?: EjournalPreviewRow[],
    signal?: AbortSignal,
    prefetchedItems?: Promise<
      Array<{ personExternalId: string; fileName?: string | null }>
    >,
  ) => {
    setQuestionnairePresenceStatus("loading");
    try {
      const items =
        (await prefetchedItems) ??
        (await api.listPersonQuestionnaires({ signal }));
      const presence = await runHeavyJob({
        type: "buildQuestionnairePresence",
        people: buildQuestionnairePresencePeople(
          rows ?? dbPreview?.rows ?? [],
        ),
        questionnaires: items.map(({ personExternalId, fileName }) => ({
          personExternalId,
          fileName,
        })),
      });
      setQuestionnaireByExternalId(presence);
      setQuestionnairePresenceStatus("ready");
      return items;
    } catch {
      setQuestionnairePresenceStatus("error");
      return [];
    }
  };

  const mapRosterLatestToRows = (latest: BackendPersonnelRosterLatest | null | undefined) => {
    if (!latest?.sheet || !Array.isArray(latest.rows)) {
      return [] as EjournalPreviewRow[];
    }
    const cached = rosterPreviewRowsCache.get(latest);
    if (cached) return cached;
    const rows = latest.rows.map((row) => ({
      __dbRowId: row.id,
      __rowNumber: row.excelRowNumber,
      ...(row.values && typeof row.values === "object" && !Array.isArray(row.values)
        ? row.values
        : {}),
    })) as EjournalPreviewRow[];
    rosterPreviewRowsCache.set(latest, rows);
    return rows;
  };

  const applyRosterMeta = (latest: BackendPersonnelRosterLatest | null | undefined) => {
    if (!latest?.sheet) {
      setRosterLabels({});
      return;
    }
    const columns = parseDbColumns(latest.sheet.columns);
    setRosterLabels(
      Object.fromEntries(
        columns.map((column) => [
          column.key,
          column.label?.trim() ||
            resolveMorningGeneralListColumnLabel(column.key) ||
            column.key,
        ]),
      ),
    );
  };

  const loadLatestPersonnelRoster = async (
    force = false,
    signal?: AbortSignal,
  ) => {
    const latest = await loadSharedRosterLatest({ force, signal });
    applyRosterMeta(latest);
    return latest;
  };

  const applyPersonnelPreview = async (
    preview: DbPreviewState,
    latestRosterRows: EjournalPreviewRow[],
    sheet: BackendEjournalImport["sheets"][number],
    options?: {
      fromCache?: boolean;
      partial?: boolean;
      alreadyMerged?: boolean;
      anketaCreatedRows?: EjournalPreviewRow[];
      isCancelled?: () => boolean;
    },
  ) => {
    const safePreview = Array.isArray(preview.rows)
      ? preview
      : { ...preview, rows: [] };
    const loadedOosRows = safePreview.rows.length;
    let mergedPreview = safePreview;
    if (!options?.alreadyMerged) {
      try {
        mergedPreview = await mergeRosterRowsIntoPreviewState(
          safePreview,
          latestRosterRows,
          options?.anketaCreatedRows,
        );
      } catch {
        mergedPreview =
          (latestRosterRows.length
            ? buildRosterOnlyPreviewState(latestRosterRows, sheet)
            : null) ?? safePreview;
      }
    }
    if (options?.isCancelled?.()) return safePreview;
    if (!Array.isArray(mergedPreview.rows)) mergedPreview = safePreview;
    const rows = mergedPreview.rows.filter(isLikelyPersonnelRow);
    const storedFocus = readPersonnelFocusTarget();
    if (storedFocus.rowId || storedFocus.externalId) {
      personnelFocusLockRef.current = storedFocus;
    }
    const focusTarget = personnelFocusLockRef.current ?? {
      rowId: "",
      externalId: "",
    };
    const focusedRow =
      (focusTarget.rowId &&
        rows.find((row) => row.__dbRowId === focusTarget.rowId)) ||
      (focusTarget.externalId &&
        rows.find(
          (row) => resolvePersonIdentityKey(row) === focusTarget.externalId,
        )) ||
      null;

    setDbPreview(mergedPreview);
    setSelectedRowId((current) => {
      if (focusedRow?.__dbRowId) return focusedRow.__dbRowId;
      if (current && rows.some((row) => row.__dbRowId === current)) {
        return current;
      }
      return rows[0]?.__dbRowId ?? "";
    });
    if (focusedRow?.__dbRowId) {
      setMobilePane("card");
    }
    // Keep focus through cache→network reload; clear only after fresh apply.
    if (focusedRow && !options?.fromCache) {
      personnelFocusLockRef.current = null;
      clearPersonnelFocusTarget();
    }

    setMessage(
      focusedRow
        ? `Відкрито картку: ${getPersonDisplayName(focusedRow) || "особу"}.`
        : options?.partial
          ? `Завантажено ООС: ${loadedOosRows} з ${mergedPreview.total} · ${sheet.name}. Довантажую…`
          : options?.fromCache
            ? `Кеш: ${rows.length} записів · ${sheet.name}. Оновлюю з БД…`
            : `Завантажено особовий склад з БД: ${rows.length} записів · ${sheet.name}.`,
    );

    return mergedPreview;
  };

  const loadPersonnel = async (
    signal?: AbortSignal,
    options?: { force?: boolean },
  ) => {
    const loadGeneration = ++personnelLoadGenerationRef.current;
    const revalidateSharedData = options?.force ?? false;
    const isLoadCancelled = () =>
      loadGeneration !== personnelLoadGenerationRef.current ||
      Boolean(signal?.aborted);
    if (typeof document !== "undefined") {
    setIsLoading(true);
    if (!dbPreview?.rows.length) {
      setMessage("Завантажую актуальну Штатку та готую список особового складу…");
    }
    // These endpoints return metadata only (no PDF/photo bodies), so start
    // them in parallel with the personnel dataset instead of delaying status
    // badges and visible thumbnails until an idle callback.
    const questionnaireItemsPromise = fetchWithCache({
      key: CacheKeys.questionnairesMeta,
      force: options?.force,
      signal,
      fetcher: () => api.listPersonQuestionnaires({ signal }),
      isChanged: jsonChanged,
    }).catch(() => []);
    void loadPersonnelPhotoIndex().catch(() => new Set<string>());
    const paintDataset = async (
      dataset: PersonnelDataset,
      fromCache: boolean,
    ) => {
      const preview = personnelDatasetToPreview(dataset);
      if (!preview || isLoadCancelled()) return;
      setRosterLabels(dataset.rosterLabels);
      await applyPersonnelPreview(preview, [], preview.sheet, {
        fromCache,
        alreadyMerged: true,
        isCancelled: isLoadCancelled,
      });
    };
    try {
      const dataset = await loadPersonnelDataset({
        force: options?.force,
        signal,
        onCached: (cached) => paintDataset(cached, true),
      });
      if (isLoadCancelled()) return;
      await paintDataset(dataset, false);
      if (isLoadCancelled()) return;
      setIsLoading(false);

      const startAttachments = () => {
        if (isLoadCancelled()) return;
        const questionnairesPromise = loadPersonnelQuestionnaireIds(
          dataset.rows,
          signal,
          questionnaireItemsPromise,
        );
        const healAttachments = () => {
          if (isLoadCancelled()) return;
          void healOrphanAttachmentsInBackground(
            dataset.rows,
            isLoadCancelled,
            Promise.resolve([]),
            questionnairesPromise,
          );
        };
        if ("requestIdleCallback" in window) {
          window.requestIdleCallback(healAttachments, { timeout: 10_000 });
        } else {
          globalThis.setTimeout(healAttachments, 3_000);
        }
      };
      startAttachments();
    } catch (error) {
      if (!isLoadCancelled()) {
        setMessage(
          error instanceof Error
            ? error.message
            : "Не вдалося завантажити особовий склад.",
        );
      }
    } finally {
      if (!isLoadCancelled()) setIsLoading(false);
    }
    return;
    }

    try {
      // Cache-first: paint roster from IndexedDB before network round-trips.
      const [cachedImports, cachedRosterInitial] = await Promise.all([
        readDataCache<BackendEjournalImport[]>(CacheKeys.ejournalImports),
        readDataCache<BackendPersonnelRosterLatest | null>(CacheKeys.rosterLatest),
      ]);
      // Locally created anketa rows are not cross-device state. Visible
      // personnel comes only from server OOS + the latest server roster.
      const anketaCreatedRows: EjournalPreviewRow[] = [];
      const cachedRoster = cachedRosterInitial;

      const cachedRosterRows = mapRosterLatestToRows(cachedRoster);
      const cachedSheet = cachedImports
        ? findEjournalPersonnelSheet(cachedImports)
        : null;
      let paintedFromCache = false;
      if (
        cachedSheet &&
        cachedImports &&
        cachedRosterRows.length > 0 &&
        !isLoadCancelled()
      ) {
        const cachedPreview = await readDataCache<DbPreviewState>(
          sheetRowsCacheKey(cachedSheet),
        );
        if (cachedPreview && Array.isArray(cachedPreview.rows)) {
          paintedFromCache = true;
          applyRosterMeta(cachedRoster);
          await applyPersonnelPreview(
            cachedPreview,
            cachedRosterRows,
            cachedSheet,
            { fromCache: true, anketaCreatedRows, isCancelled: isLoadCancelled },
          );
          setIsLoading(false);
        } else if (cachedPreview) {
          await deleteDataCache(sheetRowsCacheKey(cachedSheet));
        }
      } else if (!cachedSheet && cachedRoster && !isLoadCancelled()) {
        const rosterPreview = buildRosterOnlyPreviewState(
          cachedRosterRows,
          cachedRoster.sheet,
        );
        if (rosterPreview) {
          paintedFromCache = true;
          applyRosterMeta(cachedRoster);
          await applyPersonnelPreview(
            rosterPreview,
            cachedRosterRows,
            rosterPreview.sheet,
            { fromCache: true, anketaCreatedRows, isCancelled: isLoadCancelled },
          );
          setIsLoading(false);
        }
      }

      // Shared memory/IndexedDB is authoritative within its short TTL. A normal
      // page transition reuses it; only explicit refresh/import bypasses it.
      const importsPromise = loadSharedEjournalImports({
        force: revalidateSharedData,
        signal,
      });
      const latestRosterPromise = loadLatestPersonnelRoster(
        revalidateSharedData,
        signal,
      ).catch(() => cachedRoster);
      let rosterRows: EjournalPreviewRow[] = cachedRosterRows;
      let latestRoster = cachedRoster;
      const rosterPromise = latestRosterPromise
        .then(async (latest) => {
          latestRoster = latest;
          const rows = mapRosterLatestToRows(latest);
          rosterRows = rows;
          if (!paintedFromCache && latest?.sheet && !isLoadCancelled()) {
            const rosterPreview = buildRosterOnlyPreviewState(rows, latest.sheet);
            if (rosterPreview) {
              await applyPersonnelPreview(
                rosterPreview,
                rows,
                rosterPreview.sheet,
                {
                  fromCache: true,
                  anketaCreatedRows,
                  isCancelled: isLoadCancelled,
                },
              );
              paintedFromCache = true;
              const inStaffCount = rosterPreview.rows.reduce(
                (count, row) =>
                  isLikelyPersonnelRow(row) && isPersonnelInStaffRoster(row)
                    ? count + 1
                    : count,
                0,
              );
              setMessage(
                `Штатка: ${inStaffCount} осіб у штаті. Довантажую дані ООС у фоні…`,
              );
            }
          }
          return rows;
        })
        .catch(() => rosterRows);
      const nextImports = await importsPromise;
      if (isLoadCancelled()) return;

      const sheet = findEjournalPersonnelSheet(nextImports);
      if (!sheet) {
        const latestRosterRows = await rosterPromise;
        if (isLoadCancelled()) return;
        const rosterPreview = buildRosterOnlyPreviewState(
          latestRosterRows,
          latestRoster?.sheet ?? null,
        );
        if (rosterPreview) {
          applyRosterMeta(latestRoster);
          await applyPersonnelPreview(rosterPreview, latestRosterRows, rosterPreview.sheet, {
            anketaCreatedRows,
            isCancelled: isLoadCancelled,
          });
          setMessage(
            `Немає імпорту ЕЖООС — показано ${rosterPreview.rows.length} осіб зі Штатки. Імпортуйте ЕЖООС для повних карток ООС.`,
          );
        } else {
          setDbPreview(null);
          setMessage("У БД ще немає ЕЖООС-імпорту та «Штатки» для особового складу.");
        }
        return;
      }

      let paintedFirstNetworkPage = false;
      const preview = await loadAllEjournalSheetRows(sheet, {
        force: revalidateSharedData,
        signal,
        isCancelled: isLoadCancelled,
        onPage: paintedFromCache
          ? undefined
          : async (partial) => {
              if (paintedFirstNetworkPage) return;
              paintedFirstNetworkPage = true;
              if (isLoadCancelled()) return;
              const currentRosterRows = await rosterPromise;
              if (isLoadCancelled()) return;
              await applyPersonnelPreview(partial, currentRosterRows, sheet, {
                partial: partial.rows.length < partial.total,
                anketaCreatedRows,
                isCancelled: isLoadCancelled,
              });
            },
      });
      if (isLoadCancelled()) return;

      const latestRosterRows = await rosterPromise;
      if (isLoadCancelled()) return;
      const mergedPreview = await applyPersonnelPreview(
        preview,
        latestRosterRows,
        sheet,
        { anketaCreatedRows, isCancelled: isLoadCancelled },
      );
      if (!isLoadCancelled()) setIsLoading(false);

      const startAttachments = () => {
        if (isLoadCancelled()) return;
        // Photos are large base64 payloads; visible list rows request them lazily.
        const photosPromise = Promise.resolve([]);
        const questionnairesPromise = loadPersonnelQuestionnaireIds(
          mergedPreview.rows,
          signal,
        );
        const healAttachments = () => {
          if (isLoadCancelled()) return;
          void healOrphanAttachmentsInBackground(
            mergedPreview.rows,
            isLoadCancelled,
            photosPromise,
            questionnairesPromise,
          );
        };
        if ("requestIdleCallback" in window) {
          window.requestIdleCallback(healAttachments, { timeout: 10_000 });
        } else {
          globalThis.setTimeout(healAttachments, 3_000);
        }
      };
      // Questionnaire metadata is not needed for the first list paint. Starting
      // it while the merge is settling raises peak memory on low-RAM devices.
      if ("requestIdleCallback" in window) {
        window.requestIdleCallback(startAttachments, { timeout: 5_000 });
      } else {
        globalThis.setTimeout(startAttachments, 2_000);
      }
    } catch (error) {
      if (isLoadCancelled()) return;
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося завантажити особовий склад.",
      );
    } finally {
      if (!isLoadCancelled()) setIsLoading(false);
    }
  };

  const startPersonnelLoad = (options?: { force?: boolean }) => {
    personnelLoadControllerRef.current?.abort();
    const controller = new AbortController();
    personnelLoadControllerRef.current = controller;
    return loadPersonnel(controller.signal, options);
  };

  const healOrphanAttachmentsInBackground = async (
    rows: EjournalPreviewRow[],
    isCancelled: (() => boolean) | undefined,
    photosPromise: Promise<Array<{ personExternalId: string; photoData: string }>>,
    questionnairesPromise: Promise<Array<{ personExternalId: string }>>,
  ) => {
    try {
      if (sessionStorage.getItem(ATTACHMENT_HEAL_SESSION_KEY) === "1") return;
      await new Promise<void>((resolve) => {
        window.setTimeout(resolve, 0);
      });
      if (isCancelled?.()) return;

      const [photos, questionnaires] = await Promise.all([
        photosPromise,
        questionnairesPromise,
      ]);
      if (isCancelled?.()) return;

      const currentIds = new Set<string>();
      for (const row of rows) {
        if (!isLikelyPersonnelRow(row)) continue;
        const id = resolvePersonIdentityKey(row);
        if (id) currentIds.add(id);
      }

      const orphanIds = new Set<string>();
      for (const photo of photos) {
        const id = photo.personExternalId?.trim();
        if (id && !currentIds.has(id)) orphanIds.add(id);
      }
      for (const item of questionnaires) {
        const id = item.personExternalId?.trim();
        if (id && !currentIds.has(id)) orphanIds.add(id);
      }
      const storedPhones = readStoredPersonPhones();
      for (const id of Object.keys(storedPhones)) {
        if (id && !currentIds.has(id) && storedPhones[id]?.length) {
          orphanIds.add(id);
        }
      }

      const pairs = buildOrphanAttachmentMigrationPairs(
        rows,
        orphanIds,
        questionnaires,
      );
      if (!isCancelled?.()) {
        setPhonesByExternalId(migrateStoredPersonPhones(pairs));
        migrateStoredPersonSignatures(pairs);
      }
      if (!orphanIds.size) {
        sessionStorage.setItem(ATTACHMENT_HEAL_SESSION_KEY, "1");
        return;
      }

      const migrated = await migratePersonAttachmentsBetweenIds(pairs, {
        includeDocuments: false,
        photos,
        questionnaires,
      });
      sessionStorage.setItem(ATTACHMENT_HEAL_SESSION_KEY, "1");
      if (migrated > 0 && !isCancelled?.()) {
        availablePhotoIdsRef.current = null;
        requestedPhotoIdsRef.current.clear();
        await Promise.all([
          loadPersonnelPhotoIndex(),
          loadPersonnelQuestionnaireIds(rows),
        ]);
      }
    } catch {
      // Background heal must never block the personnel list.
    }
  };

  const missingDiskSearchPeople = useMemo(() => {
    const people: Array<{
      rowId: string;
      externalId: string;
      fullName: string;
      callSign: string;
      missingQuestionnaire: boolean;
      missingPhoto: boolean;
    }> = [];
    for (const record of personnelRows) {
      const rowId = record.row.__dbRowId ?? "";
      const externalId = record.summary.externalId;
      if (!rowId || !externalId) continue;
      const missingQuestionnaire = !questionnaireByExternalId[externalId];
      if (!missingQuestionnaire) continue;
      const missingPhoto = !photoByExternalId[externalId];
      people.push({
        rowId,
        externalId,
        fullName: record.summary.name,
        callSign: record.summary.callSign,
        missingQuestionnaire,
        missingPhoto,
      });
    }
    return people;
  }, [
    personnelRows,
    photoByExternalId,
    questionnaireByExternalId,
  ]);

  const importVkTpvDovidkyWorkbook = async (file: File | undefined) => {
    if (!file) return;

    setIsMergingVkTpvDovidky(true);
    setMessage(`Читаю «${file.name}»…`);
    try {
      const snapshot = await readWorkbookSnapshot(file);
      const records = await runParseVkTpvDovidkyHeavy(snapshot);
      let lastProgressAt = 0;
      const report = await mergeVkTpvDovidkyRecords(records, {
        onProgress: (done, total) => {
          const now = Date.now();
          if (done !== total && now - lastProgressAt < 250) return;
          lastProgressAt = now;
          setMessage(`ВК ТПВ ДОВІДКИ · ${done}/${total}`);
        },
      });
      await startPersonnelLoad();
      setMessage(
        `ВК ТПВ ДОВІДКИ · ${formatVkTpvDovidkyMergeReport(report)}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося імпортувати ВК ТПВ ДОВІДКИ.",
      );
    } finally {
      setIsMergingVkTpvDovidky(false);
    }
  };

  const mergeMissingFieldsFromAnketaData = async () => {
    setIsMergingAnketaData(true);
    setMessage("Завантажую анкетні дані…");
    try {
      let lastProgressAt = 0;
      const report = await mergeCachedAnketaToPersonnel({
        onProgress: (done, total) => {
          const now = Date.now();
          if (done !== total && now - lastProgressAt < 250) return;
          lastProgressAt = now;
          setMessage(`Доповнення з анкетних даних… ${done}/${total}`);
        },
        onStatus: setMessage,
        onCreated: () => startPersonnelLoad(),
      });
      sessionStorage.removeItem(ATTACHMENT_HEAL_SESSION_KEY);
      await startPersonnelLoad();
      holdStatusUntilRef.current = Date.now() + 20_000;
      setMessage(
        `Доповнено з анкетних даних · ${formatAnketaBulkMergeReport(report)}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося доповнити особовий склад з анкет.",
      );
    } finally {
      setIsMergingAnketaData(false);
    }
  };

  const importPersonnelRosterWorkbook = async (file: File | undefined) => {
    if (!file) return;
    setIsLoading(true);
    try {
      setMessage(`Імпортую «${file.name}» у БД персоналу…`);
      const imported = await importStaffSheetFromFile(file);
      await startPersonnelLoad({ force: true });
      holdStatusUntilRef.current = Date.now() + 20_000;
      const hasTsapenko = imported.rows.some((row) =>
        normalizeRosterText(getRosterPersonName(row)).includes("цапенко"),
      );
      setMessage(
        `Штатку імпортовано з файлу в БД: ${imported.personCount} осіб · ЦАПЕНКО: ${
          hasTsapenko ? "знайдено у файлі" : "у файлі не знайдено"
        }.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося імпортувати файл «Штатка» в БД.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    let mounted = true;
    void startPersonnelLoad();
    const refreshAfterStaffSync = () => {
      if (mounted) void startPersonnelLoad({ force: true });
    };
    window.addEventListener(STAFF_SHEET_SYNCED_EVENT, refreshAfterStaffSync);
    return () => {
      mounted = false;
      personnelLoadControllerRef.current?.abort();
      window.removeEventListener(
        STAFF_SHEET_SYNCED_EVENT,
        refreshAfterStaffSync,
      );
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const saveSelectedPerson = async () => {
    if (!selectedRow?.__dbRowId) return;

    setIsLoading(true);
    try {
      const values = Object.fromEntries(
        editableFields.map((field) => {
          const raw = editValues[field.key] ?? "";
          if (!isPositionIndexField(field.parts)) return [field.key, raw];

          // Keep Excel-compatible multiline storage for multiple indexes.
          return [
            field.key,
            raw
              .split(/\s*[·,;]\s*|\s+/)
              .map((part) => part.trim())
              .filter(Boolean)
              .join("\n"),
          ];
        }),
      );
      const updatedRow = await api.updateEjournalRowValues(
        selectedRow.__dbRowId,
        values,
      );
      setDbPreview((currentPreview) => {
        if (!currentPreview) return currentPreview;

        return {
          ...currentPreview,
          rows: currentPreview.rows.map((row) =>
            row.__dbRowId === selectedRow.__dbRowId
              ? {
                  ...row,
                  ...updatedRow.values,
                  __dbRowId: selectedRow.__dbRowId,
                }
              : row,
          ),
        };
      });
      setMessage(`Картку оновлено: ${selectedSummary.name}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося зберегти картку особи.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const submitPersonAction = async () => {
    if (!selectedRow?.__dbRowId || !activePersonAction) return;

    setIsLoading(true);
    try {
      await api.createEjournalRowAction(selectedRow.__dbRowId, {
        actionType: activePersonAction.type,
        validFrom: personActionForm.validFrom || undefined,
        validTo: personActionForm.validTo || undefined,
        reason: personActionForm.reason || undefined,
        place: personActionForm.place || undefined,
        note: personActionForm.note || undefined,
        positionIndex: personActionForm.positionIndex || undefined,
        positionTitle: personActionForm.positionTitle || undefined,
        rank: personActionForm.rank || undefined,
      });
      setMessage(
        `Статусну дію збережено: ${activePersonAction.label} · ${selectedSummary.name}.`,
      );
      setActivePersonAction(null);
      setPersonActionForm(createDefaultActionForm());
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося зберегти статусну дію.",
      );
    } finally {
      setIsLoading(false);
    }
  };
  const openPhotoCrop = (
    file: File | undefined,
    options?: { floating?: boolean },
  ) => {
    if (!file) return;

    setIsDiskFloatingCrop(Boolean(options?.floating));
    setPhotoCropFile(file);
    setIsPhotoCropOpen(true);
  };
  const savePersonPhoto = async (dataUrl: string, crop: CropRect) => {
    const externalId = selectedSummary.externalId;
    if (!externalId) {
      setMessage("Не вдалося зберегти фото: у вибраної особи немає ID.");
      return;
    }

    const compressedDataUrl = await compressPhotoDataUrl(dataUrl).catch(
      () => dataUrl,
    );

    // Show immediately even if API is slow/unavailable.
    setPhotoByExternalId((photos) => ({
      ...photos,
      [externalId]: compressedDataUrl,
    }));

    try {
      const thumbnailData = await createPhotoThumbnailDataUrl(
        compressedDataUrl,
      );
      const savedPhoto = await api.upsertPersonPhoto(externalId, {
        photoData: compressedDataUrl,
        thumbnailData,
        fileName: photoCropFile?.name,
        mimeType: "image/jpeg",
        crop,
      });
      setPhotoByExternalId((photos) => ({
        ...photos,
        [externalId]: savedPhoto?.photoData || compressedDataUrl,
      }));
      notifyPersonnelAttachmentChanged(externalId, "photo");
      setMessage(`Фото збережено в БД: ${selectedSummary.name}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? `${error.message} (фото показано локально, але не збережено в БД)`
          : "Не вдалося зберегти фото в БД.",
      );
    }
  };

  const deleteSelectedPhoto = async () => {
    const externalId = selectedSummary.externalId;
    if (!externalId || !selectedPhoto) return;
    if (!window.confirm(`Видалити фото для ${selectedSummary.name || "особи"}?`)) {
      return;
    }

    try {
      await api.deletePersonPhoto(externalId);
      setPhotoByExternalId((photos) => {
        const next = { ...photos };
        delete next[externalId];
        return next;
      });
      notifyPersonnelAttachmentChanged(externalId, "photo");
      setMessage(`Фото видалено: ${selectedSummary.name}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося видалити фото з БД.",
      );
    }
  };

  const persistPersonPhones = async (externalId: string, phones: string[]) => {
    const next = uniqueNormalizedPhones(phones);
    setPhonesByExternalId((current) => {
      const updated = { ...current, [externalId]: next };
      writeStoredPersonPhones(updated);
      return updated;
    });
    try {
      const saved = await upsertPersonPhonesDocument(
        externalId,
        next,
        phoneDocByExternalId[externalId] ?? null,
      );
      setPhoneDocByExternalId((current) => {
        if (saved) return { ...current, [externalId]: saved };
        const updated = { ...current };
        delete updated[externalId];
        return updated;
      });
      return true;
    } catch {
      return false;
    }
  };

  const addSelectedPersonPhone = async () => {
    const externalId = selectedSummary.externalId;
    const normalized = normalizeUaPhone(phoneDraft);
    if (!selectedRow || !externalId) {
      setMessage("Спочатку виберіть особу зі списку.");
      return;
    }
    if (!normalized) {
      setMessage("Вкажіть український номер, наприклад 063 123 45 67.");
      return;
    }
    if (savedPhones.includes(normalized)) {
      setPhoneDraft("");
      setMessage(`Цей номер уже збережено: ${formatUaPhoneDisplay(normalized)}.`);
      return;
    }

    setIsSavingPhone(true);
    const savedToDb = await persistPersonPhones(externalId, [
      ...savedPhones,
      normalized,
    ]);
    setIsSavingPhone(false);
    setPhoneDraft("");
    setMessage(
      savedToDb
        ? `Телефон збережено: ${formatUaPhoneDisplay(normalized)}. Номер не затреться при оновленні списку.`
        : `Телефон збережено локально: ${formatUaPhoneDisplay(normalized)}. Не вдалося записати в БД.`,
    );
  };

  const removeSelectedPersonPhone = async (phone: string) => {
    const externalId = selectedSummary.externalId;
    if (!externalId) return;
    setIsSavingPhone(true);
    const savedToDb = await persistPersonPhones(
      externalId,
      savedPhones.filter((item) => item !== phone),
    );
    setIsSavingPhone(false);
    setMessage(
      savedToDb
        ? `Телефон видалено: ${formatUaPhoneDisplay(phone)}.`
        : `Телефон прибрано локально: ${formatUaPhoneDisplay(phone)}. Не вдалося оновити БД.`,
    );
  };

  const deleteSelectedQuestionnaire = async () => {
    const externalId = selectedSummary.externalId;
    if (!externalId || !questionnaire) return;
    if (!window.confirm(`Видалити анкету для ${selectedSummary.name || "особи"}?`)) {
      return;
    }

    try {
      const deleteIds = [
        ...new Set([
          externalId,
          ...collectPersonAttachmentLookupIds(selectedRow, undefined, {
            includeLooseKeys: true,
          }),
        ]),
      ].filter(Boolean);
      await Promise.all(
        deleteIds.map((id) => api.deletePersonQuestionnaire(id).catch(() => undefined)),
      );
      setQuestionnaire(null);
      setQuestionnaireByExternalId((current) => {
        const next = { ...current };
        for (const id of deleteIds) delete next[id];
        return next;
      });
      notifyPersonnelAttachmentChanged(externalId, "questionnaire");
      closeQuestionnairePreview();
      setMessage(`Анкету видалено: ${selectedSummary.name}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося видалити анкету з БД.",
      );
    }
  };

  const focusPersonByExternalId = (externalId: string) => {
    const record = personnelRows.find(
      (item) => item.summary.externalId === externalId,
    );
    if (record?.row.__dbRowId) {
      setSelectedRowId(record.row.__dbRowId);
      setMobilePane("card");
    }
  };

  const openDiskQuestionnairePreview = (file: File, title: string) => {
    const nextUrl = URL.createObjectURL(file);
    setPendingQuestionnaireFile(null);
    setDiskPreviewFile(file);
    setQuestionnairePreviewTitle(title);
    setIsDiskFloatingPreview(true);
    setQuestionnairePreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      return nextUrl;
    });
    setIsQuestionnairePreviewOpen(true);
    setMessage(title);
  };

  const openDiskPhotoCrop = (file: File, externalId: string) => {
    focusPersonByExternalId(externalId);
    openPhotoCrop(file, { floating: true });
  };

  const closeQuestionnairePreview = () => {
    setIsQuestionnairePreviewOpen(false);
    setIsDiskFloatingPreview(false);
    setDiskPreviewFile(null);
    setPendingQuestionnaireFile(null);
    setQuestionnairePreviewTitle("");
    setQuestionnairePreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      return "";
    });
  };

  const openQuestionnairePreview = async (fileData = questionnaire?.fileData) => {
    const externalId = selectedSummary.externalId;
    let nextUrl = "";

    try {
      if (fileData) {
        nextUrl = dataUrlToObjectUrl(fileData);
      } else if (externalId && !pendingQuestionnaireFile && !diskPreviewFile) {
        nextUrl = await api.createPersonQuestionnairePreviewUrl(
          externalId,
          questionnaireExportFileName,
        );
      } else {
        return;
      }
    } catch (error) {
      setMessage(
        error instanceof Error
          ? `Не вдалося відкрити анкету: ${error.message}`
          : "Не вдалося відкрити анкету.",
      );
      return;
    }

    setPendingQuestionnaireFile(null);
    setDiskPreviewFile(null);
    setIsDiskFloatingPreview(false);
    setQuestionnairePreviewTitle(
      `Анкета · ${selectedSummary.name}${
        questionnaireExportFileName ? ` · ${questionnaireExportFileName}` : ""
      }`,
    );
    setQuestionnairePreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      return nextUrl;
    });
    setIsQuestionnairePreviewOpen(true);
  };

  const openQuestionnaireInNewTab = async () => {
    const externalId = selectedSummary.externalId;
    if (questionnaire?.fileData) {
      const url = dataUrlToObjectUrl(questionnaire.fileData);
      window.open(url, "_blank", "noopener,noreferrer");
      return;
    }
    if (
      externalId &&
      !pendingQuestionnaireFile &&
      !diskPreviewFile
    ) {
      try {
        const url = await api.createPersonQuestionnairePreviewUrl(
          externalId,
          questionnaireExportFileName,
        );
        window.open(url, "_blank", "noopener,noreferrer");
        return;
      } catch (error) {
        setMessage(
          error instanceof Error
            ? `Не вдалося відкрити анкету: ${error.message}`
            : "Не вдалося відкрити анкету.",
        );
        return;
      }
    }
    downloadCurrentQuestionnaire();
  };

  const beginQuestionnaireReview = (file: File | undefined) => {
    if (!file) return;
    if (!selectedSummary.externalId) {
      setMessage("Не вдалося додати анкету: у вибраної особи немає ID.");
      return;
    }
    if (file.type && file.type !== "application/pdf") {
      setMessage("Анкета має бути у форматі PDF.");
      return;
    }
    if (file.size > MAX_QUESTIONNAIRE_FILE_BYTES) {
      setMessage(
        `PDF завеликий для збереження в БД: ${formatFileSize(file.size)}. Максимум: ${formatFileSize(MAX_QUESTIONNAIRE_FILE_BYTES)}.`,
      );
      return;
    }

    const nextUrl = URL.createObjectURL(file);
    setPendingQuestionnaireFile(file);
    setDiskPreviewFile(null);
    setIsDiskFloatingPreview(false);
    setQuestionnairePreviewTitle(
      `Перегляд анкети перед збереженням · ${selectedSummary.name} · ${file.name}`,
    );
    setQuestionnairePreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      return nextUrl;
    });
    setIsQuestionnairePreviewOpen(true);
  };

  const uploadQuestionnaire = async (file: File) => {
    const externalId = selectedSummary.externalId;
    if (!externalId) {
      setMessage("Не вдалося зберегти анкету: у вибраної особи немає ID.");
      return;
    }

    setIsUploadingQuestionnaire(true);
    try {
      const exportFileName = sanitizeFileName(
        buildQuestionnaireExportFileName(
          selectedSummary.name,
          selectedCallSign,
        ),
      );
      const fileToSave = renameQuestionnaireFile(file, exportFileName);
      const saved = await api.upsertPersonQuestionnaireFile(
        externalId,
        fileToSave,
      );
      setQuestionnaire(saved);
      setQuestionnaireByExternalId((current) => ({
        ...current,
        [externalId]: true,
      }));
      notifyPersonnelAttachmentChanged(externalId, "questionnaire");
      setMessage(
        `Анкету збережено в БД: ${selectedSummary.name} · ${exportFileName}.`,
      );
      setPendingQuestionnaireFile(null);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося зберегти анкету в БД.",
      );
    } finally {
      setIsUploadingQuestionnaire(false);
    }
  };

  const confirmPendingQuestionnaire = async () => {
    if (!pendingQuestionnaireFile) return;
    await uploadQuestionnaire(pendingQuestionnaireFile);
    closeQuestionnairePreview();
  };

  const questionnaireCropFile = useMemo(() => {
    if (diskPreviewFile) return diskPreviewFile;
    if (pendingQuestionnaireFile) return pendingQuestionnaireFile;
    if (!questionnaire?.fileData) return null;
    return dataUrlToFile(
      questionnaire.fileData,
      questionnaire.fileName || "questionnaire.pdf",
    );
  }, [
    diskPreviewFile,
    pendingQuestionnaireFile,
    questionnaire?.fileData,
    questionnaire?.fileName,
  ]);

  return (
    <main className="main-panel personnel-page">
      <header className="topbar analytics-topbar personnel-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Особовий склад
          </Typography>
          <Typography
            className="personnel-topbar-hint"
            variant="body2"
            color="text.secondary"
          >
            Список із ЕЖООС · ручний файл «Штатка» (.xlsx) імпортується в БД.
          </Typography>
        </Box>
        <Stack className="personnel-topbar-actions" direction="row" spacing={1}>
          <Button
            variant="outlined"
            disabled={!canEdit || isLoading || isMergingAnketaData || isMergingVkTpvDovidky}
            onClick={() => void mergeMissingFieldsFromAnketaData()}
            title="Доповнити порожні поля з таблиці «Анкети» і додати осіб, яких ще немає в особовому складі"
          >
            {isMergingAnketaData ? "З анкет…" : "З анкетних даних"}
          </Button>
          <Button
            component="label"
            variant="outlined"
            disabled={!canEdit || isLoading || isMergingAnketaData || isMergingVkTpvDovidky}
            startIcon={<FileUploadOutlinedIcon />}
            title="Імпорт ВК № в ООС та ІПН в анкетні дані за ПІБ"
          >
            {isMergingVkTpvDovidky ? "ВК ТПВ…" : "ВК ТПВ ДОВІДКИ"}
            <input
              hidden
              type="file"
              accept=".xlsx,.xlsm"
              disabled={!canEdit}
              onChange={(event) => {
                void importVkTpvDovidkyWorkbook(event.target.files?.[0]);
                event.target.value = "";
              }}
            />
          </Button>
          <Button
            variant="outlined"
            disabled={!canEdit || !missingDiskSearchPeople.length}
            onClick={() => setIsDiskSearchOpen(true)}
          >
            Пошук усіх анкет
          </Button>
          <Button
            component="label"
            disabled={isLoading || !canEdit}
            startIcon={<FileUploadOutlinedIcon />}
            variant="outlined"
            title="Вибрати файл «Штатка» (.xlsx/.xlsm), імпортувати його та записати в БД"
          >
            Імпорт Штатки в БД
            <input
              hidden
              type="file"
              accept=".xlsx,.xlsm"
              disabled={!canEdit || isLoading}
              onChange={(event) => {
                void importPersonnelRosterWorkbook(event.target.files?.[0]);
                event.target.value = "";
              }}
            />
          </Button>
          <Button
            variant="outlined"
            disabled={!canEdit}
            onClick={() => void startPersonnelLoad({ force: true })}
          >
            Оновити з БД
          </Button>
        </Stack>
      </header>
      {isLoading || isMergingAnketaData || isMergingVkTpvDovidky ? (
        <LinearProgress color="primary" />
      ) : null}
      <Alert severity="info" variant="outlined" className="personnel-page-alert">
        {message}
      </Alert>

      <div className="personnel-mobile-tabs" role="tablist" aria-label="Розділи особового складу">
        <button
          type="button"
          role="tab"
          aria-selected={mobilePane === "list"}
          className={mobilePane === "list" ? "is-active" : undefined}
          onClick={() => setMobilePane("list")}
        >
          <FormatListBulletedOutlinedIcon fontSize="small" />
          Список
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={mobilePane === "card"}
          className={mobilePane === "card" ? "is-active" : undefined}
          disabled={!selectedRowId}
          onClick={() => selectedRowId && setMobilePane("card")}
        >
          <PersonOutlinedIcon fontSize="small" />
          Картка
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={mobilePane === "side"}
          className={mobilePane === "side" ? "is-active" : undefined}
          disabled={!selectedRowId}
          onClick={() => selectedRowId && setMobilePane("side")}
        >
          <ArticleOutlinedIcon fontSize="small" />
          Дії
        </button>
      </div>

      <section className={`personnel-layout mobile-pane-${mobilePane}`}>
        <aside className="analytics-panel personnel-list-panel">
          <div className="panel-heading personnel-list-heading">
            <span>Військовослужбовці · {filteredPersonnel.length}</span>
            <span className="personnel-questionnaire-summary">
              {questionnairePresenceStatus === "ready"
                ? `З анкетами · ${questionnaireCounts.withQuestionnaire}  /  Без анкет · ${questionnaireCounts.withoutQuestionnaire}`
                : questionnairePresenceStatus === "error"
                  ? "Анкети · не вдалося завантажити"
                  : "Анкети · завантаження…"}
            </span>
          </div>
          <label className="personnel-search">
            <SearchOutlinedIcon fontSize="small" />
            <input
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="ПІБ, позивний, звання, посада"
            />
          </label>
          <div
            aria-label="Фільтр за штаткою"
            className="personnel-questionnaire-filter"
            role="group"
          >
            {(
              [
                ["all", "Усі", staffCounts.all],
                ["in", "У штаті", staffCounts.in],
                ["archive", "Архів", staffCounts.archive],
              ] as const
            ).map(([value, label, count]) => (
              <button
                aria-pressed={staffFilter === value}
                className={staffFilter === value ? "is-active" : undefined}
                key={value}
                title={
                  value === "in"
                    ? "Лише особи з актуального «Загального списку» Штатки"
                    : value === "archive"
                      ? "Лише особи, додані з аркуша «Архів»"
                      : "Усі особи"
                }
                onClick={() => {
                  setStaffFilter(value);
                  setSelectedRowId("");
                }}
                type="button"
              >
                {label} · {count}
              </button>
            ))}
          </div>
          {isLoading && personnelRows.length === 0 ? (
            <div className="personnel-list-preloader" role="status">
              <LinearProgress color="primary" />
              <strong>Готую список «У штаті»…</strong>
              <span>Спочатку завантажую Штатку, тому проміжний список ООС не показується.</span>
            </div>
          ) : (
            <PersonnelVirtualList
              items={filteredPersonnel}
              selectedRowId={selectedRowId}
              photoByExternalId={photoByExternalId}
              onNeedPhotos={loadVisiblePersonnelPhotos}
              onSelect={(rowId) => {
                setSelectedRowId(rowId);
                setMobilePane("card");
              }}
              keyboardEnabled={
                !isPhotoCropOpen &&
                !isQuestionnairePreviewOpen &&
                !isPhotoLightboxOpen &&
                !isDiskSearchOpen &&
                !activePersonAction
              }
            />
          )}
          {filteredPersonnel.length === 0 && query.trim() ? (
            <p className="personnel-empty-hint">
              {staffFilter === "in"
                ? "Нічого у «У штаті». Спробуйте «Усі» або «Імпорт Штатки в БД»."
                : staffFilter === "archive"
                  ? "Нічого в «Архіві». Імпортуйте Штатку з аркушем «Архів»."
                  : "Нічого не знайдено. Спробуйте «Імпорт Штатки в БД»."}
            </p>
          ) : null}
        </aside>

        <section className="person-card-panel">
          <div className="person-card-hero">
            <div className="person-avatar">
              {selectedPhoto ? (
                <img
                  alt={selectedSummary.name}
                  src={selectedPhoto}
                  onClick={() => setIsPhotoLightboxOpen(true)}
                />
              ) : (
                <PersonSearchOutlinedIcon />
              )}
              <button
                aria-label="Відкрити анкету"
                className="person-avatar-zoom"
                disabled={!selectedRowId}
                onClick={() => {
                  if (!questionnaire?.fileData) {
                    setMessage("Анкета ще не додана.");
                    return;
                  }
                  void openQuestionnairePreview();
                }}
                title="Відкрити анкету"
                type="button"
              >
                <SearchOutlinedIcon />
              </button>
              {selectedPhoto ? (
                <button
                  aria-label="Видалити фото"
                  className="person-avatar-delete"
                  disabled={!selectedRow}
                  onClick={() => void deleteSelectedPhoto()}
                  title="Видалити фото"
                  type="button"
                >
                  <DeleteOutlineOutlinedIcon />
                </button>
              ) : null}
              <Button
                aria-label="Додати фото"
                className="person-avatar-upload"
                component="label"
                disabled={!selectedRow}
                size="small"
                startIcon={<AddPhotoAlternateOutlinedIcon />}
                title="Додати фото"
                variant="contained"
                sx={{ color: "#1a1a14" }}
              >
                Фото
                <input
                  hidden
                  type="file"
                  accept="image/*,application/pdf"
                  onChange={(event) => {
                    openPhotoCrop(event.target.files?.[0]);
                    event.target.value = "";
                  }}
                />
              </Button>
            </div>
            <div className="person-card-identity">
              {selectedCallSign || rosterStatus ? (
                <div className="person-callsign-row">
                  {selectedCallSign ? (
                    <span className="person-callsign" title="Позивний">
                      <span className="person-callsign-label">позивний</span>
                      <strong>{selectedCallSign}</strong>
                    </span>
                  ) : null}
                  {rosterStatus ? (
                    <span className="person-roster-status" title="Статус">
                      <span className="person-callsign-label">статус</span>
                      <strong>{rosterStatus}</strong>
                    </span>
                  ) : null}
                </div>
              ) : null}
              <PersonCardName name={selectedSummary.name} />
              <div className="person-action-tags">
                {selectedSummary.rank && (
                  <Chip label={selectedSummary.rank} size="small" />
                )}
                {selectedSummary.positionIndex && (
                  <Chip
                    label={`Посада: ${selectedSummary.positionIndex}`}
                    size="small"
                  />
                )}
                {selectedSummary.serviceType && (
                  <Chip label={selectedSummary.serviceType} size="small" />
                )}
              </div>
            </div>
          </div>

          <div className="person-card-scroll">
            <div className="person-action-fields">
              <span className="wide person-phones-summary">
                <strong>Телефони</strong>
                {parsedPhones.length > 0 ? (
                  <span className="person-phones-list">
                    {parsedPhones.map((phone) => {
                      const isSaved = savedPhones.includes(phone);
                      return (
                        <span className="person-phone-chip" key={phone}>
                          <Chip
                            label={formatUaPhoneDisplay(phone)}
                            size="small"
                            color="primary"
                            variant="outlined"
                          />
                          {isSaved ? (
                            <button
                              aria-label={`Видалити ${formatUaPhoneDisplay(phone)}`}
                              disabled={isSavingPhone}
                              type="button"
                              onClick={() => void removeSelectedPersonPhone(phone)}
                            >
                              <DeleteOutlineOutlinedIcon fontSize="small" />
                            </button>
                          ) : null}
                        </span>
                      );
                    })}
                  </span>
                ) : (
                  <span className="person-phones-empty">Номерів ще немає</span>
                )}
                <div className="person-phones-editor">
                  <input
                    aria-label="Номер телефону"
                    autoComplete="off"
                    disabled={!selectedRow || isSavingPhone}
                    inputMode="tel"
                    placeholder="063 123 45 67"
                    value={phoneDraft}
                    onChange={(event) => setPhoneDraft(event.target.value)}
                    onKeyDown={(event) => {
                      if (event.key !== "Enter") return;
                      event.preventDefault();
                      void addSelectedPersonPhone();
                    }}
                  />
                  <Button
                    disabled={!selectedRow || isSavingPhone}
                    size="small"
                    type="button"
                    variant="contained"
                    sx={{ color: "#1a1a14" }}
                    onClick={() => void addSelectedPersonPhone()}
                  >
                    Додати
                  </Button>
                </div>
              </span>
              <span>
                <strong>РНОКПП</strong>
                {selectedSummary.rnokpp || "—"}
              </span>
              <span>
                <strong>Дата народження</strong>
                {birthDateWithAge || "—"}
              </span>
              <span className="person-location-highlight">
                <strong>
                  <LocationOnOutlinedIcon fontSize="small" />
                  Поточне місцеперебування
                </strong>
                <span className="person-location-value">
                  {selectedSummary.location || "Не вказано"}
                </span>
              </span>
              <span>
                <strong>Посада</strong>
                {selectedSummary.positionTitle || "—"}
              </span>
              <span>
                <strong>Військовий квиток</strong>
                {selectedSummary.militaryId || "—"}
              </span>
              <span>
                <strong>Звідки прибув</strong>
                {selectedSummary.arrivedFrom || "—"}
              </span>
            </div>

            {fighterStatusFieldRows.length > 0 ? (
              <div className="person-edit-section">
                <div className="panel-heading">Статус бійців</div>
                <div className="person-roster-grid">
                  {fighterStatusFieldRows.map((field) => {
                    const tone = getFighterStatusFieldTone(field.key);
                    const label = field.label.replace(/^Статус бійців · /, "");
                    const Icon =
                      tone === "exit"
                        ? LogoutOutlinedIcon
                        : tone === "return"
                          ? LoginOutlinedIcon
                          : tone === "entry"
                            ? CalendarMonthOutlinedIcon
                            : tone === "days"
                              ? InfoOutlinedIcon
                              : tone === "status"
                                ? WarningAmberOutlinedIcon
                                : tone === "direction"
                                  ? PushPinOutlinedIcon
                                  : null;
                    return (
                      <span
                        key={field.key}
                        className={
                          tone
                            ? `person-roster-tile is-${tone}`
                            : "person-roster-tile"
                        }
                      >
                        <strong>
                          {Icon ? <Icon fontSize="small" aria-hidden /> : null}
                          {label}
                        </strong>
                        <em>{field.value}</em>
                      </span>
                    );
                  })}
                </div>
              </div>
            ) : null}

            {rosterFieldRows.length > 0 ? (
              <div className="person-edit-section">
                <div className="panel-heading">Загальний список</div>
                <div className="person-roster-grid">
                  {rosterFieldRows.map((field) => (
                    <span key={field.key}>
                      <strong>{field.label}</strong>
                      {field.value}
                    </span>
                  ))}
                </div>
              </div>
            ) : null}

            {editableFieldsBySection.map((group) => (
              <div className="person-edit-section" key={group.section}>
                <div className="panel-heading">{group.label}</div>
                <div className="person-edit-grid">
                  {group.fields.map((field) => {
                    const isWide =
                      field.kind === "multiline" ||
                      field.section === "contacts" ||
                      field.parts.includes("додаткова_інформація");
                    const isBirthDateField =
                      field.parts.includes("дата_народження");
                    const years = isBirthDateField
                      ? computeFullYearsFromBirthDate(
                          editValues[field.key] || selectedSummary.birthDate,
                        )
                      : null;

                    return (
                      <label className={isWide ? "wide" : ""} key={field.key}>
                        <span>
                          {field.label}
                          {years != null ? ` · ${years} р.` : ""}
                        </span>
                        {field.kind === "multiline" ? (
                          <textarea
                            className="sci-message-area"
                            value={editValues[field.key] ?? ""}
                            onChange={(event) =>
                              setEditValues((values) => ({
                                ...values,
                                [field.key]: event.target.value,
                              }))
                            }
                            onBlur={() => void saveSelectedPerson()}
                          />
                        ) : (
                          <input
                            value={editValues[field.key] ?? ""}
                            onChange={(event) =>
                              setEditValues((values) => ({
                                ...values,
                                [field.key]: event.target.value,
                              }))
                            }
                            onBlur={() => void saveSelectedPerson()}
                          />
                        )}
                      </label>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        </section>

        <aside className="person-side-panel">
          <div className="personnel-mobile-card-nav">
            <Button
              size="small"
              variant="outlined"
              startIcon={<ArrowLeftOutlinedIcon fontSize="small" />}
              onClick={() => setMobilePane("card")}
            >
              До картки
            </Button>
            <Button
              size="small"
              variant="outlined"
              onClick={() => setMobilePane("list")}
            >
              До списку
            </Button>
          </div>
          <div className="analytics-panel">
            <div className="panel-heading">Зміна статусу</div>
            <div className="person-action-buttons">
              {personActions.map((action) => (
                <Button
                  key={action.type}
                  disabled={!selectedRow}
                  variant={
                    activePersonAction?.type === action.type
                      ? "contained"
                      : "outlined"
                  }
                  size="small"
                  onClick={() => {
                    setActivePersonAction(action);
                    setPersonActionForm(createDefaultActionForm());
                  }}
                  sx={
                    activePersonAction?.type === action.type
                      ? { color: "#1a1a14" }
                      : undefined
                  }
                >
                  {action.label}
                </Button>
              ))}
            </div>
          </div>
          <div className="analytics-panel">
            <div className="person-documents-header">
              <div className="panel-heading">Пов’язані документи</div>
              <Button
                component="label"
                disabled={
                  !selectedRow ||
                  !selectedSummary.externalId ||
                  isUploadingQuestionnaire
                }
                size="small"
                variant="outlined"
                startIcon={<PictureAsPdfOutlinedIcon />}
              >
                Додати анкету
                <input
                  hidden
                  type="file"
                  accept="application/pdf,.pdf"
                  onChange={(event) => {
                    beginQuestionnaireReview(event.target.files?.[0]);
                    event.target.value = "";
                  }}
                />
              </Button>
            </div>
            <div className="person-document-list">
              {questionnaire ? (
                <article
                  className={[
                    "person-document-shell",
                    "is-ready",
                  ].join(" ")}
                >
                  <button
                    className="person-document-item is-ready"
                    type="button"
                    onClick={() => void openQuestionnairePreview()}
                  >
                    <PictureAsPdfOutlinedIcon />
                    <span>
                      <strong>Анкета (PDF)</strong>
                      <small>
                        {questionnaireExportFileName} · переглянути
                      </small>
                    </span>
                  </button>
                  <div className="person-document-actions">
                    <button
                      aria-label="Експорт анкети"
                      className="person-document-delete"
                      disabled={!selectedRow}
                      onClick={() => void exportCurrentQuestionnaire()}
                      title={`Експорт: ${questionnaireExportFileName}`}
                      type="button"
                    >
                      <FileDownloadOutlinedIcon />
                    </button>
                    <button
                      aria-label="Показати анкету у Finder"
                      className="person-document-delete person-document-action--finder"
                      disabled={!selectedRow}
                      onClick={() => void revealCurrentQuestionnaireInFinder()}
                      title="Показати оригінал у Finder"
                      type="button"
                    >
                      <SearchOutlinedIcon />
                    </button>
                    <button
                      aria-label="Видалити анкету"
                      className="person-document-delete"
                      disabled={!selectedRow}
                      onClick={() => void deleteSelectedQuestionnaire()}
                      title="Видалити анкету"
                      type="button"
                    >
                      <DeleteOutlineOutlinedIcon />
                    </button>
                  </div>
                </article>
              ) : (
                <div className="person-document-empty">
                  <PictureAsPdfOutlinedIcon />
                  <span>Анкета ще не додана</span>
                </div>
              )}
              {["Довідка про проходження служби", "Витяг з наказу", "Рапорт"].map(
                (item) => (
                  <div key={item}>
                    <ArticleOutlinedIcon />
                    <span>{item}</span>
                  </div>
                ),
              )}
              <button
                className={[
                  "person-document-item",
                  personRelatedDocuments.some(
                    (document) => document.type === "salaryPowerAttorney",
                  )
                    ? "is-ready"
                    : "",
                ]
                  .filter(Boolean)
                  .join(" ")}
                disabled={!selectedRow}
                type="button"
                onClick={() =>
                  selectedRow &&
                  onOpenDocuments(selectedRow, "salaryPowerAttorney")
                }
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>Довіреність зарплати</strong>
                  <small>створити документ і вести прогрес</small>
                </span>
              </button>
              <button
                className={[
                  "person-document-item",
                  personRelatedDocuments.some(
                    (document) => document.type === "ubdReport",
                  )
                    ? "is-ready"
                    : "",
                ]
                  .filter(Boolean)
                  .join(" ")}
                disabled={!selectedRow}
                type="button"
                onClick={() =>
                  selectedRow &&
                  onOpenDocuments(selectedRow, "ubdReport")
                }
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>Рапорт на УБД</strong>
                  <small>рапорт, скани документів, статус</small>
                </span>
              </button>
              <button
                className={[
                  "person-document-item",
                  personRelatedDocuments.some(
                    (document) => document.type === "ubdRestoreReport",
                  )
                    ? "is-ready"
                    : "",
                ]
                  .filter(Boolean)
                  .join(" ")}
                disabled={!selectedRow}
                type="button"
                onClick={() =>
                  selectedRow && onOpenDocuments(selectedRow, "ubdRestoreReport")
                }
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>Рапорт на відновлення УБД</strong>
                  <small>пошкоджене посвідчення, клопотання, скани</small>
                </span>
              </button>
              <button
                className={[
                  "person-document-item",
                  personRelatedDocuments.some(
                    (document) => document.type === "form6Report",
                  )
                    ? "is-ready"
                    : "",
                ]
                  .filter(Boolean)
                  .join(" ")}
                disabled={!selectedRow}
                type="button"
                onClick={() =>
                  selectedRow && onOpenDocuments(selectedRow, "form6Report")
                }
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>Форма 6</strong>
                  <small>рапорт для довідки УБД, персональні дані, скани</small>
                </span>
              </button>
              <button
                className={[
                  "person-document-item",
                  personRelatedDocuments.some(
                    (document) => document.type === "form12Report",
                  )
                    ? "is-ready"
                    : "",
                ]
                  .filter(Boolean)
                  .join(" ")}
                disabled={!selectedRow}
                type="button"
                onClick={() =>
                  selectedRow && onOpenDocuments(selectedRow, "form12Report")
                }
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>Форма 12</strong>
                  <small>рапорт Ф-12, дані бійця, підпис PNG</small>
                </span>
              </button>
              <button
                className={[
                  "person-document-item",
                  personRelatedDocuments.some(
                    (document) => document.type === "serviceCharacteristic",
                  )
                    ? "is-ready"
                    : "",
                ]
                  .filter(Boolean)
                  .join(" ")}
                disabled={!selectedRow}
                type="button"
                onClick={() =>
                  selectedRow &&
                  onOpenDocuments(selectedRow, "serviceCharacteristic")
                }
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>Службова характеристика</strong>
                  <small>звання, ПІБ, посада, текст, підпис командира</small>
                </span>
              </button>
              <button
                className={[
                  "person-document-item",
                  personRelatedDocuments.some(
                    (document) => document.type === "zhbdCertificate",
                  )
                    ? "is-ready"
                    : "",
                ]
                  .filter(Boolean)
                  .join(" ")}
                disabled={!selectedRow}
                type="button"
                onClick={() => {
                  if (!selectedRow) return;
                  const fullPosition =
                    rosterFieldRows.find((field) =>
                      field.label
                        .trim()
                        .toLocaleLowerCase("uk-UA")
                        .replace(/_/g, " ")
                        .includes("повна посада"),
                    )?.value ||
                    pickFullPositionFromPersonRow(selectedRow) ||
                    rosterFieldRows.find((field) =>
                      field.label.trim().toLocaleLowerCase("uk-UA").replace(/_/g, " ") ===
                      "посада",
                    )?.value ||
                    "";
                  onOpenDocuments(selectedRow, "zhbdCertificate", {
                    fullPosition,
                  });
                }}
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>Довідка ЖБД</strong>
                  <small>період, посада, підстава, підпис</small>
                </span>
              </button>
              <button
                className={[
                  "person-document-item",
                  personRelatedDocuments.some(
                    (document) => document.type === "temporaryMilitaryId",
                  )
                    ? "is-ready"
                    : "",
                ]
                  .filter(Boolean)
                  .join(" ")}
                disabled={!selectedRow}
                type="button"
                onClick={() =>
                  selectedRow &&
                  onOpenDocuments(selectedRow, "temporaryMilitaryId")
                }
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>Тимчасовий військовий квиток</strong>
                  <small>фото, рядок для замовлення, прогрес</small>
                </span>
              </button>
              <button
                className={[
                  "person-document-item",
                  personRelatedDocuments.some(
                    (document) => document.type === "lostMilitaryId",
                  )
                    ? "is-ready"
                    : "",
                ]
                  .filter(Boolean)
                  .join(" ")}
                disabled={!selectedRow}
                type="button"
                onClick={() =>
                  selectedRow && onOpenDocuments(selectedRow, "lostMilitaryId")
                }
              >
                <ArticleOutlinedIcon />
                <span>
                  <strong>Втрата військового квитка</strong>
                  <small>рапорт, наказ, акт розслідування</small>
                </span>
              </button>
            </div>
          </div>
        </aside>
      </section>

      {activePersonAction && (
        <section className="person-action-panel">
          <div className="person-action-form">
            <div className="panel-heading">{activePersonAction.label}</div>
            <label>
              <span>Дата від</span>
              <input
                type="date"
                value={personActionForm.validFrom}
                onChange={(event) =>
                  setPersonActionForm((form) => ({
                    ...form,
                    validFrom: event.target.value,
                  }))
                }
              />
            </label>
            <label>
              <span>Дата до</span>
              <input
                type="date"
                value={personActionForm.validTo}
                onChange={(event) =>
                  setPersonActionForm((form) => ({
                    ...form,
                    validTo: event.target.value,
                  }))
                }
              />
            </label>
            <label>
              <span>Причина / статус</span>
              <input
                value={personActionForm.reason}
                onChange={(event) =>
                  setPersonActionForm((form) => ({
                    ...form,
                    reason: event.target.value,
                  }))
                }
                placeholder={activePersonAction.label}
              />
            </label>
            <label>
              <span>Місце</span>
              <input
                value={personActionForm.place}
                onChange={(event) =>
                  setPersonActionForm((form) => ({
                    ...form,
                    place: event.target.value,
                  }))
                }
                placeholder="Місце / підрозділ / заклад"
              />
            </label>
            {activePersonAction.type === "POSITION_CHANGE" && (
              <>
                <label>
                  <span>Індекс посади</span>
                  <input
                    value={personActionForm.positionIndex}
                    onChange={(event) =>
                      setPersonActionForm((form) => ({
                        ...form,
                        positionIndex: event.target.value,
                      }))
                    }
                    placeholder={selectedSummary.positionIndex || "0600000"}
                  />
                </label>
                <label>
                  <span>Назва посади</span>
                  <input
                    value={personActionForm.positionTitle}
                    onChange={(event) =>
                      setPersonActionForm((form) => ({
                        ...form,
                        positionTitle: event.target.value,
                      }))
                    }
                    placeholder="Нова посада"
                  />
                </label>
              </>
            )}
            {activePersonAction.type === "RANK_CHANGE" && (
              <label>
                <span>Нове звання</span>
                <input
                  value={personActionForm.rank}
                  onChange={(event) =>
                    setPersonActionForm((form) => ({
                      ...form,
                      rank: event.target.value,
                    }))
                  }
                  placeholder={selectedSummary.rank || "солдат"}
                />
              </label>
            )}
            <label className="person-action-note">
              <span>Примітка</span>
              <textarea
                className="sci-message-area"
                value={personActionForm.note}
                onChange={(event) =>
                  setPersonActionForm((form) => ({
                    ...form,
                    note: event.target.value,
                  }))
                }
                placeholder="Наказ, уточнення або коментар"
              />
            </label>
            <div className="person-action-submit">
              <Button
                variant="outlined"
                onClick={() => setActivePersonAction(null)}
              >
                Скасувати
              </Button>
              <Button
                variant="contained"
                onClick={() => void submitPersonAction()}
                sx={{ color: "#1a1a14" }}
              >
                Зберегти дію
              </Button>
            </div>
          </div>
        </section>
      )}

      {isPhotoLightboxOpen && selectedPhoto ? (
        <div
          className="person-photo-lightbox"
          role="dialog"
          aria-modal="true"
          aria-label={`Фото · ${selectedSummary.name}`}
        >
          <button
            aria-label="Закрити фото"
            className="person-photo-lightbox-backdrop"
            onClick={() => setIsPhotoLightboxOpen(false)}
            type="button"
          />
          <img alt={selectedSummary.name} src={selectedPhoto} />
          <button
            className="person-photo-lightbox-close"
            onClick={() => setIsPhotoLightboxOpen(false)}
            type="button"
          >
            Закрити
          </button>
        </div>
      ) : null}

      <PhotoCropDialog
        file={photoCropFile}
        open={isPhotoCropOpen}
        floating={isDiskFloatingCrop}
        onClose={() => {
          setIsPhotoCropOpen(false);
          setIsDiskFloatingCrop(false);
        }}
        onMessage={setMessage}
        onSave={savePersonPhoto}
      />

      <QuestionnaireDiskSearchDialog
        open={isDiskSearchOpen}
        people={missingDiskSearchPeople}
        onClose={() => setIsDiskSearchOpen(false)}
        onConfirmed={(externalId) => {
          setQuestionnaireByExternalId((current) => ({
            ...current,
            [externalId]: true,
          }));
          focusPersonByExternalId(externalId);
          notifyPersonnelAttachmentChanged(externalId, "questionnaire");
          void api.getPersonQuestionnaire(externalId).then((next) => {
            setQuestionnaire(next);
          }).catch(() => undefined);
          setMessage(`Анкету підтверджено та збережено для ID ${externalId}.`);
        }}
        onAutoPhotoSaved={(externalId, photoData) => {
          setPhotoByExternalId((current) => ({
            ...current,
            [externalId]: photoData,
          }));
          if (Date.now() < holdStatusUntilRef.current) return;
          setMessage(`Фото автоматично знайдено в PDF і додано до preview для ID ${externalId}.`);
        }}
        onPreviewQuestionnaire={(file, title, externalId) => {
          focusPersonByExternalId(externalId);
          openDiskQuestionnairePreview(file, title);
        }}
        onCropPhoto={openDiskPhotoCrop}
      />

      <FloatingQuestionnairePreview
        open={isQuestionnairePreviewOpen}
        title={
          questionnairePreviewTitle ||
          `Анкета · ${selectedSummary.name}`
        }
        previewUrl={questionnairePreviewUrl}
        pendingFile={Boolean(pendingQuestionnaireFile)}
        isUploading={isUploadingQuestionnaire}
        placement={isDiskFloatingPreview ? "left" : "center"}
        defaultWidth={isDiskFloatingPreview ? 560 : 760}
        defaultHeight={isDiskFloatingPreview ? 720 : 820}
        cropFile={questionnaireCropFile}
        onClose={closeQuestionnairePreview}
        onSaveCrop={savePersonPhoto}
        onCropMessage={setMessage}
        onOpenTab={openQuestionnaireInNewTab}
        onDownload={() => void exportCurrentQuestionnaire()}
        onSave={() => void confirmPendingQuestionnaire()}
        shareFileName={questionnaireExportFileName}
        sharePersonName={selectedSummary.name}
        shareSource={currentQuestionnaireShareSource}
        onShareNotify={setMessage}
      />
    </main>
  );
}
