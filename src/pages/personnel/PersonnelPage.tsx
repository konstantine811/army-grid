import { useEffect, useMemo, useState } from "react";
import {
  Alert,
  Box,
  Button,
  Chip,
  Dialog,
  DialogActions,
  DialogContent,
  DialogTitle,
  LinearProgress,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { AddPhotoAlternateOutlinedIcon } from "@/components/sci/icons";
import { ArticleOutlinedIcon } from "@/components/sci/icons";
import { DeleteOutlineOutlinedIcon } from "@/components/sci/icons";
import { FileUploadOutlinedIcon } from "@/components/sci/icons";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import { PersonSearchOutlinedIcon } from "@/components/sci/icons";
import { PictureAsPdfOutlinedIcon } from "@/components/sci/icons";
import { SearchOutlinedIcon } from "@/components/sci/icons";
import {
  api,
  type BackendEjournalImport,
  type BackendPersonDocument,
  type BackendPersonQuestionnaire,
} from "../../api";
import {
  buildFighterStatusAdditions,
  extractFighterStatusFieldRows,
  FIGHTER_STATUS_FIELDS,
  findFighterStatusAddition,
  findFighterStatusSheet,
  normalizeRosterMatchText,
} from "./fighterStatusImport";
import {
  hasRowData,
  readWorkbookSnapshot,
  valueToDisplay,
} from "../../excelRoundTrip";
import { cellValueToJson } from "../../shared/format";
import type { DbPreviewState, EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { buildImportColumns, parseDbColumns } from "../ejournal/ejournalUtils";
import { PhotoCropDialog, type CropRect } from "./PhotoCropDialog";
import { FloatingQuestionnairePreview } from "./FloatingQuestionnairePreview";
import { QuestionnaireShareButton } from "./QuestionnaireShareButton";
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
  buildSelfAttachmentMigrationPairs,
  buildOrphanAttachmentMigrationPairs,
  collectPersonExternalIdCandidates,
  getPersonDisplayName,
  getPersonExternalId,
  getPersonFieldValue,
  inferRosterFieldLabel,
  isLikelyPersonnelRow,
  migratePersonAttachmentsBetweenIds,
  normalizePersonBirthKey,
  normalizeUaPhone,
  resolveMorningGeneralListColumnLabel,
  resolvePersonIdentityKey,
  resolvePersonRankTitle,
  isPositionIndexField,
  loadAllEjournalSheetRows,
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
  applyEnrichmentToPreviewRow,
  syncEnrichmentToPerson,
} from "./personEnrichment";
import { migrateStoredPersonSignatures } from "./personSignatureStore";
import { parseQuestionnairePdf } from "../questionnaire-parser/questionnairePdfParser";

const PERSONNEL_FOCUS_KEY = "army-grid:focus-personnel";
const ATTACHMENT_HEAL_SESSION_KEY = "army-grid:attachments-healed";
const MAX_QUESTIONNAIRE_FILE_BYTES = 350 * 1024 * 1024;

const formatFileSize = (bytes: number) => {
  if (!Number.isFinite(bytes)) return "";
  const mb = bytes / 1024 / 1024;
  return `${mb.toFixed(mb >= 10 ? 0 : 1)} MB`;
};

const ROSTER_FIELD_PREFIX = "roster__";
const normalizeRosterText = normalizeRosterMatchText;

const getRosterValue = (row: EjournalPreviewRow, keyParts: string[]) => {
  const key = Object.keys(row).find((item) =>
    keyParts.every((part) => item.toLocaleLowerCase("uk-UA").includes(part)),
  );
  return key ? valueToDisplay(row[key] as Parameters<typeof valueToDisplay>[0]).trim() : "";
};

const normalizePersonnelSearchText = (value: unknown) =>
  valueToDisplay(value as Parameters<typeof valueToDisplay>[0])
    .replace(/[ʼ’']/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const getRawCallSignSearchValues = (row: EjournalPreviewRow) =>
  collectPersonCallSignFieldValues(row);

const getRosterAdditions = (rosterRow: EjournalPreviewRow) =>
  Object.fromEntries(
    Object.entries(rosterRow)
      .filter(([key, value]) =>
        !key.startsWith("__") &&
        valueToDisplay(value as Parameters<typeof valueToDisplay>[0]).trim(),
      )
      .map(([key, value]) => [`${ROSTER_FIELD_PREFIX}${key}`, value]),
  );

const buildRosterOnlyPersonnelRow = (rosterRow: EjournalPreviewRow) => {
  const name =
    getRosterValue(rosterRow, ["піб"]) ||
    getRosterValue(rosterRow, ["прізвище"]);
  const identityKey = resolvePersonIdentityKey({
    ...rosterRow,
    прізвище: name,
    ПІБ: name,
  });
  const rowKey = identityKey || normalizeRosterText(name);

  return {
    __dbRowId: `roster:${rowKey}`,
    __rowNumber: rosterRow.__rowNumber,
    id: identityKey,
    "прізвище": name,
    "ПІБ": name,
    "звання": resolvePersonRankTitle(rosterRow),
    "Звання": resolvePersonRankTitle(rosterRow),
    "позивний": getRosterValue(rosterRow, ["позив"]),
    "Позивний": getRosterValue(rosterRow, ["позив"]),
    "індекс_посади": getRosterValue(rosterRow, ["індекс", "посади"]),
    "Індекс посади": getRosterValue(rosterRow, ["індекс", "посади"]),
    "місце_дислокації": getRosterValue(rosterRow, ["перебування"]),
    "Місце дислокації": getRosterValue(rosterRow, ["перебування"]),
    ...getRosterAdditions(rosterRow),
  } as EjournalPreviewRow;
};

const mergeRosterRowsIntoPreview = (
  preview: DbPreviewState,
  rosterRows: EjournalPreviewRow[],
) => {
  if (!rosterRows.length) return preview;

  const rosterById = new Map<string, EjournalPreviewRow>();
  const rosterByName = new Map<string, EjournalPreviewRow>();
  const usedRosterRows = new Set<EjournalPreviewRow>();
  rosterRows.forEach((row) => {
    const id = getPersonExternalId(row);
    const name = getRosterValue(row, ["піб"]) || getRosterValue(row, ["прізвище"]);
    if (id) rosterById.set(id, row);
    if (name) rosterByName.set(normalizeRosterText(name), row);
  });

  const mergedRows = preview.rows.map((row) => {
    try {
      const spreadsheetId = getPersonExternalId(row);
      const name = getPersonDisplayName(row);
      const rosterRow =
        (spreadsheetId && rosterById.get(spreadsheetId)) ||
        rosterByName.get(normalizeRosterText(name));

      if (!rosterRow) return row;
      usedRosterRows.add(rosterRow);
      return { ...row, ...getRosterAdditions(rosterRow) };
    } catch {
      return row;
    }
  });
  const rosterOnlyRows = rosterRows
    .filter((row) => !usedRosterRows.has(row))
    .filter((row) => getRosterValue(row, ["піб"]) || getRosterValue(row, ["прізвище"]))
    .flatMap((row) => {
      try {
        return [buildRosterOnlyPersonnelRow(row)];
      } catch {
        return [];
      }
    });

  return {
    ...preview,
    rows: [...mergedRows, ...rosterOnlyRows],
  };
};

const personRecordMatchKeys = (record: PersonnelRecord) => {
  const nameKey = normalizeRosterText(record.summary.name);
  const birthKey = normalizePersonBirthKey(record.summary.birthDate);
  const callSignKey = normalizeRosterText(record.summary.callSign);
  const keys: string[] = [];
  if (nameKey && birthKey) keys.push(`name-birth:${nameKey}:${birthKey}`);
  if (callSignKey) keys.push(`call:${callSignKey}`);
  if (nameKey) keys.push(`name:${nameKey}`);
  return keys;
};

const mapUniqueRecordsByMatchKey = (
  records: PersonnelRecord[],
  prefix: string,
) => {
  const buckets = new Map<string, PersonnelRecord[]>();
  records.forEach((record) => {
    if (!record.summary.externalId) return;
    for (const key of personRecordMatchKeys(record)) {
      if (!key.startsWith(prefix)) continue;
      buckets.set(key, [...(buckets.get(key) ?? []), record]);
    }
  });

  return new Map(
    [...buckets.entries()]
      .filter(([, items]) => items.length === 1)
      .map(([key, items]) => [key, items[0] as PersonnelRecord]),
  );
};

const buildAttachmentMigrationPairs = (
  previousRecords: PersonnelRecord[],
  nextRows: EjournalPreviewRow[],
) => {
  const nextRecords = nextRows
    .filter(isLikelyPersonnelRow)
    .map((row) => ({ row, summary: buildPersonSummary(row) }));

  const usedPrevious = new Set<string>();
  const usedNext = new Set<string>();
  const pairs: Array<{
    name: string;
    fromExternalId: string;
    toExternalId: string;
  }> = [];

  const matchByPrefix = (prefix: string) => {
    const previousByKey = mapUniqueRecordsByMatchKey(previousRecords, prefix);
    const nextByKey = mapUniqueRecordsByMatchKey(nextRecords, prefix);
    for (const [key, nextRecord] of nextByKey.entries()) {
      const previousRecord = previousByKey.get(key);
      if (!previousRecord) continue;
      const fromExternalId = previousRecord.summary.externalId;
      const toExternalId = nextRecord.summary.externalId;
      if (
        !fromExternalId ||
        !toExternalId ||
        fromExternalId === toExternalId ||
        usedPrevious.has(fromExternalId) ||
        usedNext.has(toExternalId)
      ) {
        continue;
      }
      usedPrevious.add(fromExternalId);
      usedNext.add(toExternalId);
      pairs.push({
        name: nextRecord.summary.name,
        fromExternalId,
        toExternalId,
      });
      for (const candidateId of collectPersonExternalIdCandidates(
        previousRecord.row,
      )) {
        pairs.push({
          name: nextRecord.summary.name,
          fromExternalId: candidateId,
          toExternalId,
        });
      }
    }
  };

  matchByPrefix("name-birth:");
  matchByPrefix("call:");
  matchByPrefix("name:");

  return pairs;
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
      | "ubdRestoreReport"
      | "temporaryMilitaryId",
  ) => void;
}) {
  const [imports, setImports] = useState<BackendEjournalImport[]>([]);
  const [dbPreview, setDbPreview] = useState<DbPreviewState | null>(null);
  const [rosterLabels, setRosterLabels] = useState<Record<string, string>>({});
  const [rosterImportName, setRosterImportName] = useState("");
  const [selectedRowId, setSelectedRowId] = useState("");
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
  const [isPullingFromQuestionnaire, setIsPullingFromQuestionnaire] =
    useState(false);
  const [questionnaireByExternalId, setQuestionnaireByExternalId] = useState<
    Record<string, true>
  >({});
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
  const [message, setMessage] = useState(`API: ${api.baseUrl}`);
  const [isLoading, setIsLoading] = useState(false);
  const personnelRows = useMemo<PersonnelRecord[]>(
    () =>
      (dbPreview?.rows ?? [])
        .filter(isLikelyPersonnelRow)
        .map((row) => ({ row, summary: buildPersonListSummary(row) })),
    [dbPreview],
  );
  const filteredPersonnel = useMemo(() => {
    const normalizedQuery = normalizePersonnelSearchText(query);

    if (!normalizedQuery) return personnelRows;

    return personnelRows.filter((record) => {
      const searchableText = normalizePersonnelSearchText(
        [
        record.summary.name,
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

      return searchableText.includes(normalizedQuery);
    });
  }, [personnelRows, phonesByExternalId, query]);
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
  const questionnaireExportFileName = useMemo(
    () =>
      sanitizeFileName(
        buildQuestionnaireExportFileName(selectedSummary.name, selectedCallSign),
      ),
    [selectedCallSign, selectedSummary.name],
  );
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
        fields: editableFields.filter((field) => field.section === section),
      }))
      .filter((group) => group.fields.length > 0);
  }, [editableFields]);
  const rosterFieldRows = useMemo(
    () =>
      Object.entries(selectedRow ?? {})
        .filter(([key, value]) =>
          key.startsWith(ROSTER_FIELD_PREFIX) &&
          !key.includes("fighter_status_") &&
          valueToDisplay(value as Parameters<typeof valueToDisplay>[0]).trim(),
        )
        .map(([key, value]) => {
          const sourceKey = key.slice(ROSTER_FIELD_PREFIX.length);
          const displayed = valueToDisplay(value as Parameters<typeof valueToDisplay>[0]).trim();
          return {
            key,
            label: inferRosterFieldLabel(sourceKey, displayed, rosterLabels),
            value: displayed,
          };
        }),
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
        editableFields.map((field) => [
          field.key,
          formatPersonFieldValue(selectedRow?.[field.key], field),
        ]),
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
            console.warn(
              "Періодів виходу в БД не знайдено (absences, locationPeriods, absent sheet, fighter status).",
            );
          }
        }

        if (!profile?.person) {
          console.warn(
            "Запис Person у таблиці persons відсутній — періоди зʼявляться після статусних дій з ЕЖООС.",
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
    if (!externalId || photoByExternalId[externalId]) return;

    let isCancelled = false;
    void api
      .getPersonPhoto(externalId)
      .then((photo) => {
        if (isCancelled || !photo?.photoData) return;
        setPhotoByExternalId((photos) => ({
          ...photos,
          [externalId]: photo.photoData,
        }));
      })
      .catch(() => {
        // No saved photo yet — keep silent, avatar stays placeholder.
      });

    return () => {
      isCancelled = true;
    };
  }, [photoByExternalId, selectedSummary.externalId]);

  useEffect(() => {
    const externalId = selectedSummary.externalId;
    if (!externalId) {
      setQuestionnaire(null);
      setPersonRelatedDocuments([]);
      return;
    }

    let isCancelled = false;
    setQuestionnaire(null);
    void api
      .getPersonQuestionnaire(externalId)
      .then((next) => {
        if (!isCancelled) setQuestionnaire(next);
      })
      .catch(() => {
        if (!isCancelled) setQuestionnaire(null);
      });
    void api
      .listPersonDocuments(externalId)
      .then((documents) => {
        if (!isCancelled) setPersonRelatedDocuments(documents);
      })
      .catch(() => {
        if (!isCancelled) setPersonRelatedDocuments([]);
      });

    return () => {
      isCancelled = true;
    };
  }, [selectedSummary.externalId]);

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
        const response = await fetch(
          api.getPersonQuestionnaireFileUrl(
            externalId,
            questionnaireExportFileName,
            true,
          ),
        );
        if (!response.ok) {
          throw new Error("Не вдалося завантажити PDF з сервера.");
        }
        downloadBlob(await response.blob(), questionnaireExportFileName);
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

  const loadPersonnelPhotos = async () => {
    try {
      const photos = await api.listPersonPhotos();
      setPhotoByExternalId((current) => ({
        ...Object.fromEntries(
          photos
            .filter((photo) => photo.personExternalId && photo.photoData)
            .map((photo) => [photo.personExternalId, photo.photoData]),
        ),
        ...current,
      }));
      return photos;
    } catch {
      return [];
    }
  };

  const loadPersonnelQuestionnaireIds = async () => {
    try {
      const items = await api.listPersonQuestionnaires();
      setQuestionnaireByExternalId(
        Object.fromEntries(
          items
            .filter((item) => item.personExternalId)
            .map((item) => [item.personExternalId, true as const]),
        ),
      );
      return items;
    } catch {
      return [];
    }
  };

  const loadLatestPersonnelRoster = async () => {
    const latest = await api.getLatestPersonnelRoster();
    if (!latest?.sheet) {
      setRosterLabels({});
      setRosterImportName("");
      return [] as EjournalPreviewRow[];
    }

    const columns = parseDbColumns(latest.sheet.columns);
    const rows = latest.rows.map((row) => ({
      __dbRowId: row.id,
      __rowNumber: row.excelRowNumber,
      ...(row.values && typeof row.values === "object" && !Array.isArray(row.values)
        ? row.values
        : {}),
    })) as EjournalPreviewRow[];

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
    setRosterImportName(latest.sourceFileName || latest.importName);
    return rows;
  };

  const migrateAttachmentsToNewExternalIds = async (
    pairs: Array<{
      name: string;
      fromExternalId: string;
      toExternalId: string;
    }>,
    includeDocuments = true,
  ) => {
    try {
      const nextPhones = migrateStoredPersonPhones(pairs);
      setPhonesByExternalId(nextPhones);
      migrateStoredPersonSignatures(pairs);
      const migrated = await migratePersonAttachmentsBetweenIds(pairs, {
        includeDocuments,
      });
      if (migrated > 0) {
        await Promise.all([
          loadPersonnelPhotos(),
          loadPersonnelQuestionnaireIds(),
        ]);
      }
      return migrated;
    } catch {
      return 0;
    }
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

      sessionStorage.setItem(ATTACHMENT_HEAL_SESSION_KEY, "1");
      const pairs = buildOrphanAttachmentMigrationPairs(rows, orphanIds);
      if (!isCancelled?.()) {
        setPhonesByExternalId(migrateStoredPersonPhones(pairs));
        migrateStoredPersonSignatures(pairs);
      }
      if (!orphanIds.size) return;

      const migrated = await migratePersonAttachmentsBetweenIds(pairs, {
        includeDocuments: false,
        photos,
        questionnaires,
      });
      if (migrated > 0 && !isCancelled?.()) {
        await Promise.all([
          loadPersonnelPhotos(),
          loadPersonnelQuestionnaireIds(),
        ]);
      }
    } catch {
      // Background heal must never block the personnel list.
    }
  };

  const filteredWithQuestionnaireCount = useMemo(
    () =>
      filteredPersonnel.reduce((count, record) => {
        const externalId = record.summary.externalId;
        return externalId && questionnaireByExternalId[externalId]
          ? count + 1
          : count;
      }, 0),
    [filteredPersonnel, questionnaireByExternalId],
  );

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
      const missingPhoto = !photoByExternalId[externalId];
      if (!missingQuestionnaire && !missingPhoto) continue;
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
  }, [personnelRows, photoByExternalId, questionnaireByExternalId]);

  const loadPersonnel = async (isCancelled?: () => boolean) => {
    setIsLoading(true);
    try {
      const nextImports = await api.listEjournalImports();
      if (isCancelled?.()) return;

      const sheet = findEjournalPersonnelSheet(nextImports);
      setImports(nextImports);
      if (!sheet) {
        setDbPreview(null);
        setMessage("У БД ще немає ЕЖООС-імпорту для особового складу.");
        return;
      }

      const [preview, latestRosterRows] = await Promise.all([
        loadAllEjournalSheetRows(sheet),
        loadLatestPersonnelRoster().catch(() => [] as EjournalPreviewRow[]),
      ]);
      if (isCancelled?.()) return;

      let mergedPreview = preview;
      try {
        mergedPreview = mergeRosterRowsIntoPreview(preview, latestRosterRows);
      } catch {
        mergedPreview = preview;
      }
      const rows = mergedPreview.rows.filter(isLikelyPersonnelRow);

      // Show the roster immediately — photos and rematch must not block first paint.
      const focusTarget = readPersonnelFocusTarget();
      const focusedRow =
        (focusTarget.rowId &&
          rows.find((row) => row.__dbRowId === focusTarget.rowId)) ||
        (focusTarget.externalId &&
          rows.find(
            (row) => resolvePersonIdentityKey(row) === focusTarget.externalId,
          )) ||
        null;

      setDbPreview(mergedPreview);
      setSelectedRowId(focusedRow?.__dbRowId ?? rows[0]?.__dbRowId ?? "");

      if (focusedRow) {
        clearPersonnelFocusTarget();
      }

      setMessage(
        focusedRow
          ? `Відкрито картку: ${getPersonDisplayName(focusedRow) || "особу"}.`
          : `Завантажено особовий склад з БД: ${rows.length} записів · ${sheet.name}.`,
      );
      if (!isCancelled?.()) setIsLoading(false);

      const startAttachments = () => {
        if (isCancelled?.()) return;
        const photosPromise = loadPersonnelPhotos();
        const questionnairesPromise = loadPersonnelQuestionnaireIds();
        void healOrphanAttachmentsInBackground(
          mergedPreview.rows,
          isCancelled,
          photosPromise,
          questionnairesPromise,
        );
      };
      window.requestAnimationFrame(() => {
        window.setTimeout(startAttachments, 0);
      });
    } catch (error) {
      if (isCancelled?.()) return;
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося завантажити особовий склад.",
      );
    } finally {
      if (!isCancelled?.()) setIsLoading(false);
    }
  };

  const importPersonnelRoster = async (file: File | undefined) => {
    if (!file) return;

    setIsLoading(true);
    try {
      const previousRecords = personnelRows;
      const snapshot = await readWorkbookSnapshot(file);
      const rosterSheet = snapshot.sheets.find((sheet) =>
        /загальний\s*список/i.test(sheet.sheetName),
      );
      if (!rosterSheet) {
        setMessage("У файлі не знайдено аркуш «Загальний список».");
        return;
      }

      const fighterStatusSheet = findFighterStatusSheet(snapshot.sheets);
      const fighterStatusAdditions = fighterStatusSheet
        ? buildFighterStatusAdditions(fighterStatusSheet)
        : new Map<string, Record<string, unknown>>();
      const columns = [
        ...buildImportColumns(rosterSheet),
        ...FIGHTER_STATUS_FIELDS.map((field, index) => ({
          key: field.key,
          label: field.label,
          order: rosterSheet.columnCount + index,
          originalIndex: rosterSheet.columnCount + index,
          letter: "",
        })),
      ];
      const rosterColumns = buildImportColumns(rosterSheet);
      let matchedFighterStatusCount = 0;
      const rows = rosterSheet.rows
        .filter((row) => hasRowData(row.values))
        .map((row) => {
          const values = Object.fromEntries(
            rosterColumns.map((column, index) => [
              column.key,
              cellValueToJson(row.values[index]),
            ]),
          );
          const statusAddition = findFighterStatusAddition(
            values,
            fighterStatusAdditions,
          );
          if (statusAddition) matchedFighterStatusCount += 1;

          return {
            excelRowNumber: row.excelRowNumber,
            values: {
              ...values,
              ...(statusAddition ?? {}),
            },
          };
        });

      const created = await api.importPersonnelRoster({
        name: snapshot.fileName.replace(/\.(xlsx|xlsm)$/i, ""),
        sourceFileName: snapshot.fileName,
        notes: "Імпорт Загальний список для доповнення карток особового складу.",
        sheets: [
          {
            name: rosterSheet.sheetName,
            sheetIndex: rosterSheet.sheetIndex,
            columns,
            rows,
          },
        ],
      });

      const latestRosterRows = await loadLatestPersonnelRoster();
      let nextPreview = dbPreview;
      try {
        nextPreview = dbPreview
          ? mergeRosterRowsIntoPreview(dbPreview, latestRosterRows)
          : null;
      } catch {
        nextPreview = dbPreview;
      }
      setDbPreview(nextPreview);
      setIsLoading(false);
      await new Promise<void>((resolve) => {
        window.setTimeout(resolve, 0);
      });

      const migrationPairs = nextPreview
        ? [
            ...buildAttachmentMigrationPairs(previousRecords, nextPreview.rows),
            ...buildSelfAttachmentMigrationPairs(nextPreview.rows),
          ]
        : [];
      const migrationResult = await migrateAttachmentsToNewExternalIds(
        migrationPairs,
        true,
      );
      sessionStorage.removeItem(ATTACHMENT_HEAL_SESSION_KEY);

      setMessage(
        `Імпортовано Загальний список: ${created.totalRows} рядків · ${snapshot.fileName}. Статус бійців: ${fighterStatusSheet ? `аркуш «${fighterStatusSheet.sheetName.trim()}», ${fighterStatusAdditions.size} записів, привʼязано ${matchedFighterStatusCount}` : "аркуш не знайдено"}. Перенесено привʼязок: ${migrationResult}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося імпортувати Загальний список.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    let cancelled = false;
    void loadPersonnel(() => cancelled);
    return () => {
      cancelled = true;
    };
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

    // Show immediately even if API is slow/unavailable.
    setPhotoByExternalId((photos) => ({
      ...photos,
      [externalId]: dataUrl,
    }));

    try {
      const savedPhoto = await api.upsertPersonPhoto(externalId, {
        photoData: dataUrl,
        fileName: photoCropFile?.name,
        mimeType: "image/jpeg",
        crop,
      });
      setPhotoByExternalId((photos) => ({
        ...photos,
        [externalId]: savedPhoto?.photoData || dataUrl,
      }));
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

  const pullContactsFromQuestionnaire = async () => {
    const externalId = selectedSummary.externalId;
    const rowId = selectedRow?.__dbRowId ? String(selectedRow.__dbRowId) : "";
    if (!selectedRow || !externalId || !rowId) {
      setMessage("Спочатку виберіть особу зі списку.");
      return;
    }

    setIsPullingFromQuestionnaire(true);
    try {
      let file: File | null = pendingQuestionnaireFile || diskPreviewFile || null;
      if (!file) {
        const full = await api.getPersonQuestionnaire(externalId);
        if (!full?.fileData) {
          setMessage("Немає збереженої анкети — спочатку додайте PDF анкети.");
          return;
        }
        file = dataUrlToFile(
          full.fileData,
          full.fileName || "anketa.pdf",
        );
      }

      setMessage("Читаю анкету…");
      const parsed = await parseQuestionnairePdf(file, { useOcr: false });
      const byKey = Object.fromEntries(
        parsed.fields.map((field) => [field.key, field.value.trim()]),
      );
      const enrichment = await syncEnrichmentToPerson({
        personExternalId: externalId,
        rowId,
        row: selectedRow,
        patch: {
          rnokpp: byKey.rnokpp,
          address: byKey.actualAddress || byKey.registrationAddress,
          phones: extractPhones(byKey.phones || ""),
        },
        existingPhones: phonesByExternalId[externalId] ?? [],
        phoneDocument: phoneDocByExternalId[externalId] ?? null,
      });

      setPhonesByExternalId((current) => {
        const next = { ...current, [externalId]: enrichment.phones };
        writeStoredPersonPhones(next);
        return next;
      });
      if (enrichment.phoneDocument) {
        setPhoneDocByExternalId((current) => ({
          ...current,
          [externalId]: enrichment.phoneDocument!,
        }));
      }

      if (Object.keys(enrichment.fieldUpdates).length) {
        setDbPreview((currentPreview) => {
          if (!currentPreview) return currentPreview;
          return {
            ...currentPreview,
            rows: currentPreview.rows.map((row) =>
              row.__dbRowId === selectedRow.__dbRowId
                ? applyEnrichmentToPreviewRow(row, enrichment.fieldUpdates)
                : row,
            ),
          };
        });
        setEditValues((current) => ({
          ...current,
          ...enrichment.fieldUpdates,
        }));
      }

      const fieldKeys = Object.keys(enrichment.fieldUpdates);
      const parts = [
        enrichment.phonesAdded.length
          ? `телефони +${enrichment.phonesAdded.length}`
          : "",
        fieldKeys.some((key) => key.toLocaleLowerCase("uk-UA").includes("рнокпп"))
          ? "РНОКПП"
          : "",
        fieldKeys.some((key) => key.toLocaleLowerCase("uk-UA").includes("адрес"))
          ? "адреса"
          : "",
      ].filter(Boolean);

      setMessage(
        parts.length
          ? `З анкети додано (без перезапису наявного): ${parts.join(", ")}.`
          : "У анкеті немає нових даних для порожніх полів — наявне не змінено.",
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? `Не вдалося підтягнути з анкети: ${error.message}`
          : "Не вдалося підтягнути дані з анкети.",
      );
    } finally {
      setIsPullingFromQuestionnaire(false);
    }
  };

  const deleteSelectedQuestionnaire = async () => {
    const externalId = selectedSummary.externalId;
    if (!externalId || !questionnaire) return;
    if (!window.confirm(`Видалити анкету для ${selectedSummary.name || "особи"}?`)) {
      return;
    }

    try {
      await api.deletePersonQuestionnaire(externalId);
      setQuestionnaire(null);
      setQuestionnaireByExternalId((current) => {
        const next = { ...current };
        delete next[externalId];
        return next;
      });
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
    if (record?.row.__dbRowId) setSelectedRowId(record.row.__dbRowId);
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

  const openQuestionnairePreview = (fileData = questionnaire?.fileData) => {
    const externalId = selectedSummary.externalId;
    let nextUrl = "";

    if (
      externalId &&
      fileData &&
      !pendingQuestionnaireFile &&
      !diskPreviewFile
    ) {
      nextUrl = api.getPersonQuestionnaireFileUrl(
        externalId,
        questionnaireExportFileName,
      );
    } else if (fileData) {
      nextUrl = dataUrlToObjectUrl(fileData);
    } else {
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

  const openQuestionnaireInNewTab = () => {
    const externalId = selectedSummary.externalId;
    if (
      externalId &&
      questionnaire?.fileData &&
      !pendingQuestionnaireFile &&
      !diskPreviewFile
    ) {
      window.open(
        api.getPersonQuestionnaireFileUrl(externalId, questionnaireExportFileName),
        "_blank",
        "noopener,noreferrer",
      );
      return;
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
          selectedSummary.callSign,
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

  const openPhotoCropFromQuestionnaire = () => {
    if (diskPreviewFile) {
      openPhotoCrop(diskPreviewFile, { floating: true });
      return;
    }
    if (pendingQuestionnaireFile) {
      openPhotoCrop(pendingQuestionnaireFile);
      return;
    }
    if (!questionnaire?.fileData) {
      setMessage("Немає PDF анкети, з якого можна вирізати фото.");
      return;
    }
    openPhotoCrop(
      dataUrlToFile(
        questionnaire.fileData,
        questionnaire.fileName || "questionnaire.pdf",
      ),
    );
  };

  return (
    <main className="main-panel personnel-page">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Особовий склад
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Список із ЕЖООС · картка особи · редагування staging-даних
          </Typography>
        </Box>
        <Stack direction="row" spacing={1}>
          <Button
            variant="outlined"
            disabled={!missingDiskSearchPeople.length}
            onClick={() => setIsDiskSearchOpen(true)}
          >
            Пошук усіх анкет
          </Button>
          <Button
            component="label"
            disabled={isLoading}
            startIcon={<FileUploadOutlinedIcon />}
            variant="outlined"
          >
            Імпорт Загальний список
            <input
              hidden
              type="file"
              accept=".xlsx,.xlsm"
              onChange={(event) => {
                void importPersonnelRoster(event.target.files?.[0]);
                event.target.value = "";
              }}
            />
          </Button>
          <Button variant="outlined" onClick={() => void loadPersonnel()}>
            Оновити з БД
          </Button>
        </Stack>
      </header>
      {isLoading && <LinearProgress color="primary" />}
      <Alert severity="info" variant="outlined" className="personnel-page-alert">
        {message}
      </Alert>

      <section className="personnel-layout">
        <aside className="analytics-panel personnel-list-panel">
          <div className="panel-heading">
            Військовослужбовці · {filteredPersonnel.length}
            <span className="personnel-list-questionnaire-count">
              {" "}
              · з анкетами {filteredWithQuestionnaireCount}
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
          <PersonnelVirtualList
            items={filteredPersonnel}
            selectedRowId={selectedRowId}
            photoByExternalId={photoByExternalId}
            onSelect={setSelectedRowId}
            keyboardEnabled={
              !isPhotoCropOpen &&
              !isQuestionnairePreviewOpen &&
              !isDiskSearchOpen &&
              !activePersonAction
            }
          />
        </aside>

        <section className="person-card-panel">
          <div className="person-card-hero">
            <div className="person-avatar">
              {selectedPhoto ? (
                <img alt={selectedSummary.name} src={selectedPhoto} />
              ) : (
                <PersonSearchOutlinedIcon />
              )}
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
            <div>
              <div className="panel-heading">Картка особи</div>
              {selectedCallSign ? (
                <div className="person-callsign-row">
                  <span className="person-callsign" title="Позивний">
                    <span className="person-callsign-label">позивний</span>
                    <strong>{selectedCallSign}</strong>
                  </span>
                </div>
              ) : null}
              <Typography component="h2" variant="h4">
                {selectedSummary.name}
              </Typography>
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
                <Button
                  disabled={
                    !selectedRow ||
                    !selectedSummary.externalId ||
                    isPullingFromQuestionnaire
                  }
                  size="small"
                  type="button"
                  variant="outlined"
                  sx={{ mt: 1, alignSelf: "flex-start" }}
                  onClick={() => void pullContactsFromQuestionnaire()}
                >
                  {isPullingFromQuestionnaire
                    ? "Читаю анкету…"
                    : "Підтягнути з анкети"}
                </Button>
              </span>
              <span>
                <strong>Дата народження</strong>
                {selectedSummary.birthDate || "—"}
              </span>
              <span>
                <strong>РНОКПП</strong>
                {selectedSummary.rnokpp || "—"}
              </span>
              <span>
                <strong>Дислокація</strong>
                {selectedSummary.location || "—"}
              </span>
              <span>
                <strong>Контракт</strong>
                {[selectedSummary.contractFrom, selectedSummary.contractTo]
                  .filter(Boolean)
                  .join(" — ") || "—"}
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
                  {fighterStatusFieldRows.map((field) => (
                    <span key={field.key}>
                      <strong>{field.label.replace(/^Статус бійців · /, "")}</strong>
                      {field.value}
                    </span>
                  ))}
                </div>
              </div>
            ) : null}

            {rosterFieldRows.length > 0 ? (
              <div className="person-edit-section">
                <div className="panel-heading">
                  Загальний список
                  {rosterImportName ? ` · ${rosterImportName}` : ""}
                </div>
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

                    return (
                      <label className={isWide ? "wide" : ""} key={field.key}>
                        <span>{field.label}</span>
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
                          />
                        )}
                      </label>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>

          <Stack
            className="person-card-actions"
            direction="row"
            spacing={1}
            sx={{ justifyContent: "flex-end" }}
          >
            <Button
              component="label"
              disabled={!selectedRow || !selectedSummary.externalId || isUploadingQuestionnaire}
              variant="outlined"
              startIcon={<PictureAsPdfOutlinedIcon />}
            >
              {questionnaire ? "Замінити анкету" : "Додати анкету"}
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
            <Button
              disabled={!selectedRow}
              variant="outlined"
              onClick={() => selectedRow && onOpenDocuments(selectedRow)}
            >
              Створити документ
            </Button>
            <Button
              disabled={!selectedRow}
              variant="contained"
              onClick={() => void saveSelectedPerson()}
              sx={{ color: "#1a1a14" }}
            >
              Зберегти зміни
            </Button>
          </Stack>
        </section>

        <aside className="person-side-panel">
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
                    onClick={() => openQuestionnairePreview()}
                  >
                    <PictureAsPdfOutlinedIcon />
                    <span>
                      <strong>Анкета (PDF)</strong>
                      <small>
                        {questionnaireExportFileName} · переглянути
                      </small>
                    </span>
                  </button>
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
                    aria-label="Видалити анкету"
                    className="person-document-delete"
                    disabled={!selectedRow}
                    onClick={() => void deleteSelectedQuestionnaire()}
                    title="Видалити анкету"
                    type="button"
                  >
                    <DeleteOutlineOutlinedIcon />
                  </button>
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
          setMessage(`Фото автоматично знайдено в PDF і додано до preview для ID ${externalId}.`);
        }}
        onPreviewQuestionnaire={(file, title, externalId) => {
          focusPersonByExternalId(externalId);
          openDiskQuestionnairePreview(file, title);
        }}
        onCropPhoto={openDiskPhotoCrop}
      />

      <FloatingQuestionnairePreview
        open={isQuestionnairePreviewOpen && isDiskFloatingPreview}
        title={
          questionnairePreviewTitle ||
          `Анкета · ${selectedSummary.name}`
        }
        previewUrl={questionnairePreviewUrl}
        pendingFile={false}
        isUploading={false}
        placement="left"
        onClose={closeQuestionnairePreview}
        onCrop={() => openPhotoCropFromQuestionnaire()}
        onOpenTab={openQuestionnaireInNewTab}
        onDownload={() => void exportCurrentQuestionnaire()}
        shareFileName={questionnaireExportFileName}
        sharePersonName={selectedSummary.name}
        shareSource={currentQuestionnaireShareSource}
        onShareNotify={setMessage}
      />

      <Dialog
        fullWidth
        maxWidth="md"
        open={isQuestionnairePreviewOpen && !isDiskFloatingPreview}
        onClose={closeQuestionnairePreview}
        slotProps={{ paper: { className: "questionnaire-preview-dialog" } }}
      >
        <DialogTitle>
          {pendingQuestionnaireFile
            ? `Перегляд анкети перед збереженням · ${selectedSummary.name}`
            : `Анкета · ${selectedSummary.name}`}
          {pendingQuestionnaireFile
            ? ` · ${pendingQuestionnaireFile.name}`
            : questionnaireExportFileName
              ? ` · ${questionnaireExportFileName}`
              : ""}
        </DialogTitle>
        <DialogContent>
          {pendingQuestionnaireFile ? (
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              Перевірте, що це потрібна анкета. Можна одразу вирізати фото з PDF,
              потім зберегти анкету в БД.
            </Typography>
          ) : questionnairePreviewUrl ? (
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              Для збереження на диск натисніть «Експорт PDF». Збереження через Cmd+S
              у переглядачі може дати випадкову назву файлу.
            </Typography>
          ) : null}
          {questionnairePreviewUrl ? (
            <iframe
              className="questionnaire-preview-frame"
              src={questionnairePreviewUrl}
              title="Перегляд анкети PDF"
            />
          ) : (
            <Typography variant="body2" color="text.secondary">
              Немає PDF для перегляду.
            </Typography>
          )}
        </DialogContent>
        <DialogActions>
          <Button
            variant="outlined"
            disabled={!questionnairePreviewUrl}
            startIcon={<AddPhotoAlternateOutlinedIcon />}
            onClick={() => openPhotoCropFromQuestionnaire()}
          >
            Вирізати фото
          </Button>
          <Button
            variant="contained"
            disabled={!questionnairePreviewUrl}
            onClick={() => void exportCurrentQuestionnaire()}
            sx={{ color: "#1a1a14" }}
          >
            Експорт PDF
          </Button>
          <QuestionnaireShareButton
            disabled={!questionnairePreviewUrl}
            fileName={questionnaireExportFileName}
            personName={selectedSummary.name}
            source={currentQuestionnaireShareSource}
            onNotify={setMessage}
          />
          {!pendingQuestionnaireFile && !diskPreviewFile ? (
            <Button
              variant="outlined"
              disabled={!questionnairePreviewUrl}
              onClick={openQuestionnaireInNewTab}
            >
              Відкрити в новій вкладці
            </Button>
          ) : null}
          {pendingQuestionnaireFile ? (
            <>
              <Button variant="outlined" onClick={closeQuestionnairePreview}>
                Скасувати
              </Button>
              <Button
                variant="contained"
                disabled={isUploadingQuestionnaire}
                onClick={() => void confirmPendingQuestionnaire()}
                sx={{ color: "#1a1a14" }}
              >
                {isUploadingQuestionnaire ? "Збереження…" : "Зберегти анкету"}
              </Button>
            </>
          ) : (
            <Button
              variant="contained"
              onClick={closeQuestionnairePreview}
              sx={{ color: "#1a1a14" }}
            >
              Закрити
            </Button>
          )}
        </DialogActions>
      </Dialog>

      <section className="panel table-panel personnel-source-panel">
        <div className="panel-heading">
          Джерело: {dbPreview?.sheet.name ?? "2. ООС"} · імпортів ЕЖООС:{" "}
          {imports.length}
        </div>
        <div className="personnel-raw-preview">
          {selectedRow ? (
            <pre>{JSON.stringify(selectedRow, null, 2)}</pre>
          ) : (
            <Typography variant="body2" color="text.secondary">
              Виберіть людину зі списку.
            </Typography>
          )}
        </div>
      </section>
    </main>
  );
}
