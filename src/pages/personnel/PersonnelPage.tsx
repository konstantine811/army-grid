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
import { PersonSearchOutlinedIcon } from "@/components/sci/icons";
import { PictureAsPdfOutlinedIcon } from "@/components/sci/icons";
import { SearchOutlinedIcon } from "@/components/sci/icons";
import {
  api,
  type BackendEjournalImport,
  type BackendPersonQuestionnaire,
} from "../../api";
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
import { PersonnelVirtualList } from "./PersonnelVirtualList";
import { QuestionnaireDiskSearchDialog } from "./QuestionnaireDiskSearchDialog";
import {
  PERSON_CARD_FIELDS,
  PERSON_SECTION_LABELS,
  buildPersonSummary,
  createDefaultActionForm,
  dataUrlToFile,
  dataUrlToObjectUrl,
  extractPersonCallSign,
  extractPhones,
  findEjournalPersonnelSheet,
  formatPersonFieldValue,
  formatUaPhoneDisplay,
  getPersonExternalId,
  isLikelyPersonnelRow,
  isPositionIndexField,
  loadAllEjournalSheetRows,
  personActions,
  resolvePersonFieldKey,
  type PersonAction,
  type PersonActionForm,
  type PersonFieldDef,
  type PersonnelRecord,
} from "./personnelUtils";

const PERSONNEL_FOCUS_KEY = "army-grid:focus-personnel";
const MAX_QUESTIONNAIRE_FILE_BYTES = 350 * 1024 * 1024;

const formatFileSize = (bytes: number) => {
  if (!Number.isFinite(bytes)) return "";
  const mb = bytes / 1024 / 1024;
  return `${mb.toFixed(mb >= 10 ? 0 : 1)} MB`;
};

const ROSTER_FIELD_PREFIX = "roster__";

const normalizeRosterText = (value: unknown) =>
  valueToDisplay(value as Parameters<typeof valueToDisplay>[0])
    .replace(/[ʼ’']/g, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[.,;:№#"/\\|()[\]{}]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const getRosterValue = (row: EjournalPreviewRow, keyParts: string[]) => {
  const key = Object.keys(row).find((item) =>
    keyParts.every((part) => item.toLocaleLowerCase("uk-UA").includes(part)),
  );
  return key ? valueToDisplay(row[key] as Parameters<typeof valueToDisplay>[0]).trim() : "";
};

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
  const externalId = getRosterValue(rosterRow, ["id"]);
  const name =
    getRosterValue(rosterRow, ["піб"]) ||
    getRosterValue(rosterRow, ["прізвище"]);
  const rowKey = externalId || normalizeRosterText(name);

  return {
    __dbRowId: `roster:${rowKey}`,
    __rowNumber: rosterRow.__rowNumber,
    id: externalId,
    "прізвище": name,
    "ПІБ": name,
    "звання": getRosterValue(rosterRow, ["звання"]),
    "Звання": getRosterValue(rosterRow, ["звання"]),
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
    const id = getRosterValue(row, ["id"]);
    const name = getRosterValue(row, ["піб"]) || getRosterValue(row, ["прізвище"]);
    if (id && id !== "0") rosterById.set(id, row);
    if (name) rosterByName.set(normalizeRosterText(name), row);
  });

  const mergedRows = preview.rows.map((row) => {
    const externalId = getPersonExternalId(row);
    const name = buildPersonSummary(row).name;
    const rosterRow =
      (externalId && rosterById.get(externalId)) ||
      rosterByName.get(normalizeRosterText(name));

    if (!rosterRow) return row;
    usedRosterRows.add(rosterRow);
    return { ...row, ...getRosterAdditions(rosterRow) };
  });
  const rosterOnlyRows = rosterRows
    .filter((row) => !usedRosterRows.has(row))
    .filter((row) => getRosterValue(row, ["піб"]) || getRosterValue(row, ["прізвище"]))
    .map(buildRosterOnlyPersonnelRow);

  return {
    ...preview,
    rows: [...mergedRows, ...rosterOnlyRows],
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
    mode?: "default" | "salaryPowerAttorney" | "ubdReport",
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
  const [questionnaireByExternalId, setQuestionnaireByExternalId] = useState<
    Record<string, true>
  >({});
  const [photoCropFile, setPhotoCropFile] = useState<File | null>(null);
  const [isPhotoCropOpen, setIsPhotoCropOpen] = useState(false);
  const [questionnaire, setQuestionnaire] =
    useState<BackendPersonQuestionnaire | null>(null);
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
        .map((row) => ({ row, summary: buildPersonSummary(row) })),
    [dbPreview],
  );
  const filteredPersonnel = useMemo(() => {
    const normalizedQuery = query.trim().toLowerCase();

    if (!normalizedQuery) return personnelRows;

    return personnelRows.filter((record) =>
      [
        record.summary.name,
        record.summary.rank,
        record.summary.externalId,
        record.summary.positionIndex,
        record.summary.location,
        record.summary.rnokpp,
        record.summary.phones.join(" "),
        record.summary.additionalInfo,
        record.summary.militaryId,
      ]
        .join(" ")
        .toLowerCase()
        .includes(normalizedQuery),
    );
  }, [personnelRows, query]);
  const selectedRecord = useMemo(
    () =>
      personnelRows.find((record) => record.row.__dbRowId === selectedRowId) ??
      filteredPersonnel[0] ??
      null,
    [filteredPersonnel, personnelRows, selectedRowId],
  );
  const selectedSummary = selectedRecord?.summary ?? buildPersonSummary(null);
  const selectedRow = selectedRecord?.row ?? null;
  const selectedPhoto =
    (selectedSummary.externalId && photoByExternalId[selectedSummary.externalId]) ||
    "";
  const selectedCallSign = useMemo(
    () =>
      extractPersonCallSign(
        selectedSummary.callSign,
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
          valueToDisplay(value as Parameters<typeof valueToDisplay>[0]).trim(),
        )
        .map(([key, value]) => {
          const sourceKey = key.slice(ROSTER_FIELD_PREFIX.length);
          return {
            key,
            label: rosterLabels[sourceKey] || sourceKey,
            value: valueToDisplay(value as Parameters<typeof valueToDisplay>[0]).trim(),
          };
        }),
    [rosterLabels, selectedRow],
  );
  const additionalInfoKey = useMemo(
    () => resolvePersonFieldKey(selectedRow, ["додаткова_інформація"]),
    [selectedRow],
  );
  /** Live-parsed separate property from «Додаткова інформація» (including draft edits). */
  const parsedPhones = useMemo(() => {
    const draft =
      (additionalInfoKey && editValues[additionalInfoKey]) ||
      selectedSummary.additionalInfo;
    return extractPhones(draft);
  }, [additionalInfoKey, editValues, selectedSummary.additionalInfo]);

  useEffect(() => {
    setEditValues(
      Object.fromEntries(
        editableFields.map((field) => [
          field.key,
          formatPersonFieldValue(selectedRow?.[field.key], field),
        ]),
      ),
    );
  }, [editableFields, selectedRow]);

  useEffect(() => {
    if (selectedRecord?.row.__dbRowId && !selectedRowId) {
      setSelectedRowId(selectedRecord.row.__dbRowId);
    }
  }, [selectedRecord, selectedRowId]);

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

    return () => {
      isCancelled = true;
    };
  }, [selectedSummary.externalId]);

  useEffect(() => {
    return () => {
      if (questionnairePreviewUrl) URL.revokeObjectURL(questionnairePreviewUrl);
    };
  }, [questionnairePreviewUrl]);

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
    } catch {
      // List previews stay as placeholders if the bulk endpoint is unavailable.
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
    } catch {
      // Missing list endpoint should not block the page.
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
      Object.fromEntries(columns.map((column) => [column.key, column.label || column.key])),
    );
    setRosterImportName(latest.sourceFileName || latest.importName);
    return rows;
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

      const mergedPreview = mergeRosterRowsIntoPreview(preview, latestRosterRows);
      const rows = mergedPreview.rows.filter(isLikelyPersonnelRow);
      // Read focus only after data is ready so a cancelled StrictMode pass
      // cannot clear the target before the active mount applies it.
      const focusTarget = readPersonnelFocusTarget();
      const focusedRow =
        (focusTarget.rowId &&
          rows.find((row) => row.__dbRowId === focusTarget.rowId)) ||
        (focusTarget.externalId &&
          rows.find(
            (row) => getPersonExternalId(row) === focusTarget.externalId,
          )) ||
        null;

      setDbPreview(mergedPreview);
      setSelectedRowId(focusedRow?.__dbRowId ?? rows[0]?.__dbRowId ?? "");

      if (focusedRow) {
        clearPersonnelFocusTarget();
      }

      void loadPersonnelPhotos();
      void loadPersonnelQuestionnaireIds();
      setMessage(
        focusedRow
          ? `Відкрито картку: ${buildPersonSummary(focusedRow).name}.`
          : `Завантажено особовий склад з БД: ${rows.length} записів · ${sheet.name}.`,
      );
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
      const snapshot = await readWorkbookSnapshot(file);
      const rosterSheet = snapshot.sheets.find((sheet) =>
        /загальний\s*список/i.test(sheet.sheetName),
      );
      if (!rosterSheet) {
        setMessage("У файлі не знайдено аркуш «Загальний список».");
        return;
      }

      const columns = buildImportColumns(rosterSheet);
      const rows = rosterSheet.rows
        .filter((row) => hasRowData(row.values))
        .map((row) => ({
          excelRowNumber: row.excelRowNumber,
          values: Object.fromEntries(
            columns.map((column, index) => [
              column.key,
              cellValueToJson(row.values[index]),
            ]),
          ),
        }));

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
      setDbPreview((current) =>
        current ? mergeRosterRowsIntoPreview(current, latestRosterRows) : current,
      );
      setMessage(
        `Імпортовано Загальний список: ${created.totalRows} рядків · ${snapshot.fileName}.`,
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
      setMessage(`Фото збережено в БД: ${selectedSummary.name} · ID ${externalId}.`);
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
      setMessage(`Фото видалено: ${selectedSummary.name} · ID ${externalId}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося видалити фото з БД.",
      );
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
      setMessage(`Анкету видалено: ${selectedSummary.name} · ID ${externalId}.`);
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
      if (current) URL.revokeObjectURL(current);
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
      if (current) URL.revokeObjectURL(current);
      return "";
    });
  };

  const openQuestionnairePreview = (fileData = questionnaire?.fileData) => {
    if (!fileData) return;
    const nextUrl = dataUrlToObjectUrl(fileData);
    setPendingQuestionnaireFile(null);
    setDiskPreviewFile(null);
    setIsDiskFloatingPreview(false);
    setQuestionnairePreviewTitle(
      `Анкета · ${selectedSummary.name}${
        questionnaire?.fileName ? ` · ${questionnaire.fileName}` : ""
      }`,
    );
    setQuestionnairePreviewUrl((current) => {
      if (current) URL.revokeObjectURL(current);
      return nextUrl;
    });
    setIsQuestionnairePreviewOpen(true);
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
      if (current) URL.revokeObjectURL(current);
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
      const saved = await api.upsertPersonQuestionnaireFile(externalId, file);
      setQuestionnaire(saved);
      setQuestionnaireByExternalId((current) => ({
        ...current,
        [externalId]: true,
      }));
      setMessage(
        `Анкету збережено в БД: ${selectedSummary.name} · ${file.name}.`,
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
          <Button
            disabled={!selectedRow}
            variant="contained"
            onClick={() => selectedRow && onOpenDocuments(selectedRow)}
            sx={{ color: "#1a1a14" }}
          >
            Сформувати документ
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
              placeholder="ПІБ, ID, звання, посада"
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
                {selectedSummary.externalId && (
                  <Chip label={`ID: ${selectedSummary.externalId}`} size="small" />
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
                    {parsedPhones.map((phone, phoneIndex) => (
                      <Chip
                        key={`${phone}-${phoneIndex}`}
                        label={formatUaPhoneDisplay(phone)}
                        size="small"
                        color="primary"
                        variant="outlined"
                      />
                    ))}
                  </span>
                ) : (
                  "—"
                )}
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
                <article className="person-document-shell">
                  <button
                    className="person-document-item"
                    type="button"
                    onClick={() => openQuestionnairePreview()}
                  >
                    <PictureAsPdfOutlinedIcon />
                    <span>
                      <strong>Анкета (PDF)</strong>
                      <small>
                        {questionnaire.fileName || "questionnaire.pdf"} · натисніть,
                        щоб переглянути
                      </small>
                    </span>
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
                className="person-document-item"
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
                className="person-document-item"
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
        onOpenTab={() =>
          questionnairePreviewUrl &&
          window.open(questionnairePreviewUrl, "_blank", "noopener,noreferrer")
        }
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
            : questionnaire?.fileName
              ? ` · ${questionnaire.fileName}`
              : ""}
        </DialogTitle>
        <DialogContent>
          {pendingQuestionnaireFile ? (
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              Перевірте, що це потрібна анкета. Можна одразу вирізати фото з PDF,
              потім зберегти анкету в БД.
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
            variant="outlined"
            disabled={!questionnairePreviewUrl}
            onClick={() =>
              questionnairePreviewUrl &&
              window.open(questionnairePreviewUrl, "_blank", "noopener,noreferrer")
            }
          >
            Відкрити в новій вкладці
          </Button>
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
