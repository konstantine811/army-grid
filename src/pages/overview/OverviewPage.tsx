import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import writeXlsxFile, { type SheetData } from "write-excel-file/browser";
import {
  Alert,
  Box,
  Button,
  Chip,
  LinearProgress,
  MenuItem,
  Stack,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import { BusinessCenterOutlinedIcon } from "@/components/sci/icons";
import { CalendarMonthOutlinedIcon } from "@/components/sci/icons";
import { LocalHospitalOutlinedIcon } from "@/components/sci/icons";
import { BeachAccessOutlinedIcon } from "@/components/sci/icons";
import { PersonOutlinedIcon } from "@/components/sci/icons";
import { SearchOutlinedIcon } from "@/components/sci/icons";
import { ShieldOutlinedIcon } from "@/components/sci/icons";
import {
  api,
  type BackendPersonnelOverview,
  type BackendPersonnelOverviewRow,
  type BackendPersonnelRosterLatest,
} from "../../api";
import {
  openPersonnelInNewTab,
} from "../../app/navigation";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
  readDataCache,
} from "../../data/idbDataCache";
import { loadSharedRosterLatest } from "../../data/sharedAppData";
import {
  createStoredZipBlob,
  dataUrlToUint8Array,
  downloadBlob,
  sanitizeFileName,
} from "../../shared/browserExport";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  buildQuestionnaireExportFileName,
} from "../personnel/personnelUtils";
import { parseDbColumns } from "../ejournal/ejournalUtils";
import {
  fillMissingOverviewPhotos,
  resolveOverviewPhoto,
} from "./overviewPhotos";
import {
  applyPersonnelAssetsToOverview,
  loadPersonnelRowsForOverview,
} from "./overviewPersonnelAssets";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import {
  overviewNameMatchesQuery,
  parseOverviewNameQueries,
} from "./overviewNameSearch";
import {
  OverviewVirtualTable,
  type OverviewPersonDocumentSummary,
  type OverviewQuestionnaireTarget,
} from "./OverviewVirtualTable";
import { loadPersonnelOverviewInBatches } from "./overviewBatchLoad";
import {
  buildOverviewMetrics,
  summarizeStaffFromRoster,
} from "./overviewRosterMerge";
import { runHeavyJob } from "../../workers/runHeavyJob";
import { buildOverviewSideStats } from "./overviewSideStats";
import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";

const STATUS_FILTERS = [
  { value: "ALL", label: "Усі статуси" },
  { value: "ON_DUTY", label: "На службі" },
  { value: "BUSINESS_TRIP", label: "Відрядження" },
  { value: "LEAVE", label: "Відпустка" },
  { value: "MEDICAL", label: "Лікування" },
  { value: "AWOL", label: "СЗЧ" },
  { value: "MISSING", label: "Безвісти / втрати" },
] as const;

const PERIOD_FILTERS = [
  { value: "ALL", label: "Увесь період" },
  { value: "7", label: "Останні 7 днів" },
  { value: "30", label: "Останні 30 днів" },
  { value: "90", label: "Останні 90 днів" },
] as const;

const SOURCE_FILTERS = [
  { value: "staff", label: "Штатка" },
  { value: "ejoos", label: "ЕЖООС" },
  { value: "all", label: "Усі джерела" },
] as const;

type OverviewSourceFilter = (typeof SOURCE_FILTERS)[number]["value"];

const normalizeRosterText = normalizeRosterMatchText;

const rosterLabelsFromLatest = (latest: BackendPersonnelRosterLatest | null) => {
  const labels: Record<string, string> = {};
  for (const column of parseDbColumns(latest?.sheet?.columns)) {
    if (column.key) labels[column.key] = column.label;
  }
  return labels;
};

const withStaffOverviewStatus = (
  row: BackendPersonnelOverviewRow,
): BackendPersonnelOverviewRow => {
  if (!row.staffStatus) return row;
  return {
    ...row,
    status: row.staffStatus,
    statusLabel: row.staffStatusLabel || row.statusLabel,
  };
};

const rosterRowsFromLatest = (latest: BackendPersonnelRosterLatest | null) => {
  if (!latest?.sheet) return [] as EjournalPreviewRow[];
  return latest.rows.map((row) => ({
    __dbRowId: row.id,
    __rowNumber: row.excelRowNumber,
    ...(row.values && typeof row.values === "object" && !Array.isArray(row.values)
      ? row.values
      : {}),
  })) as EjournalPreviewRow[];
};

const rosterSourceLabel = (latest: BackendPersonnelRosterLatest | null) =>
  latest?.sourceFileName?.trim() || latest?.importName?.trim() || "Штатка";

const rosterColumnsFromLatest = (latest: BackendPersonnelRosterLatest | null) =>
  parseDbColumns(latest?.sheet?.columns);

const loadLatestPersonnelRosterRows = async (force = false) => {
  const latest = await loadSharedRosterLatest({ force });
  return {
    rows: rosterRowsFromLatest(latest),
    labels: rosterLabelsFromLatest(latest),
    columns: rosterColumnsFromLatest(latest),
    label: rosterSourceLabel(latest),
  };
};

const buildCallSignByExternalId = (rosterRows: EjournalPreviewRow[]) => {
  const map: Record<string, string> = {};
  for (const row of rosterRows) {
    const summary = buildPersonSummary(row);
    if (summary.externalId && summary.callSign) {
      map[summary.externalId] = summary.callSign;
    }
  }
  return map;
};

export function OverviewPage() {
  const [data, setData] = useState<BackendPersonnelOverview | null>(null);
  const [photos, setPhotos] = useState<Record<string, string>>({});
  const [questionnaireByExternalId, setQuestionnaireByExternalId] = useState<
    Record<string, true>
  >({});
  const [
    questionnaireSourceIdByExternalId,
    setQuestionnaireSourceIdByExternalId,
  ] = useState<Record<string, string>>({});
  const [documentsByExternalId, setDocumentsByExternalId] = useState<
    Record<string, OverviewPersonDocumentSummary>
  >({});
  const [callSignByExternalId, setCallSignByExternalId] = useState<
    Record<string, string>
  >({});
  const [query, setQuery] = useState("");
  const [source, setSource] = useState<OverviewSourceFilter>("staff");
  const [battalion, setBattalion] = useState("ALL");
  const [rosterLabel, setRosterLabel] = useState("Штатка");
  const [unit, setUnit] = useState("ALL");
  const [status, setStatus] = useState("ALL");
  const [period, setPeriod] = useState("30");
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState(`API: ${api.baseUrl}`);
  const rosterRowsRef = useRef<EjournalPreviewRow[]>([]);
  const personnelRowsRef = useRef<EjournalPreviewRow[]>([]);
  const rosterColumnsRef = useRef<
    Array<{ key: string; letter?: string; originalIndex?: number }>
  >([]);
  const photosRef = useRef<Record<string, string>>({});
  const requestedPhotoKeysRef = useRef(new Set<string>());
  const [rosterEpoch, setRosterEpoch] = useState(0);
  photosRef.current = photos;

  const load = async (
    signal?: { cancelled: boolean },
    options?: { force?: boolean },
  ) => {
    const alive = () => !signal?.cancelled;
    const force = Boolean(options?.force);
    setIsLoading(true);
    try {
      const applyOverview = async (
        overview: BackendPersonnelOverview,
        rosterRows: EjournalPreviewRow[],
        rosterLabels: Record<string, string> = {},
        rosterColumns: Array<{
          key: string;
          letter?: string;
          originalIndex?: number;
        }> = [],
        fromCache = false,
      ) => {
        let mergedOverview = overview;
        try {
          mergedOverview = await runHeavyJob({
            type: "mergeOverview",
            overview,
            rosterRows,
            rosterLabels,
            columns: rosterColumns,
          });
        } catch {
          mergedOverview = overview;
        }
        if (!alive()) return mergedOverview;
        rosterRowsRef.current = rosterRows;
        if (!personnelRowsRef.current.length) {
          personnelRowsRef.current = rosterRows;
        }
        rosterColumnsRef.current = rosterColumns;
        setRosterEpoch((value) => value + 1);
        setData(mergedOverview);
        try {
          setCallSignByExternalId(buildCallSignByExternalId(rosterRows));
        } catch {
          setCallSignByExternalId({});
        }
        setMessage(
          fromCache
            ? `Кеш огляду · оновлюю з БД…`
            : mergedOverview.importName
              ? `ЕЖООС: ${mergedOverview.importName} · Штатка: ${rosterRows.length ? "є" : "немає"} · ${mergedOverview.rows.length} записів`
              : rosterRows.length
                ? `Штатка: ${rosterRows.length} осіб · імпорту ЕЖООС немає`
                : "Немає імпорту ЕЖООС і Штатки",
        );
        return mergedOverview;
      };

      let assetsSeq = 0;
      const applyPersonnelAssets = async (
        overviewRows: BackendPersonnelOverviewRow[],
        nextRosterRows: EjournalPreviewRow[],
      ) => {
        const seq = ++assetsSeq;
        const [photoList, questionnaireList, documentList] = await Promise.all([
          api.listPersonPhotos().catch(() => []),
          fetchWithCache({
            key: CacheKeys.questionnairesMeta,
            force,
            fetcher: () => api.listPersonQuestionnaires(),
            isChanged: jsonChanged,
          }).catch(() => []),
          fetchWithCache({
            key: CacheKeys.documentsAll,
            force,
            fetcher: () => api.listAllPersonDocuments(),
            isChanged: jsonChanged,
          }).catch(() => []),
        ]);
        if (!alive() || seq !== assetsSeq) return;

        const paintAssets = (people: EjournalPreviewRow[]) => {
          try {
            const assets = applyPersonnelAssetsToOverview(
              overviewRows,
              people,
              Array.isArray(questionnaireList) ? questionnaireList : [],
              Array.isArray(photoList) ? photoList : [],
              Array.isArray(documentList) ? documentList : [],
            );
            photosRef.current = { ...photosRef.current, ...assets.photos };
            setPhotos(photosRef.current);
            setQuestionnaireByExternalId(assets.questionnairePresence);
            setQuestionnaireSourceIdByExternalId(assets.questionnaireSourceIds);
            setDocumentsByExternalId(assets.documents);
          } catch (error) {
            console.warn("[Огляд] Не вдалося підставити дані з Особового складу", error);
          }
        };

        personnelRowsRef.current = nextRosterRows;
        paintAssets(nextRosterRows);

        const personnelRows = await loadPersonnelRowsForOverview(nextRosterRows);
        if (!alive() || seq !== assetsSeq) return;
        personnelRowsRef.current = personnelRows.length
          ? personnelRows
          : nextRosterRows;
        paintAssets(personnelRowsRef.current);
      };

      const [cachedOverview, cachedRoster] = await Promise.all([
        readDataCache<BackendPersonnelOverview>(CacheKeys.overview),
        readDataCache<BackendPersonnelRosterLatest | null>(CacheKeys.rosterLatest),
      ]);
      let rosterRows: EjournalPreviewRow[] = [];
      let rosterLabels: Record<string, string> = {};
      let rosterColumns: Array<{
        key: string;
        letter?: string;
        originalIndex?: number;
      }> = [];
      if (cachedOverview) {
        rosterRows = rosterRowsFromLatest(cachedRoster);
        rosterLabels = rosterLabelsFromLatest(cachedRoster);
        rosterColumns = rosterColumnsFromLatest(cachedRoster);
        if (cachedRoster) setRosterLabel(rosterSourceLabel(cachedRoster));
        const painted = await applyOverview(
          cachedOverview,
          rosterRows,
          rosterLabels,
          rosterColumns,
          true,
        );
        setIsLoading(false);
        void applyPersonnelAssets(painted.rows, rosterRows);
      }
      const paintedFromCache = Boolean(cachedOverview);

      const rosterPromise = loadLatestPersonnelRosterRows(force)
        .then((result) => {
          rosterRows = result.rows;
          rosterLabels = result.labels;
          rosterColumns = result.columns;
          setRosterLabel(result.label);
          return result.rows;
        })
        .catch(() => rosterRows);

      const overview = await fetchWithCache({
        key: CacheKeys.overview,
        force,
        fetcher: () =>
          loadPersonnelOverviewInBatches({
            onPage: paintedFromCache
              ? undefined
              : async (partial, meta) => {
                  await applyOverview(
                    partial,
                    rosterRows,
                    rosterLabels,
                    rosterColumns,
                    !meta.complete,
                  );
                  if (meta.complete) return;
                  setMessage(
                    partial.importName
                      ? `Джерело: ${partial.importName} · ${meta.done} з ${meta.total}. Довантажую…`
                      : `Огляд: ${meta.done} з ${meta.total}. Довантажую…`,
                  );
                },
          }),
        isChanged: jsonChanged,
      });
      rosterRows = await rosterPromise;
      if (!alive()) return;
      const mergedOverview = await applyOverview(
        overview,
        rosterRows,
        rosterLabels,
        rosterColumns,
      );

      void applyPersonnelAssets(mergedOverview.rows, rosterRows);
    } catch (error) {
      if (!alive()) return;
      setMessage(
        error instanceof Error ? error.message : "Не вдалося завантажити огляд",
      );
    } finally {
      if (alive()) setIsLoading(false);
    }
  };

  useEffect(() => {
    const signal = { cancelled: false };
    void load(signal);
    return () => {
      signal.cancelled = true;
    };
  }, []);

  const nameQueries = useMemo(() => parseOverviewNameQueries(query), [query]);
  const isNameListSearch = nameQueries.length > 1;

  const sourceRows = useMemo(() => {
    if (!data) return [] as BackendPersonnelOverviewRow[];
    const rows =
      source === "staff"
        ? data.rows.filter((row) => {
            if (!row.inStaff) return false;
            if (battalion === "ALL") return true;
            return (row.battalion || "") === battalion;
          })
        : source === "ejoos"
          ? data.rows.filter((row) => row.fromEjoos)
          : data.rows;
    return source === "staff" ? rows.map(withStaffOverviewStatus) : rows;
  }, [battalion, data, source]);

  const staffSummary = useMemo(
    () =>
      summarizeStaffFromRoster(
        rosterRowsRef.current,
        rosterColumnsRef.current,
        battalion,
      ),
    [battalion, rosterEpoch],
  );

  /** Pasted FIO list looks through every loaded person, not only the current source. */
  const nameSearchRows = useMemo(() => {
    if (!isNameListSearch) return sourceRows;
    if (!data) return sourceRows;
    return data.rows.map((row) =>
      row.inNovaStaff ? withStaffOverviewStatus(row) : row,
    );
  }, [data, isNameListSearch, sourceRows]);

  const sourceUnits = useMemo(
    () =>
      [...new Set(sourceRows.map((row) => row.unit).filter(Boolean))].sort((a, b) =>
        a.localeCompare(b, "uk", { numeric: true, sensitivity: "base" }),
      ),
    [sourceRows],
  );

  const metrics = useMemo(() => buildOverviewMetrics(sourceRows), [sourceRows]);

  const filteredRows = useMemo(() => {
    const maxDays = period === "ALL" ? null : Number(period);

    return nameSearchRows.filter((row) => {
      if (unit !== "ALL" && row.unit !== unit) return false;
      if (status !== "ALL" && row.status !== status) return false;
      if (
        !isNameListSearch &&
        source !== "staff" &&
        maxDays != null &&
        row.status !== "ON_DUTY" &&
        (row.days == null || row.days > maxDays)
      ) {
        return false;
      }
      if (!nameQueries.length) return true;

      if (isNameListSearch) {
        return nameQueries.some((nameQuery) =>
          overviewNameMatchesQuery(row.name, nameQuery),
        );
      }

      const normalizedQuery = normalizeRosterText(nameQueries[0] ?? "");
      if (!normalizedQuery) return true;
      if (overviewNameMatchesQuery(row.name, nameQueries[0] ?? "")) return true;

      return [
        row.name,
        row.externalId,
        row.rank,
        row.positionTitle,
        row.unit,
        row.statusLabel,
        row.fighterDirection,
        row.fighterExitDate,
        row.fighterReturnDate,
        row.fighterTotalDays,
        row.fighterStatus,
        row.externalId && documentsByExternalId[row.externalId]
          ? documentsByExternalId[row.externalId].labels.join(" ")
          : "",
      ]
        .join(" ")
        .split(" ")
        .map(normalizeRosterText)
        .join(" ")
        .includes(normalizedQuery);
    });
  }, [
    documentsByExternalId,
    isNameListSearch,
    nameQueries,
    period,
    source,
    nameSearchRows,
    status,
    unit,
  ]);

  const sideStats = useMemo(
    () => buildOverviewSideStats(filteredRows, data),
    [data, filteredRows],
  );

  const nameListMatchStats = useMemo(() => {
    if (!isNameListSearch || !data) return null;
    const matched: string[] = [];
    const missing: string[] = [];
    for (const nameQuery of nameQueries) {
      const hit = nameSearchRows.some((row) =>
        overviewNameMatchesQuery(row.name, nameQuery),
      );
      if (hit) matched.push(nameQuery);
      else missing.push(nameQuery);
    }
    return { matched, missing, total: nameQueries.length };
  }, [isNameListSearch, nameQueries, nameSearchRows]);

  const openQuestionnaire = async (target: OverviewQuestionnaireTarget) => {
    if (!target.externalId) return;
    if (!target.hasQuestionnaire) {
      openPersonnelInNewTab({
        rowId: target.rowId,
        externalId: target.externalId,
      });
      setMessage(`Відкрито картку ${target.name} — можна додати анкету.`);
      return;
    }

    try {
      setMessage(`Відкриваю анкету: ${target.name}…`);
      const url = await api.createPersonQuestionnairePreviewUrl(
        questionnaireSourceIdByExternalId[target.externalId] ??
          target.externalId,
        buildQuestionnaireExportFileName(target.name),
      );
      window.open(url, "_blank", "noopener,noreferrer");
      setMessage(`Анкета відкрита: ${target.name}`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? `Не вдалося відкрити анкету: ${error.message}`
          : "Не вдалося відкрити анкету.",
      );
    }
  };

  const onNeedPhoto = useCallback((row: BackendPersonnelOverviewRow) => {
    if (resolveOverviewPhoto(row, photosRef.current)) return;
    const requestKey = row.externalId || row.id || row.name;
    if (!requestKey || requestedPhotoKeysRef.current.has(requestKey)) return;
    requestedPhotoKeysRef.current.add(requestKey);
    void fillMissingOverviewPhotos(
      [row],
      personnelRowsRef.current.length
        ? personnelRowsRef.current
        : rosterRowsRef.current,
      photosRef.current,
    ).then((next) => {
      photosRef.current = next;
      setPhotos(next);
      if (
        !resolveOverviewPhoto(row, next) &&
        !personnelRowsRef.current.length &&
        !rosterRowsRef.current.length
      ) {
        requestedPhotoKeysRef.current.delete(requestKey);
      }
    });
  }, []);

  useEffect(() => {
    if (!filteredRows.length || filteredRows.length > 40) return;
    const missing = filteredRows.filter((row) => {
      const requestKey = row.externalId || row.id || row.name;
      if (!requestKey || requestedPhotoKeysRef.current.has(requestKey)) {
        return false;
      }
      return !resolveOverviewPhoto(row, photosRef.current);
    });
    if (!missing.length) return;
    for (const row of missing) {
      requestedPhotoKeysRef.current.add(row.externalId || row.id || row.name);
    }
    void fillMissingOverviewPhotos(
      missing,
      personnelRowsRef.current.length
        ? personnelRowsRef.current
        : rosterRowsRef.current,
      photosRef.current,
    ).then((next) => {
      photosRef.current = next;
      setPhotos(next);
    });
  }, [filteredRows, rosterEpoch]);

  const exportOverviewTable = async (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => {
    const exportedAt = new Date().toISOString().slice(0, 10);
    const visibleColumns = context.columns;
    const hasQuestionnaireColumn = visibleColumns.some(
      (column) => column.id === "questionnaire",
    );

    const sheetData: SheetData = [
        visibleColumns.map((column) => ({
          value: column.label,
          fontWeight: "bold" as const,
          align: "center" as const,
          alignVertical: "center" as const,
        })),
        ...context.rows.map((row) =>
          visibleColumns.map((column) => ({
            value: column.value(row),
            alignVertical: "center" as const,
            wrap: true,
          })),
        ),
      ];

    await writeXlsxFile(sheetData).toFile(
      `Огляд особового складу ${exportedAt}.xlsx`,
    );

    if (!hasQuestionnaireColumn) {
      setMessage(`Експортовано таблицю: ${context.rows.length} рядків.`);
      return;
    }

    const questionnaireRows = context.rows.filter(
      (row) => row.externalId && questionnaireByExternalId[row.externalId],
    );
    const files: Array<{ name: string; data: Uint8Array }> = [];

    for (const row of questionnaireRows) {
      const questionnaire = await api.getPersonQuestionnaire(
        questionnaireSourceIdByExternalId[row.externalId] ?? row.externalId,
      );
      if (!questionnaire?.fileData) continue;
      const fileName = sanitizeFileName(
        buildQuestionnaireExportFileName(
          row.name,
          row.externalId ? callSignByExternalId[row.externalId] : undefined,
        ),
      );
      files.push({
        name: fileName.toLocaleLowerCase("uk-UA").endsWith(".pdf")
          ? fileName
          : `${fileName}.pdf`,
        data: dataUrlToUint8Array(questionnaire.fileData),
      });
    }

    if (files.length) {
      downloadBlob(
        createStoredZipBlob(files),
        `Анкети огляд ${exportedAt}.zip`,
      );
    }
    setMessage(
      `Експортовано таблицю: ${context.rows.length} рядків · анкет у ZIP: ${files.length}.`,
    );
  };

  return (
    <main className="main-panel overview-page">
      <div className="overview-screen">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Огляд
          </Typography>
          <Typography variant="body2" color="text.secondary">
            {message}
          </Typography>
        </Box>
        <Stack direction="row" spacing={1}>
          <Button variant="outlined" onClick={() => void load(undefined, { force: true })}>
            Оновити
          </Button>
        </Stack>
      </header>

      {isLoading && <LinearProgress color="primary" />}

      <section className="overview-metrics">
        <article className="overview-metric-card">
          <span>
            <PersonOutlinedIcon fontSize="small" />
            {source === "staff" ? "У штаті" : "Усього"}
          </span>
          <strong>{metrics?.total ?? "—"}</strong>
          {source === "staff" && staffSummary.positions > 0 ? (
            <em className="overview-metric-note">
              штат {staffSummary.positions}
              {staffSummary.vacant > 0
                ? ` · вакант ${staffSummary.vacant}`
                : ""}
            </em>
          ) : null}
        </article>
        <article className="overview-metric-card tone-ok">
          <span>
            <ShieldOutlinedIcon fontSize="small" /> На службі
          </span>
          <strong>{metrics?.onDuty ?? "—"}</strong>
        </article>
        <article className="overview-metric-card tone-trip">
          <span>
            <BusinessCenterOutlinedIcon fontSize="small" /> Відрядження
          </span>
          <strong>{metrics?.businessTrip ?? "—"}</strong>
        </article>
        <article className="overview-metric-card tone-leave">
          <span>
            <BeachAccessOutlinedIcon fontSize="small" /> Відпустка
          </span>
          <strong>{metrics?.leave ?? "—"}</strong>
        </article>
        <article className="overview-metric-card tone-medical">
          <span>
            <LocalHospitalOutlinedIcon fontSize="small" /> Лікування
          </span>
          <strong>{metrics?.medical ?? "—"}</strong>
        </article>
      </section>

      <section className="overview-toolbar">
        <label
          className={`overview-search${isNameListSearch ? " is-multiline" : ""}`}
        >
          <SearchOutlinedIcon fontSize="small" />
          <textarea
            value={query}
            rows={isNameListSearch ? Math.min(6, nameQueries.length + 1) : 1}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="ПІБ або список (по одному в рядок) — Ctrl+V"
            spellCheck={false}
          />
          {query.trim() ? (
            <button
              type="button"
              className="overview-search-clear"
              aria-label="Очистити пошук"
              onClick={() => setQuery("")}
            >
              ✕
            </button>
          ) : null}
        </label>
        <TextField
          select
          size="small"
          className="overview-filter"
          label="Джерело"
          value={source}
          onChange={(event) => {
            const next = event.target.value as OverviewSourceFilter;
            setSource(next);
            setUnit("ALL");
            setBattalion("ALL");
          }}
        >
          {SOURCE_FILTERS.map((item) => (
            <MenuItem key={item.value} value={item.value}>
              {item.value === "staff"
                ? rosterLabel === "Штатка"
                  ? "Штатка"
                  : `Штатка · ${rosterLabel}`
                : item.value === "ejoos"
                  ? data?.importName
                    ? `ЕЖООС · ${data.importName}`
                    : "ЕЖООС"
                  : item.label}
            </MenuItem>
          ))}
        </TextField>
        {source === "staff" ? (
          <TextField
            select
            size="small"
            className="overview-filter"
            label="Батальйон"
            value={
              battalion === "ALL" || staffSummary.battalions.includes(battalion)
                ? battalion
                : "ALL"
            }
            onChange={(event) => setBattalion(event.target.value)}
          >
            <MenuItem value="ALL">Усі</MenuItem>
            {staffSummary.battalions.map((item) => (
              <MenuItem key={item} value={item}>
                {item}
              </MenuItem>
            ))}
          </TextField>
        ) : null}
        <TextField
          select
          size="small"
          className="overview-filter"
          label="Підрозділ"
          value={unit}
          onChange={(event) => setUnit(event.target.value)}
        >
          <MenuItem value="ALL">Усі підрозділи</MenuItem>
          {sourceUnits.map((item) => (
            <MenuItem key={item} value={item}>
              {item}
            </MenuItem>
          ))}
        </TextField>
        <TextField
          select
          size="small"
          className="overview-filter"
          label="Статус"
          value={status}
          onChange={(event) => setStatus(event.target.value)}
        >
          {STATUS_FILTERS.map((item) => (
            <MenuItem key={item.value} value={item.value}>
              {item.label}
            </MenuItem>
          ))}
        </TextField>
        <TextField
          select
          size="small"
          className="overview-filter"
          label="Період"
          value={period}
          onChange={(event) => setPeriod(event.target.value)}
          slotProps={{
            input: {
              startAdornment: (
                <CalendarMonthOutlinedIcon
                  fontSize="small"
                  sx={{ mr: 1, color: "text.secondary" }}
                />
              ),
            },
          }}
        >
          {PERIOD_FILTERS.map((item) => (
            <MenuItem key={item.value} value={item.value}>
              {item.label}
            </MenuItem>
          ))}
        </TextField>
      </section>

      {nameListMatchStats ? (
        <Alert
          severity={nameListMatchStats.missing.length ? "warning" : "success"}
          className="overview-name-list-alert"
        >
          Список: знайдено {nameListMatchStats.matched.length} з{" "}
          {nameListMatchStats.total}
          {nameListMatchStats.missing.length
            ? ` · не знайдено: ${nameListMatchStats.missing.join("; ")}`
            : ""}
        </Alert>
      ) : null}

      <div className="overview-table-panel">
        <OverviewVirtualTable
          rows={filteredRows}
          photos={photos}
          onNeedPhoto={onNeedPhoto}
          questionnaireByExternalId={questionnaireByExternalId}
          documentsByExternalId={documentsByExternalId}
            onOpenQuestionnaire={(target) => void openQuestionnaire(target)}
          emptyMessage={
            isLoading && !data
              ? "Завантаження огляду..."
              : "Немає записів за поточними фільтрами."
          }
          onExport={(context) => void exportOverviewTable(context)}
        />
        <footer className="overview-table-footer">
          <span>
            Показано всі {filteredRows.length} з {sourceRows.length}
          </span>
        </footer>
      </div>
      </div>

      <aside className="overview-side">
        <section className="overview-side-card overview-critical-card">
          <div className="panel-heading">Критичні терміни</div>
          <ul className="overview-critical-list">
            {sideStats.critical.map((item) => {
              const splitAt = item.text.indexOf(":");
              const name =
                splitAt >= 0 ? item.text.slice(0, splitAt).trim() : item.text;
              const meta = splitAt >= 0 ? item.text.slice(splitAt + 1).trim() : "";
              return (
                <li key={item.id} className={`tone-${item.severity}`}>
                  <span className="overview-critical-name">{name}</span>
                  {meta ? (
                    <span className="overview-critical-meta">{meta}</span>
                  ) : null}
                </li>
              );
            })}
            {sideStats.critical.length === 0 && (
              <li className="tone-info">Критичних термінів немає</li>
            )}
          </ul>
        </section>

        <section className="overview-side-card">
          <div className="panel-heading">Зміни сьогодні</div>
          <div className="overview-today-changes">
            <Chip
              className="overview-status-chip tone-ok"
              label={`+${sideStats.todayChanges.onDuty}`}
              size="small"
            />
            <Chip
              className="overview-status-chip tone-trip"
              label={`+${sideStats.todayChanges.businessTrip}`}
              size="small"
            />
            <Chip
              className="overview-status-chip tone-leave"
              label={`+${sideStats.todayChanges.leave}`}
              size="small"
            />
            <Chip
              className="overview-status-chip tone-medical"
              label={`+${sideStats.todayChanges.medical}`}
              size="small"
            />
          </div>
          <Typography variant="body2" color="text.secondary">
            Усього змін: {sideStats.todayChanges.total}
          </Typography>
        </section>

        <section className="overview-side-card overview-updates-card">
          <div className="panel-heading">Оновлення сьогодні</div>
          <strong>{String(sideStats.todayUpdates).padStart(2, "0")}</strong>
          <span>записів оновлено</span>
        </section>
      </aside>
    </main>
  );
}
