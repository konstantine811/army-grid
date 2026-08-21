import { useEffect, useMemo, useState } from "react";
import writeXlsxFile, { type SheetData } from "write-excel-file/browser";
import {
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
import { FileUploadOutlinedIcon } from "@/components/sci/icons";
import { LocalHospitalOutlinedIcon } from "@/components/sci/icons";
import { BeachAccessOutlinedIcon } from "@/components/sci/icons";
import { PersonOutlinedIcon } from "@/components/sci/icons";
import { SearchOutlinedIcon } from "@/components/sci/icons";
import { ShieldOutlinedIcon } from "@/components/sci/icons";
import {
  api,
  type BackendPersonDocument,
  type BackendPersonnelOverview,
  type BackendPersonnelOverviewRow,
} from "../../api";
import { valueToDisplay } from "../../excelRoundTrip";
import {
  createStoredZipBlob,
  dataUrlToUint8Array,
  downloadBlob,
  sanitizeFileName,
} from "../../shared/browserExport";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonIdentityFingerprint,
  buildPersonSummary,
  buildQuestionnaireExportFileName,
  getPersonExternalId,
  isUnstablePersonExternalId,
  resolvePersonIdentityKey,
  resolvePersonRankTitle,
} from "../personnel/personnelUtils";
import { PERSON_PHONES_DOCUMENT_TYPE } from "../personnel/personPhonesStore";
import {
  getRosterFighterStatusOverviewFields,
  normalizeRosterMatchText,
} from "../personnel/fighterStatusImport";
import {
  OverviewVirtualTable,
  type OverviewPersonDocumentSummary,
  type OverviewPersonTarget,
} from "./OverviewVirtualTable";
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

const normalizeRosterText = normalizeRosterMatchText;

const overviewDocumentTypeLabel = (type: string) =>
  type === "ubdReport"
    ? "Рапорт на УБД"
    : type === "form6Report"
      ? "Форма 6"
      : type === "form12Report"
        ? "Форма 12"
      : type === "ubdRestoreReport"
        ? "Рапорт на відновлення УБД"
        : type === "salaryPowerAttorney"
        ? "Довіреність зарплати"
        : type === "temporaryMilitaryId"
          ? "Тимчасовий військовий квиток"
          : type;

const buildDocumentsByExternalId = (documents: BackendPersonDocument[]) => {
  const map: Record<string, OverviewPersonDocumentSummary> = {};

  for (const document of documents) {
    if (document.type === PERSON_PHONES_DOCUMENT_TYPE) continue;
    const externalId = document.personExternalId?.trim();
    if (!externalId) continue;

    const current = map[externalId] ?? { count: 0, labels: [] };
    current.count += 1;
    const label =
      document.title?.trim() || overviewDocumentTypeLabel(document.type);
    if (label && !current.labels.includes(label)) current.labels.push(label);
    map[externalId] = current;
  }

  return map;
};

const getRosterValue = (row: EjournalPreviewRow, keyParts: string[]) => {
  const key = Object.keys(row).find((item) =>
    keyParts.every((part) => item.toLocaleLowerCase("uk-UA").includes(part)),
  );
  return key ? valueToDisplay(row[key] as Parameters<typeof valueToDisplay>[0]).trim() : "";
};

const getRosterFighterStatusFields = getRosterFighterStatusOverviewFields;

const mapRosterStatus = (status: string) => {
  const normalized = normalizeRosterText(status);
  if (normalized.includes("відряд")) {
    return { status: "BUSINESS_TRIP", statusLabel: "Відрядження" };
  }
  if (normalized.includes("відпуст")) {
    return { status: "LEAVE", statusLabel: "Відпустка" };
  }
  if (normalized.includes("ліку") || normalized.includes("шпит")) {
    return { status: "MEDICAL", statusLabel: "Лікування" };
  }
  if (normalized.includes("сзч")) {
    return { status: "AWOL", statusLabel: "СЗЧ" };
  }
  if (normalized.includes("безв") || normalized.includes("зник")) {
    return { status: "MISSING", statusLabel: "Безвісти" };
  }
  return { status: "ON_DUTY", statusLabel: "На службі" };
};

const rosterRowToOverviewRow = (
  rosterRow: EjournalPreviewRow,
): BackendPersonnelOverviewRow | null => {
  const name =
    getRosterValue(rosterRow, ["піб"]) ||
    getRosterValue(rosterRow, ["прізвище"]);
  const identityKey = resolvePersonIdentityKey({
    ...rosterRow,
    прізвище: name,
    ПІБ: name,
  });
  const rowKey = identityKey || normalizeRosterText(name);
  if (!rowKey || !name) return null;

  const mappedStatus = mapRosterStatus(getRosterValue(rosterRow, ["статус"]));
  return {
    id: `roster:${rowKey}`,
    externalId: identityKey,
    name,
    rank: resolvePersonRankTitle(rosterRow) || getRosterValue(rosterRow, ["звання"]),
    unit:
      getRosterValue(rosterRow, ["перебування"]) ||
      getRosterValue(rosterRow, ["підрозділ"]) ||
      "—",
    status: mappedStatus.status,
    statusLabel: mappedStatus.statusLabel,
    validFrom: null,
    days: null,
    plannedReturn: null,
    place: "",
    updatedAt: "",
    ...getRosterFighterStatusFields(rosterRow),
  };
};

const buildOverviewMetrics = (rows: BackendPersonnelOverviewRow[]) => ({
  total: rows.length,
  onDuty: rows.filter((row) => row.status === "ON_DUTY").length,
  businessTrip: rows.filter((row) => row.status === "BUSINESS_TRIP").length,
  leave: rows.filter((row) => row.status === "LEAVE").length,
  medical: rows.filter((row) => row.status === "MEDICAL").length,
  awol: rows.filter((row) => row.status === "AWOL").length,
  other: rows.filter(
    (row) =>
      !["ON_DUTY", "BUSINESS_TRIP", "LEAVE", "MEDICAL", "AWOL"].includes(
        String(row.status),
      ),
  ).length,
});

const mergeRosterRowsIntoOverview = (
  overview: BackendPersonnelOverview,
  rosterRows: EjournalPreviewRow[],
): BackendPersonnelOverview => {
  if (!rosterRows.length) return overview;

  const rosterById = new Map<string, EjournalPreviewRow>();
  const rosterByName = new Map<string, EjournalPreviewRow>();
  const usedRosterRows = new Set<EjournalPreviewRow>();
  rosterRows.forEach((row) => {
    const id = getPersonExternalId(row);
    const name =
      getRosterValue(row, ["піб"]) ||
      getRosterValue(row, ["прізвище"]);
    if (id) rosterById.set(id, row);
    if (name) rosterByName.set(normalizeRosterText(name), row);
  });

  const mergedRows = overview.rows.map((row) => {
    const rosterRow =
      (row.externalId &&
        !isUnstablePersonExternalId(row.externalId) &&
        rosterById.get(row.externalId)) ||
      rosterByName.get(normalizeRosterText(row.name));
    const stableFromRoster = rosterRow
      ? resolvePersonIdentityKey(rosterRow)
      : "";
    const stableFromOverview =
      row.externalId && !isUnstablePersonExternalId(row.externalId)
        ? row.externalId
        : "";
    const fingerprint = buildPersonIdentityFingerprint(row.name);
    const externalId =
      stableFromRoster || stableFromOverview || fingerprint || row.externalId;

    if (!rosterRow) {
      return externalId === row.externalId ? row : { ...row, externalId };
    }

    usedRosterRows.add(rosterRow);
    let rosterRank = "";
    try {
      rosterRank = resolvePersonRankTitle(rosterRow);
    } catch {
      rosterRank = "";
    }
    return {
      ...row,
      externalId,
      ...(rosterRank ? { rank: rosterRank } : {}),
      ...getRosterFighterStatusFields(rosterRow),
    };
  });

  const rosterOnlyRows = rosterRows
    .filter((row) => !usedRosterRows.has(row))
    .filter((row) => {
      const name =
        getRosterValue(row, ["піб"]) ||
        getRosterValue(row, ["прізвище"]);
      return Boolean(name);
    })
    .map(rosterRowToOverviewRow)
    .filter((row): row is BackendPersonnelOverviewRow => Boolean(row));

  if (!rosterOnlyRows.length) {
    return {
      ...overview,
      rows: mergedRows,
    };
  }

  const rows = [...mergedRows, ...rosterOnlyRows];
  return {
    ...overview,
    rows,
    metrics: buildOverviewMetrics(rows),
    units: [...new Set(rows.map((row) => row.unit).filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, "uk", { numeric: true, sensitivity: "base" }),
    ),
  };
};

const loadLatestPersonnelRosterRows = async () => {
  const latest = await api.getLatestPersonnelRoster();
  if (!latest?.sheet) return [] as EjournalPreviewRow[];

  return latest.rows.map((row) => ({
    __dbRowId: row.id,
    __rowNumber: row.excelRowNumber,
    ...(row.values && typeof row.values === "object" && !Array.isArray(row.values)
      ? row.values
      : {}),
  })) as EjournalPreviewRow[];
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

export function OverviewPage({
  onOpenImport,
  onOpenPersonnel,
}: {
  onOpenImport?: () => void;
  onOpenPersonnel?: (target: OverviewPersonTarget) => void;
}) {
  const [data, setData] = useState<BackendPersonnelOverview | null>(null);
  const [photos, setPhotos] = useState<Record<string, string>>({});
  const [questionnaireByExternalId, setQuestionnaireByExternalId] = useState<
    Record<string, true>
  >({});
  const [documentsByExternalId, setDocumentsByExternalId] = useState<
    Record<string, OverviewPersonDocumentSummary>
  >({});
  const [callSignByExternalId, setCallSignByExternalId] = useState<
    Record<string, string>
  >({});
  const [query, setQuery] = useState("");
  const [unit, setUnit] = useState("ALL");
  const [status, setStatus] = useState("ALL");
  const [period, setPeriod] = useState("30");
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState(`API: ${api.baseUrl}`);

  const load = async () => {
    setIsLoading(true);
    try {
      const [overview, photoList, questionnaireList, rosterRows] = await Promise.all([
        api.getPersonnelOverview(),
        api.listPersonPhotos().catch(() => []),
        api.listPersonQuestionnaires().catch(() => []),
        loadLatestPersonnelRosterRows().catch(() => [] as EjournalPreviewRow[]),
      ]);
      let mergedOverview = overview;
      try {
        mergedOverview = mergeRosterRowsIntoOverview(overview, rosterRows);
      } catch {
        mergedOverview = overview;
      }
      setData(mergedOverview);
      setPhotos(
        Object.fromEntries(
          photoList.map((item) => [item.personExternalId, item.photoData]),
        ),
      );
      setQuestionnaireByExternalId(
        Object.fromEntries(
          questionnaireList
            .filter((item) => item.personExternalId)
            .map((item) => [item.personExternalId, true as const]),
        ),
      );
      try {
        setCallSignByExternalId(buildCallSignByExternalId(rosterRows));
      } catch {
        setCallSignByExternalId({});
      }
      setMessage(
        mergedOverview.importName
          ? `Джерело: ${mergedOverview.importName} · ООС + Загальний список`
          : "Немає імпорту ЕЖООС",
      );
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося завантажити огляд",
      );
    } finally {
      setIsLoading(false);
    }

    try {
      const documentList = await api.listAllPersonDocuments();
      setDocumentsByExternalId(buildDocumentsByExternalId(documentList));
    } catch {
      setDocumentsByExternalId({});
    }
  };

  useEffect(() => {
    void load();
  }, []);

  const filteredRows = useMemo(() => {
    if (!data) return [] as BackendPersonnelOverviewRow[];
    const normalizedQuery = normalizeRosterText(query);
    const maxDays = period === "ALL" ? null : Number(period);

    return data.rows.filter((row) => {
      if (unit !== "ALL" && row.unit !== unit) return false;
      if (status !== "ALL" && row.status !== status) return false;
      if (
        maxDays != null &&
        row.status !== "ON_DUTY" &&
        (row.days == null || row.days > maxDays)
      ) {
        return false;
      }
      if (!normalizedQuery) return true;
      return [
        row.name,
        row.externalId,
        row.rank,
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
  }, [data, documentsByExternalId, period, query, status, unit]);

  const metrics = data?.metrics;

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
      const questionnaire = await api.getPersonQuestionnaire(row.externalId);
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
          <Button variant="outlined" onClick={() => void load()}>
            Оновити
          </Button>
          <Button
            variant="contained"
            startIcon={<FileUploadOutlinedIcon />}
            onClick={onOpenImport}
            sx={{ color: "#1a1a14" }}
          >
            Імпортувати
          </Button>
        </Stack>
      </header>

      {isLoading && <LinearProgress color="primary" />}

      <section className="overview-metrics">
        <article className="overview-metric-card">
          <span>
            <PersonOutlinedIcon fontSize="small" /> Усього
          </span>
          <strong>{metrics?.total ?? "—"}</strong>
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
        <label className="overview-search">
          <SearchOutlinedIcon fontSize="small" />
          <input
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="Пошук за ПІБ"
          />
        </label>
        <TextField
          select
          size="small"
          label="Підрозділ"
          value={unit}
          onChange={(event) => setUnit(event.target.value)}
        >
          <MenuItem value="ALL">Усі підрозділи</MenuItem>
          {(data?.units ?? []).map((item) => (
            <MenuItem key={item} value={item}>
              {item}
            </MenuItem>
          ))}
        </TextField>
        <TextField
          select
          size="small"
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

      <section className="overview-layout">
        <div className="overview-table-panel">
          <OverviewVirtualTable
            rows={filteredRows}
            photos={photos}
            questionnaireByExternalId={questionnaireByExternalId}
            documentsByExternalId={documentsByExternalId}
            onOpenPersonnel={onOpenPersonnel}
            emptyMessage={
              isLoading && !data
                ? "Завантаження огляду..."
                : "Немає записів за поточними фільтрами."
            }
            onExport={(context) => void exportOverviewTable(context)}
          />
          <footer className="overview-table-footer">
            <span>
              Показано всі {filteredRows.length} з {data?.rows.length ?? 0}
            </span>
          </footer>
        </div>

        <aside className="overview-side">
          <section className="overview-side-card overview-critical-card">
            <div className="panel-heading">Критичні терміни</div>
            <ul className="overview-critical-list">
              {(data?.critical ?? []).map((item) => (
                <li key={item.id} className={`tone-${item.severity}`}>
                  {item.text}
                </li>
              ))}
              {(data?.critical?.length ?? 0) === 0 && (
                <li className="tone-info">Критичних термінів немає</li>
              )}
            </ul>
          </section>

          <section className="overview-side-card">
            <div className="panel-heading">Зміни сьогодні</div>
            <div className="overview-today-changes">
              <Chip
                className="overview-status-chip tone-ok"
                label={`+${data?.todayChanges.onDuty ?? 0}`}
                size="small"
              />
              <Chip
                className="overview-status-chip tone-trip"
                label={`+${data?.todayChanges.businessTrip ?? 0}`}
                size="small"
              />
              <Chip
                className="overview-status-chip tone-leave"
                label={`+${data?.todayChanges.leave ?? 0}`}
                size="small"
              />
              <Chip
                className="overview-status-chip tone-medical"
                label={`+${data?.todayChanges.medical ?? 0}`}
                size="small"
              />
            </div>
            <Typography variant="body2" color="text.secondary" sx={{ mt: 1 }}>
              Усього змін: {data?.todayChanges.total ?? 0}
            </Typography>
          </section>

          <section className="overview-side-card overview-updates-card">
            <div className="panel-heading">Оновлення сьогодні</div>
            <strong>{String(data?.todayUpdates ?? 0).padStart(2, "0")}</strong>
            <span>записів оновлено</span>
          </section>
        </aside>
      </section>
    </main>
  );
}
