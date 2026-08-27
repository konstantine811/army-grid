import { useMemo, useState } from "react";
import { Box, Button, Chip, Stack, Typography } from "@/components/sci/SciPrimitives";
import {
  parsePbArchive,
  parsePbMovements,
  parsePbShPeople,
} from "./ejoosSyncPlan";
import { useEjoosWorkspace } from "./EjoosWorkspaceContext";

type PbSourceTab = "sh" | "movements" | "archive";

type Column<T> = {
  key: keyof T;
  label: string;
  className?: string;
};

const SH_COLUMNS = [
  { key: "excelRow", label: "Рядок" },
  { key: "fullName", label: "ПІБ", className: "is-text" },
  { key: "personId", label: "ID" },
  { key: "rank", label: "Звання" },
  { key: "positionIndex", label: "Індекс" },
  { key: "positionTitle", label: "Посада", className: "is-text" },
  { key: "status", label: "Статус" },
  { key: "arrivedFrom", label: "Звідки прибув", className: "is-text" },
] as const;

const MOVEMENT_COLUMNS = [
  { key: "excelRow", label: "Рядок" },
  { key: "movementNumber", label: "№" },
  { key: "type", label: "Тип" },
  { key: "fullName", label: "ПІБ", className: "is-text" },
  { key: "personId", label: "ID" },
  { key: "rank", label: "Звання" },
  { key: "previousIndex", label: "Було" },
  { key: "nextIndex", label: "Стало" },
  { key: "destination", label: "Куди", className: "is-text" },
  { key: "orderNumber", label: "Наказ" },
  { key: "orderDate", label: "Дата" },
  { key: "status", label: "Статус" },
  { key: "note", label: "Примітка", className: "is-text" },
] as const;

const ARCHIVE_COLUMNS = [
  { key: "excelRow", label: "Рядок" },
  { key: "periodNumber", label: "№" },
  { key: "fullName", label: "ПІБ", className: "is-text" },
  { key: "personId", label: "ID" },
  { key: "rank", label: "Звання" },
  { key: "positionTitle", label: "Посада", className: "is-text" },
  { key: "absenceType", label: "Вибуття", className: "is-text" },
  { key: "departDate", label: "З дати" },
  { key: "place", label: "Куди", className: "is-text" },
  { key: "plannedReturn", label: "План" },
  { key: "returnDate", label: "Факт" },
  { key: "orderNumber", label: "Наказ" },
  { key: "orderDate", label: "Дата" },
] as const;

const SOURCE_TABS: { id: PbSourceTab; label: string }[] = [
  { id: "sh", label: "sh" },
  { id: "movements", label: "Рух" },
  { id: "archive", label: "archive" },
];

function SourceTable<T extends Record<string, unknown>>({
  columns,
  rows,
  emptyText,
}: {
  columns: readonly Column<T>[];
  rows: T[];
  emptyText: string;
}) {
  if (!rows.length) {
    return (
      <Box className="ejoos-pb-empty">
        <Typography variant="body2" className="ejoos-muted">
          {emptyText}
        </Typography>
      </Box>
    );
  }

  return (
    <div className="ejoos-pb-source-table-wrap">
      <table className="ejoos-pb-source-table">
        <thead>
          <tr>
            {columns.map((column) => (
              <th key={String(column.key)} className={column.className}>
                {column.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, index) => (
            <tr key={`${String(row.excelRow ?? "")}-${index}`}>
              {columns.map((column) => (
                <td key={String(column.key)} className={column.className}>
                  {String(row[column.key] ?? "")}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export function EjoosPbSourcePanel() {
  const {
    pbSnapshot,
    pbSources,
    session,
    isLoading,
    loadPbFromDb,
    setTab,
  } = useEjoosWorkspace();
  const [sourceTab, setSourceTab] = useState<PbSourceTab>("sh");

  const shRows = useMemo(
    () => (pbSnapshot ? parsePbShPeople(pbSnapshot) : []),
    [pbSnapshot],
  );
  const movementRows = useMemo(
    () => (pbSnapshot ? parsePbMovements(pbSnapshot) : []),
    [pbSnapshot],
  );
  const archiveRows = useMemo(
    () => (pbSnapshot ? parsePbArchive(pbSnapshot) : []),
    [pbSnapshot],
  );

  const pbCurrent = pbSources?.current;

  if (!pbSnapshot) {
    return (
      <Stack spacing={1.5} className="ejoos-pb-source">
        <Box>
          <Typography variant="h6">1ПБ</Typography>
          <Typography variant="body2" className="ejoos-muted">
            Немає відкритого 1ПБ. Завантажте файл з аркушами sh, Рух, archive
            або відкрийте останній збережений з БД.
          </Typography>
        </Box>
        <Stack direction="row" spacing={1} flexWrap="wrap">
          <Button variant="outlined" onClick={() => setTab("import")}>
            Перейти до імпорту
          </Button>
          {pbCurrent ? (
            <Button
              variant="contained"
              disabled={isLoading}
              onClick={() => void loadPbFromDb(pbCurrent.id)}
              sx={{ color: "#1a1a14" }}
            >
              Відкрити 1ПБ з БД
            </Button>
          ) : null}
        </Stack>
      </Stack>
    );
  }

  const activeRows =
    sourceTab === "sh"
      ? shRows
      : sourceTab === "movements"
        ? movementRows
        : archiveRows;
  const activeColumns =
    sourceTab === "sh"
      ? SH_COLUMNS
      : sourceTab === "movements"
        ? MOVEMENT_COLUMNS
        : ARCHIVE_COLUMNS;
  const emptyText =
    sourceTab === "sh"
      ? "Аркуш sh не знайдено або немає рядків з людьми."
      : sourceTab === "movements"
        ? "Аркуш Рух не знайдено або немає рухів."
        : "Аркуш archive не знайдено або немає періодів відсутності.";

  return (
    <Stack spacing={1.5} className="ejoos-pb-source">
      <Stack
        direction={{ xs: "column", md: "row" }}
        justifyContent="space-between"
        alignItems={{ xs: "stretch", md: "flex-start" }}
        spacing={1}
      >
        <Box>
          <Typography variant="h6">1ПБ · {pbSnapshot.fileName}</Typography>
          <Typography variant="body2" className="ejoos-muted">
            Розбір джерела для побудови змін в ЕЖООС. Видимі тут рядки ще не
            змінюють журнал, вони тільки показують, що саме прочитано з файлу.
          </Typography>
        </Box>
        <Stack direction="row" spacing={1} flexWrap="wrap">
          <Button variant="outlined" onClick={() => setTab("import")}>
            Імпорт
          </Button>
          <Button
            variant="contained"
            disabled={!session}
            onClick={() => setTab("changes")}
            sx={{ color: "#1a1a14" }}
          >
            Операції
          </Button>
        </Stack>
      </Stack>

      <div className="ejoos-stat-grid">
        <div className="ejoos-stat-card">
          <span>sh</span>
          <strong>{shRows.length}</strong>
        </div>
        <div className="ejoos-stat-card">
          <span>Рух</span>
          <strong>{movementRows.length}</strong>
        </div>
        <div className="ejoos-stat-card">
          <span>archive</span>
          <strong>{archiveRows.length}</strong>
        </div>
        <div className="ejoos-stat-card">
          <span>Операції</span>
          <strong>{session?.counters.changes ?? 0}</strong>
          {session ? (
            <span>
              авто {session.counters.autoReady} · перевірити{" "}
              {session.counters.needsReview} · конфлікти {session.counters.errors}
            </span>
          ) : (
            <span>Відкрийте ЕЖООС, щоб побудувати план.</span>
          )}
        </div>
      </div>

      <Stack direction="row" spacing={1} flexWrap="wrap">
        {SOURCE_TABS.map((item) => (
          <Button
            key={item.id}
            size="small"
            variant={sourceTab === item.id ? "contained" : "outlined"}
            onClick={() => setSourceTab(item.id)}
            sx={sourceTab === item.id ? { color: "#1a1a14" } : undefined}
          >
            {item.label}
          </Button>
        ))}
        <Chip size="small" variant="outlined" label={`${activeRows.length} рядків`} />
      </Stack>

      <SourceTable
        columns={activeColumns}
        rows={activeRows as Array<Record<string, unknown>>}
        emptyText={emptyText}
      />
    </Stack>
  );
}
