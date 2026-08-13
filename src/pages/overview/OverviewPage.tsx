import { useEffect, useMemo, useState } from "react";
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
  type BackendPersonnelOverview,
  type BackendPersonnelOverviewRow,
} from "../../api";
import {
  OverviewVirtualTable,
  type OverviewPersonTarget,
} from "./OverviewVirtualTable";

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

export function OverviewPage({
  onOpenImport,
  onOpenPersonnel,
}: {
  onOpenImport?: () => void;
  onOpenPersonnel?: (target: OverviewPersonTarget) => void;
}) {
  const [data, setData] = useState<BackendPersonnelOverview | null>(null);
  const [photos, setPhotos] = useState<Record<string, string>>({});
  const [query, setQuery] = useState("");
  const [unit, setUnit] = useState("ALL");
  const [status, setStatus] = useState("ALL");
  const [period, setPeriod] = useState("30");
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState(`API: ${api.baseUrl}`);

  const load = async () => {
    setIsLoading(true);
    try {
      const [overview, photoList] = await Promise.all([
        api.getPersonnelOverview(),
        api.listPersonPhotos().catch(() => []),
      ]);
      setData(overview);
      setPhotos(
        Object.fromEntries(
          photoList.map((item) => [item.personExternalId, item.photoData]),
        ),
      );
      setMessage(
        overview.importName
          ? `Джерело: ${overview.importName} · ООС + тимчасово відсутні`
          : "Немає імпорту ЕЖООС",
      );
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося завантажити огляд",
      );
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    void load();
  }, []);

  const filteredRows = useMemo(() => {
    if (!data) return [] as BackendPersonnelOverviewRow[];
    const normalizedQuery = query.trim().toLowerCase();
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
      return [row.name, row.externalId, row.rank, row.unit, row.statusLabel]
        .join(" ")
        .toLowerCase()
        .includes(normalizedQuery);
    });
  }, [data, period, query, status, unit]);

  const metrics = data?.metrics;

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
            onOpenPersonnel={onOpenPersonnel}
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
