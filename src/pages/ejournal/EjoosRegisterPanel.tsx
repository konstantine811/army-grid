import { useMemo, useState } from "react";
import { Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import { filterByQuery } from "./ejoosLiveViews";
import { useEjoosWorkspace } from "./EjoosWorkspaceContext";
import { useEjoosLiveView } from "./useEjoosLiveView";

type RegisterKind =
  | "absents"
  | "arrivals"
  | "excluded"
  | "timesheet"
  | "shpo";

const META: Record<
  RegisterKind,
  { title: string; empty: string }
> = {
  absents: {
    title: "Тимчасово відсутні",
    empty: "Немає відкритих періодів відсутності.",
  },
  arrivals: {
    title: "Тимчасово прибулі",
    empty: "Немає записів у «Тимч. прибулі».",
  },
  excluded: {
    title: "Виключені",
    empty: "Немає записів у «Виключені».",
  },
  timesheet: {
    title: "Табель",
    empty: "Немає рядків табеля для обраного дня.",
  },
  shpo: {
    title: "Штат / Посади",
    empty: "Немає посад у ШПО.",
  },
};

export function EjoosRegisterPanel({ kind }: { kind: RegisterKind }) {
  const { setTab } = useEjoosWorkspace();
  const { view, hasLiveFile, hasSnapshot } = useEjoosLiveView();
  const [query, setQuery] = useState("");
  const meta = META[kind];

  const rows = useMemo(() => {
    if (kind === "absents") {
      return filterByQuery(view.absentsOpen, query).map((row) => ({
        key: String(row.excelRow),
        title: row.fullName || row.personId || "—",
        meta: [row.rank, row.positionIndex, row.personId]
          .filter(Boolean)
          .join(" · "),
        detail: `${row.ground || "—"} → ${row.place || "—"} · з ${row.departDate || "?"}`,
      }));
    }
    if (kind === "arrivals") {
      return filterByQuery(view.arrivals, query).map((row) => ({
        key: String(row.excelRow),
        title: row.fullName || row.personId || "—",
        meta: [row.rank, row.positionIndex, row.personId]
          .filter(Boolean)
          .join(" · "),
        detail: `${row.fromUnit || "—"} · ${row.arriveDate || "—"}`,
      }));
    }
    if (kind === "excluded") {
      return filterByQuery(view.excluded, query).map((row) => ({
        key: String(row.excelRow),
        title: row.fullName || row.personId || "—",
        meta: [row.rank, row.personId].filter(Boolean).join(" · "),
        detail: `${row.destination || "—"} · наказ ${row.orderNumber || "—"} ${row.orderDate || ""}`.trim(),
      }));
    }
    if (kind === "timesheet") {
      return filterByQuery(view.timesheet, query).map((row) => ({
        key: String(row.excelRow),
        title: row.fullName || "(вакансія)",
        meta: [row.rank, row.positionIndex, row.personId]
          .filter(Boolean)
          .join(" · "),
        detail: `День ${view.timesheetDayLabel}: ${row.dayValue || "—"}`,
      }));
    }
    return filterByQuery(view.shpo, query).map((row) => ({
      key: String(row.excelRow),
      title: row.fullName || "(вакансія)",
      meta: [row.rank, row.positionIndex, row.personId]
        .filter(Boolean)
        .join(" · "),
      detail: row.fullName ? "Зайнято" : "Вакансія",
    }));
  }, [kind, query, view]);

  if (!hasLiveFile) {
    return (
      <Stack spacing={1.5}>
        <Typography variant="h6">{meta.title}</Typography>
        <Typography variant="body2" className="ejoos-muted">
          Спочатку завантажте канонічний ЕЖООС у БД.
        </Typography>
        <Button variant="outlined" onClick={() => setTab("import")}>
          На Головну
        </Button>
      </Stack>
    );
  }

  return (
    <Stack spacing={1.5}>
      <div>
        <Typography variant="h6">{meta.title}</Typography>
        <Typography variant="body2" className="ejoos-muted">
          Читання з live ЕЖООС · {rows.length} записів
          {kind === "timesheet" || kind === "shpo"
            ? ` · день ${view.timesheetDayLabel}`
            : ""}
          {!hasSnapshot ? " · завантаження…" : ""}
        </Typography>
      </div>

      <input
        className="ejoos-search"
        value={query}
        onChange={(event) => setQuery(event.target.value)}
        placeholder="Пошук…"
      />

      <div className="ejoos-register-list">
        {rows.map((row) => (
          <div key={row.key} className="ejoos-register-row">
            <div className="ejoos-change-row-main">
              <strong>{row.title}</strong>
              <span className="ejoos-change-meta">{row.meta}</span>
            </div>
            <div className="ejoos-change-will">{row.detail}</div>
          </div>
        ))}
        {!rows.length ? (
          <Typography variant="body2" className="ejoos-muted">
            {meta.empty}
          </Typography>
        ) : null}
      </div>
    </Stack>
  );
}
