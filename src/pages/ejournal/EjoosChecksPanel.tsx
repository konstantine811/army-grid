import { Button, Chip, Stack, Typography } from "@/components/sci/SciPrimitives";
import { useEjoosWorkspace } from "./ejoosWorkspaceState";
import { useEjoosLiveView } from "./useEjoosLiveView";

const severityChip = (
  severity: "ok" | "warn" | "error",
): "success" | "warning" | "error" | "default" => {
  if (severity === "ok") return "success";
  if (severity === "warn") return "warning";
  return "error";
};

export function EjoosChecksPanel() {
  const { setTab, session } = useEjoosWorkspace();
  const { view, hasLiveFile } = useEjoosLiveView();

  const blockers = view.checks.filter((item) => item.severity === "error");
  const warnings = view.checks.filter((item) => item.severity === "warn");

  return (
    <Stack spacing={2} className="ejoos-checks">
      <div>
        <Typography variant="h6">Перевірка перед експортом</Typography>
        <Typography variant="body2" className="ejoos-muted">
          Чеклист по live ЕЖООС
          {view.timesheetDayLabel ? ` · день ${view.timesheetDayLabel}` : ""}.
        </Typography>
      </div>

      {!hasLiveFile ? (
        <Typography variant="body2" className="ejoos-muted">
          Немає ЕЖООС у БД — завантажте на Головній.
        </Typography>
      ) : null}

      <div className="ejoos-stat-grid">
        <div className="ejoos-stat-card">
          <span className="ejoos-stat-label">Блокери</span>
          <strong>{blockers.length}</strong>
        </div>
        <div className="ejoos-stat-card">
          <span className="ejoos-stat-label">Попередження</span>
          <strong>{warnings.length}</strong>
        </div>
        <div className="ejoos-stat-card">
          <span className="ejoos-stat-label">В строю (день)</span>
          <strong>{view.counts.onDutyToday}</strong>
        </div>
        <div className="ejoos-stat-card">
          <span className="ejoos-stat-label">Відкриті відсутності</span>
          <strong>{view.counts.absentsOpen}</strong>
        </div>
      </div>

      <Stack spacing={1}>
        {view.checks.map((item) => (
          <div
            key={item.id}
            className={`ejoos-check-item severity-${item.severity}`}
          >
            <Chip
              size="small"
              color={severityChip(item.severity)}
              label={
                item.severity === "ok"
                  ? "OK"
                  : item.severity === "warn"
                    ? "Увага"
                    : "Блок"
              }
            />
            <div>
              <strong>{item.title}</strong>
              <span className="ejoos-muted">{item.detail}</span>
            </div>
          </div>
        ))}
      </Stack>

      <div>
        <Typography variant="subtitle1">Що змінилося сьогодні</Typography>
        <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
          {session
            ? "Підтверджені в поточній сесії 1ПБ (ще можуть бути не застосовані)."
            : "З протоколу останньої застосованої версії."}
        </Typography>
        {view.todayChanges.length ? (
          <div className="ejoos-register-list">
            {view.todayChanges.map((change, index) => (
              <div key={`${change.fullName}-${index}`} className="ejoos-register-row">
                <div className="ejoos-change-row-main">
                  <strong>{change.fullName}</strong>
                  <span className="ejoos-change-meta">{change.sheet}</span>
                </div>
                <div className="ejoos-change-diff">
                  <span>{change.before}</span>
                  <span aria-hidden>→</span>
                  <span>{change.after}</span>
                </div>
              </div>
            ))}
          </div>
        ) : (
          <Typography variant="body2" className="ejoos-muted">
            Поки немає зафіксованих змін за сьогодні.
          </Typography>
        )}
      </div>

      <Stack direction="row" spacing={1} flexWrap="wrap">
        {session ? (
          <Button variant="contained" onClick={() => setTab("import")}>
            До змін
          </Button>
        ) : null}
        <Button
          variant="outlined"
          disabled={blockers.length > 0}
          onClick={() => setTab("import")}
        >
          До експорту
        </Button>
        <Button variant="text" onClick={() => setTab("import")}>
          Особовий склад
        </Button>
      </Stack>
    </Stack>
  );
}
