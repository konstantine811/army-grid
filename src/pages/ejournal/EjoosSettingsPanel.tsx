import { useMemo, useState } from "react";
import {
  Alert,
  Button,
  Chip,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  EJOOS_TIMESHEET_CODES,
  FIELD_SOURCE_LABELS,
  type EjoosFieldAuthority,
  type EjoosFieldSource,
  type EjoosOperatorSettings,
  type EjoosStatusConfidence,
  type EjoosStatusRule,
  type EjoosTimesheetCode,
} from "./ejoosRules";
import {
  mapPbStatusToEjoosWithRules,
  readOperatorSettings,
  resetOperatorSettings,
  writeOperatorSettings,
} from "./ejoosStatusMap";
import { useEjoosWorkspace } from "./ejoosWorkspaceState";
import { api } from "../../api";

const CONFIDENCE_OPTIONS: EjoosStatusConfidence[] = [
  "high",
  "review",
  "manual",
];

const SOURCE_OPTIONS = Object.keys(FIELD_SOURCE_LABELS) as EjoosFieldSource[];

export function EjoosSettingsPanel() {
  const { live } = useEjoosWorkspace();
  const [settings, setSettings] = useState<EjoosOperatorSettings>(() =>
    readOperatorSettings(),
  );
  const [probe, setProbe] = useState("В строю");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [savingRemote, setSavingRemote] = useState(false);
  const [normStats, setNormStats] = useState<{
    persons: number;
    absences: number;
    timesheet: number;
    syncedAt?: string;
  } | null>(null);

  const probeResult = useMemo(
    () => mapPbStatusToEjoosWithRules(probe, settings.statusRules),
    [probe, settings.statusRules],
  );

  const persistLocal = (next: EjoosOperatorSettings) => {
    const saved = writeOperatorSettings(next);
    setSettings(saved);
    setMessage("Збережено локально. Новий аналіз 1ПБ візьме ці правила.");
    setError("");
  };

  const saveRemote = async () => {
    setSavingRemote(true);
    setError("");
    try {
      const saved = writeOperatorSettings(settings);
      setSettings(saved);
      await api.putEjournalOperatorSettings({
        unitLabel: saved.unitLabel || "1ПБ",
        settings: saved,
      });
      setMessage("Налаштування збережено в БД.");
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Не вдалося зберегти налаштування в БД",
      );
    } finally {
      setSavingRemote(false);
    }
  };

  const loadRemote = async () => {
    setSavingRemote(true);
    setError("");
    try {
      const remote = await api.getEjournalOperatorSettings("1ПБ");
      if (remote?.settings) {
        const saved = writeOperatorSettings(
          remote.settings as EjoosOperatorSettings,
        );
        setSettings(saved);
        setMessage("Підтягнуто налаштування з БД.");
      } else {
        setMessage("У БД ще немає збережених налаштувань — локальні лишаються.");
      }
      const norm = await api.getEjournalNormalized("1ПБ");
      setNormStats({
        persons: norm.counts.persons,
        absences: norm.counts.absences,
        timesheet: norm.counts.timesheet,
        syncedAt: norm.syncedAt ?? undefined,
      });
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Не вдалося прочитати з БД",
      );
    } finally {
      setSavingRemote(false);
    }
  };

  const updateRule = (id: string, patch: Partial<EjoosStatusRule>) => {
    setSettings((current) => ({
      ...current,
      statusRules: current.statusRules.map((rule) =>
        rule.id === id ? { ...rule, ...patch } : rule,
      ),
    }));
  };

  const updateAuthority = (
    field: EjoosFieldAuthority["field"],
    patch: Partial<EjoosFieldAuthority>,
  ) => {
    setSettings((current) => ({
      ...current,
      fieldAuthorities: current.fieldAuthorities.map((item) =>
        item.field === field ? { ...item, ...patch } : item,
      ),
    }));
  };

  return (
    <Stack spacing={2} className="ejoos-settings">
      <div>
        <Typography variant="h6">Налаштування ЕЖООС</Typography>
        <Typography variant="body2" className="ejoos-muted">
          Правила статусів і авторитетні джерела полів. Excel лишається
          import/export; оперативна логіка — тут.
        </Typography>
      </div>

      {message ? (
        <Alert severity="success" variant="outlined">
          {message}
        </Alert>
      ) : null}
      {error ? (
        <Alert severity="error" variant="outlined">
          {error}
        </Alert>
      ) : null}

      <Stack direction="row" spacing={1} flexWrap="wrap">
        <Button variant="contained" onClick={() => persistLocal(settings)}>
          Зберегти локально
        </Button>
        <Button
          variant="outlined"
          disabled={savingRemote}
          onClick={() => void saveRemote()}
        >
          Зберегти в БД
        </Button>
        <Button
          variant="outlined"
          disabled={savingRemote}
          onClick={() => void loadRemote()}
        >
          З БД
        </Button>
        <Button
          variant="text"
          onClick={() => {
            const next = resetOperatorSettings();
            setSettings(next);
            setMessage("Скинуто до стандартних правил.");
          }}
        >
          Скинути
        </Button>
      </Stack>

      <div className="ejoos-stat-grid">
        <div className="ejoos-stat-card">
          <span className="ejoos-stat-label">Правил статусів</span>
          <strong>{settings.statusRules.filter((r) => r.enabled).length}</strong>
          <span>увімкнено з {settings.statusRules.length}</span>
        </div>
        <div className="ejoos-stat-card">
          <span className="ejoos-stat-label">Норм. БД</span>
          <strong>{normStats?.persons ?? "—"}</strong>
          <span>
            осіб
            {normStats?.syncedAt
              ? ` · ${new Date(normStats.syncedAt).toLocaleString("uk-UA")}`
              : live?.current
                ? " · синхронізується після seed/apply"
                : ""}
          </span>
        </div>
      </div>

      <section className="ejoos-settings-section">
        <Typography variant="subtitle1">Перевірка мапінгу</Typography>
        <input
          className="ejoos-search"
          value={probe}
          onChange={(event) => setProbe(event.target.value)}
          placeholder="Вставте СТАТУС з 1ПБ…"
        />
        <div className="ejoos-probe-result">
          <Chip size="small" label={probeResult.label} />
          <Chip
            size="small"
            color={
              probeResult.confidence === "high"
                ? "success"
                : probeResult.confidence === "review"
                  ? "warning"
                  : "default"
            }
            label={probeResult.timesheetCode || "без коду"}
          />
          <span className="ejoos-muted">{probeResult.reason}</span>
        </div>
      </section>

      <section className="ejoos-settings-section">
        <Typography variant="subtitle1">Правила 1ПБ → табель</Typography>
        <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
          Перше збіжне правило за пріоритетом перемагає. Невідоме → ручний вибір.
        </Typography>
        <div className="ejoos-rules-list">
          {settings.statusRules
            .slice()
            .sort((a, b) => a.priority - b.priority)
            .map((rule) => (
              <div key={rule.id} className="ejoos-rule-row">
                <label className="ejoos-check-label">
                  <input
                    type="checkbox"
                    checked={rule.enabled}
                    onChange={(event) =>
                      updateRule(rule.id, { enabled: event.target.checked })
                    }
                  />
                  <strong>{rule.label}</strong>
                </label>
                <div className="ejoos-rule-grid">
                  <label>
                    Пріоритет
                    <input
                      type="number"
                      value={rule.priority}
                      onChange={(event) =>
                        updateRule(rule.id, {
                          priority: Number(event.target.value) || 0,
                        })
                      }
                    />
                  </label>
                  <label>
                    Код
                    <select
                      value={rule.timesheetCode ?? ""}
                      onChange={(event) =>
                        updateRule(rule.id, {
                          timesheetCode: (event.target.value ||
                            null) as EjoosTimesheetCode | null,
                        })
                      }
                    >
                      <option value="">—</option>
                      {EJOOS_TIMESHEET_CODES.map((code) => (
                        <option key={code} value={code}>
                          {code}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Впевненість
                    <select
                      value={rule.confidence}
                      onChange={(event) =>
                        updateRule(rule.id, {
                          confidence: event.target
                            .value as EjoosStatusConfidence,
                        })
                      }
                    >
                      {CONFIDENCE_OPTIONS.map((item) => (
                        <option key={item} value={item}>
                          {item}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="ejoos-rule-wide">
                    matchAny (через кому)
                    <input
                      value={rule.matchAny.join(", ")}
                      onChange={(event) =>
                        updateRule(rule.id, {
                          matchAny: event.target.value
                            .split(",")
                            .map((part) => part.trim().toLowerCase())
                            .filter(Boolean),
                        })
                      }
                    />
                  </label>
                  <label className="ejoos-rule-wide">
                    Підстава відсутніх
                    <input
                      value={rule.absenceGround ?? ""}
                      onChange={(event) =>
                        updateRule(rule.id, {
                          absenceGround: event.target.value || null,
                        })
                      }
                    />
                  </label>
                </div>
              </div>
            ))}
        </div>
      </section>

      <section className="ejoos-settings-section">
        <Typography variant="subtitle1">Авторитетні джерела полів</Typography>
        <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
          Що вважаємо каноном при конфлікті між 1ПБ, ЕЖООС, карткою та анкетою.
        </Typography>
        <div className="ejoos-rules-list">
          {settings.fieldAuthorities.map((item) => (
            <div key={item.field} className="ejoos-rule-row">
              <strong>{item.label}</strong>
              <span className="ejoos-muted">{item.note}</span>
              <div className="ejoos-rule-grid">
                <label>
                  Primary
                  <select
                    value={item.primary}
                    onChange={(event) =>
                      updateAuthority(item.field, {
                        primary: event.target.value as EjoosFieldSource,
                      })
                    }
                  >
                    {SOURCE_OPTIONS.map((source) => (
                      <option key={source} value={source}>
                        {FIELD_SOURCE_LABELS[source]}
                      </option>
                    ))}
                  </select>
                </label>
                <label className="ejoos-rule-wide">
                  Fallbacks (через кому: pb_sh, ejoos, personnel_card, anketa)
                  <input
                    value={item.fallbacks.join(", ")}
                    onChange={(event) =>
                      updateAuthority(item.field, {
                        fallbacks: event.target.value
                          .split(",")
                          .map((part) => part.trim())
                          .filter((part): part is EjoosFieldSource =>
                            SOURCE_OPTIONS.includes(part as EjoosFieldSource),
                          ),
                      })
                    }
                  />
                </label>
              </div>
            </div>
          ))}
        </div>
      </section>
    </Stack>
  );
}
