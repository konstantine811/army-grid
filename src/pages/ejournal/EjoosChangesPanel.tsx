import { memo, useMemo, useState } from "react";
import {
  Box,
  Button,
  Chip,
  Divider,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  type PersonChange,
  type PersonChangeCategory,
  buildSheetImpacts,
} from "./ejoosPersonDiff";
import { useEjoosWorkspace } from "./EjoosWorkspaceContext";

type ChangeFilter =
  | "ALL"
  | "arrival"
  | "status"
  | "position"
  | "data"
  | "error";

const FILTERS: { id: ChangeFilter; label: string }[] = [
  { id: "ALL", label: "Усі" },
  { id: "arrival", label: "Нові" },
  { id: "status", label: "Статус" },
  { id: "position", label: "Посада" },
  { id: "data", label: "Дані" },
  { id: "error", label: "Помилки" },
];

const categoryLabel: Record<PersonChangeCategory, string> = {
  status: "Статус",
  position: "Посада",
  arrival: "Новий",
  data: "Дані",
  error: "Помилка",
  mixed: "Змішане",
};

const severityColor = (
  severity: PersonChange["severity"],
): "success" | "warning" | "error" | "default" => {
  if (severity === "ready") return "success";
  if (severity === "needs_input") return "warning";
  return "error";
};

const severityLabel = (severity: PersonChange["severity"]) => {
  if (severity === "ready") return "Авто";
  if (severity === "needs_input") return "Перевірити";
  return "Конфлікт";
};

const PersonRow = memo(function PersonRow({
  person,
  selected,
  onSelect,
}: {
  person: PersonChange;
  selected: boolean;
  onSelect: () => void;
}) {
  return (
    <button
      type="button"
      className={
        selected
          ? `ejoos-change-row is-selected severity-${person.severity}`
          : `ejoos-change-row severity-${person.severity}`
      }
      onClick={onSelect}
    >
      <div className="ejoos-change-row-main">
        <strong>{person.fullName}</strong>
        <span className="ejoos-change-meta">
          {[person.rank, person.positionIndex || person.personId]
            .filter(Boolean)
            .join(" · ")}
        </span>
      </div>
      <div className="ejoos-change-diff">
        <span>{person.summaryBefore}</span>
        <span aria-hidden>→</span>
        <span>{person.summaryAfter}</span>
      </div>
      <div className="ejoos-change-will">
        {person.ejoosWillDo[0] || "—"}
        {person.ejoosWillDo.length > 1
          ? ` (+${person.ejoosWillDo.length - 1})`
          : ""}
      </div>
      <div className="ejoos-change-sheet-chips">
        {(person.sheetImpacts ?? buildSheetImpacts(person.ops))
          .filter(
            (item) =>
              item.effect !== "untouched" && item.effect !== "skip",
          )
          .map((item) => (
            <span
              key={item.sheetKey}
              className={`ejoos-mini-chip effect-${item.effect}`}
              title={item.detail}
            >
              {item.sheetLabel.replace(/^\d+\.\s*/, "")}
            </span>
          ))}
      </div>
      <div className="ejoos-change-badges">
        <Chip
          size="small"
          label={categoryLabel[person.category]}
          variant="outlined"
        />
        <Chip
          size="small"
          color={severityColor(person.severity)}
          label={severityLabel(person.severity)}
        />
        {person.decision === "accepted" ? (
          <Chip size="small" color="success" label="✓" />
        ) : null}
        {person.decision === "rejected" ? (
          <Chip size="small" label="✕" />
        ) : null}
      </div>
    </button>
  );
});

function PersonDetail({
  person,
  onAccept,
  onApplyNow,
  onReject,
  onClose,
  onPatchPayload,
  isLoading,
}: {
  person: PersonChange;
  onAccept: () => void;
  onApplyNow: () => void;
  onReject: () => void;
  onClose: () => void;
  onPatchPayload: (opId: string, patch: Record<string, string>) => void;
  isLoading: boolean;
}) {
  const excludeOp = person.ops.find((op) => op.kind === "exclude_transfer");
  const needsDestination = Boolean(
    excludeOp && !excludeOp.payload.destination?.trim(),
  );
  const bySheet = useMemo(() => {
    const map = new Map<string, typeof person.sheetActions>();
    person.sheetActions.forEach((action) => {
      const list = map.get(action.sheet) ?? [];
      list.push(action);
      map.set(action.sheet, list);
    });
    return [...map.entries()];
  }, [person.sheetActions]);

  const sheetImpacts = useMemo(
    () =>
      person.sheetImpacts?.length
        ? person.sheetImpacts
        : buildSheetImpacts(person.ops),
    [person.ops, person.sheetImpacts],
  );

  return (
    <Box className="ejoos-change-card">
      <Stack
        direction="row"
        justifyContent="space-between"
        alignItems="flex-start"
        spacing={1}
      >
        <Box>
          <Typography variant="h6">{person.fullName}</Typography>
          <Typography variant="body2" className="ejoos-muted">
            {[person.rank, person.positionIndex, person.personId]
              .filter(Boolean)
              .join(" · ")}
          </Typography>
          {person.decision === "accepted" ? (
            <Chip size="small" color="success" label="Підтверджено" sx={{ mt: 0.5 }} />
          ) : null}
        </Box>
        <Button size="small" onClick={onClose}>
          Закрити
        </Button>
      </Stack>

      <div className="ejoos-change-summary-pair">
        <div>
          <span className="ejoos-stat-label">Було (ЕЖООС)</span>
          <strong>{person.summaryBefore}</strong>
        </div>
        <div>
          <span className="ejoos-stat-label">Стало (1ПБ)</span>
          <strong>{person.summaryAfter}</strong>
        </div>
      </div>

      {excludeOp ? (
        <Box className="ejoos-transfer-box" sx={{ mt: 1.5 }}>
          <Typography variant="subtitle2">
            Переведення з обліку (ПЕРЕВ → Виключені)
          </Typography>
          <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
            Не внутрішня зміна посади. Алгоритм: Виключені → Табель → ШПО/ООС.
            Тимч. відсутні/прибулі не чіпаємо.
          </Typography>
          <ol className="ejoos-transfer-steps">
            {person.ejoosWillDo.map((step) => (
              <li key={step}>{step}</li>
            ))}
          </ol>
          <div className="ejoos-sheet-block" style={{ marginBottom: "0.75rem" }}>
            <Typography variant="caption" className="ejoos-muted">
              З ШПО беремо
            </Typography>
            <strong>
              {[
                excludeOp.payload.fromRank || excludeOp.rank,
                excludeOp.payload.fromName || excludeOp.fullName,
                excludeOp.payload.fromPositionIndex || excludeOp.positionIndex,
                excludeOp.payload.fromPersonId || excludeOp.personId,
              ]
                .filter(Boolean)
                .join(" · ") || "— (ШПО не знайдено)"}
            </strong>
            {excludeOp.payload.shpoExcelRow ? (
              <span className="ejoos-muted">
                рядок ШПО {excludeOp.payload.shpoExcelRow}
              </span>
            ) : (
              <span className="ejoos-muted">
                ШПО не знайдено автоматично — ПІБ/ID з Рух все одно застосуються
              </span>
            )}
          </div>
          <div className="ejoos-transfer-fields">
            <label>
              Куди вибув / документи
              <input
                value={excludeOp.payload.destination || ""}
                onChange={(event) =>
                  onPatchPayload(excludeOp.id, {
                    destination: event.target.value,
                    documentsDest: event.target.value,
                  })
                }
                placeholder="обовʼязково: куди з Рух"
              />
            </label>
            <label>
              Дата виключення
              <input
                value={
                  excludeOp.payload.excludeDate ||
                  excludeOp.payload.orderDate ||
                  ""
                }
                onChange={(event) =>
                  onPatchPayload(excludeOp.id, {
                    excludeDate: event.target.value,
                  })
                }
                placeholder="дд.мм.рррр"
              />
            </label>
            <label>
              № наказу
              <input
                value={excludeOp.payload.orderNumber || ""}
                onChange={(event) =>
                  onPatchPayload(excludeOp.id, {
                    orderNumber: event.target.value,
                  })
                }
              />
            </label>
            <label>
              Дата наказу
              <input
                value={excludeOp.payload.orderDate || ""}
                onChange={(event) =>
                  onPatchPayload(excludeOp.id, {
                    orderDate: event.target.value,
                  })
                }
                placeholder="дд.мм.рррр"
              />
            </label>
          </div>
          {needsDestination ? (
            <Typography variant="body2" color="warning.main" sx={{ mt: 1 }}>
              Заповніть «Куди вибув» — інакше застосувати не можна.
            </Typography>
          ) : null}
        </Box>
      ) : null}

      <Typography variant="subtitle2" sx={{ mt: 1.5, mb: 0.5 }}>
        Карта аркушів ЕЖООС
      </Typography>
      <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
        Що саме зміниться в журналі при застосуванні для цього бійця.
      </Typography>
      <div className="ejoos-sheet-impact-grid">
        {sheetImpacts.map((item) => (
          <div
            key={item.sheetKey}
            className={`ejoos-sheet-impact effect-${item.effect}`}
          >
            <div className="ejoos-sheet-impact-head">
              <strong>{item.sheetLabel}</strong>
              <span className="ejoos-sheet-impact-badge">{item.effectLabel}</span>
            </div>
            <p>{item.detail}</p>
            {item.rowHint ? (
              <span className="ejoos-sheet-impact-row">{item.rowHint}</span>
            ) : null}
          </div>
        ))}
      </div>

      <Typography variant="subtitle2" sx={{ mt: 1.5, mb: 0.5 }}>
        Деталі ops
      </Typography>
      <Stack spacing={1}>
        {bySheet.map(([sheet, actions]) => (
          <Box key={sheet} className="ejoos-sheet-block">
            <Typography variant="caption" className="ejoos-muted">
              {sheet}
            </Typography>
            {actions.map((action) => (
              <div key={action.opId} className="ejoos-sheet-action">
                <span>
                  {action.before} → {action.after}
                </span>
                <span className="ejoos-sheet-why">{action.why}</span>
              </div>
            ))}
          </Box>
        ))}
      </Stack>

      <Stack direction="row" spacing={1} sx={{ mt: 2 }} flexWrap="wrap">
        {excludeOp ? (
          <Button
            variant="contained"
            disabled={
              isLoading ||
              person.severity === "conflict" ||
              needsDestination
            }
            onClick={onApplyNow}
            sx={{ color: "#1a1a14" }}
          >
            Застосувати переведення зараз
          </Button>
        ) : (
          <Button
            variant="contained"
            disabled={person.severity === "conflict" || isLoading}
            onClick={onAccept}
            sx={{ color: "#1a1a14" }}
          >
            Підтвердити
          </Button>
        )}
        {!excludeOp ? null : (
          <Button
            variant="outlined"
            disabled={person.severity === "conflict" || needsDestination || isLoading}
            onClick={onAccept}
          >
            Лише в чергу
          </Button>
        )}
        <Button variant="outlined" onClick={onReject} disabled={isLoading}>
          Відхилити
        </Button>
      </Stack>
    </Box>
  );
}

export function EjoosChangesPanel() {
  const {
    session,
    selectedPersonId,
    setSelectedPersonId,
    setDecision,
    patchOpPayload,
    acceptReady,
    applyAccepted,
    acceptAndApplyPerson,
    setTab,
    isLoading,
  } = useEjoosWorkspace();
  const [filter, setFilter] = useState<ChangeFilter>("ALL");
  const [query, setQuery] = useState("");

  const people = session?.people ?? [];
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return people.filter((person) => {
      if (filter === "error") {
        if (person.severity !== "conflict" && person.category !== "error") {
          return false;
        }
      } else if (filter !== "ALL") {
        if (person.category === filter) {
          /* exact match */
        } else if (
          filter === "data" &&
          (person.category === "data" || person.category === "mixed")
        ) {
          /* data + mixed */
        } else if (person.category === "mixed") {
          const kinds = new Set(person.ops.map((op) => op.kind));
          const statusHit =
            kinds.has("timesheet_day") ||
            kinds.has("absent_upsert") ||
            kinds.has("absent_close");
          const positionHit =
            kinds.has("position_change") || kinds.has("shpo_occupant");
          if (filter === "status" && !statusHit) return false;
          else if (filter === "position" && !positionHit) return false;
          else if (filter === "arrival" && !kinds.has("arrival")) return false;
          else if (
            filter !== "status" &&
            filter !== "position" &&
            filter !== "arrival"
          ) {
            return false;
          }
        } else {
          return false;
        }
      }
      if (!q) return true;
      return (
        person.fullName.toLowerCase().includes(q) ||
        person.personId.toLowerCase().includes(q) ||
        person.positionIndex.toLowerCase().includes(q)
      );
    });
  }, [people, filter, query]);

  const selectedPerson =
    people.find((p) => p.id === selectedPersonId) ?? null;

  const acceptedCount = people.filter((p) => p.decision === "accepted").length;

  if (!session) {
    return (
      <Stack spacing={1.5}>
        <Typography variant="h6">Зміни</Typography>
        <Typography variant="body2" className="ejoos-muted">
          Немає активного аналізу. Завантажте 1ПБ — з’явиться список змін по
          людях.
        </Typography>
        <Button variant="outlined" onClick={() => setTab("import")}>
          Перейти до імпорту 1ПБ
        </Button>
      </Stack>
    );
  }

  return (
    <Stack spacing={1.5} className="ejoos-changes">
      <Stack
        direction={{ xs: "column", md: "row" }}
        justifyContent="space-between"
        alignItems={{ xs: "stretch", md: "center" }}
        spacing={1}
      >
        <Box>
          <Typography variant="h6">
            Зміни · {session.plan.timesheetDayLabel}
          </Typography>
          <Typography variant="body2" className="ejoos-muted">
            {session.pbFileName} · {session.counters.changes} людей ·{" "}
            {session.counters.autoReady} авто · {session.counters.needsReview}{" "}
            перевірити · підтверджено {acceptedCount}
          </Typography>
        </Box>
        <Stack direction="row" spacing={1} style={{ flexWrap: "wrap" }}>
          <Button size="small" variant="outlined" onClick={acceptReady}>
            Підтвердити всі «зелені»
          </Button>
          <Button
            size="small"
            variant="contained"
            disabled={!acceptedCount || isLoading}
            onClick={() => void applyAccepted()}
            sx={{ color: "#1a1a14" }}
          >
            Застосувати підтверджені ({acceptedCount})
          </Button>
        </Stack>
      </Stack>

      <Stack
        direction="row"
        spacing={1}
        style={{ flexWrap: "wrap", alignItems: "center" }}
      >
        {FILTERS.map((item) => (
          <Button
            key={item.id}
            size="small"
            variant={filter === item.id ? "contained" : "outlined"}
            onClick={() => setFilter(item.id)}
            sx={filter === item.id ? { color: "#1a1a14" } : undefined}
          >
            {item.label}
          </Button>
        ))}
        <input
          className="ejoos-search"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder="Пошук ПІБ / ID / індекс"
        />
      </Stack>

      <div className="ejoos-changes-layout">
        <div className="ejoos-change-list">
          {filtered.length === 0 ? (
            <Typography variant="body2" className="ejoos-muted" sx={{ p: 2 }}>
              Немає змін за фільтром
            </Typography>
          ) : (
            filtered.map((person) => (
              <PersonRow
                key={person.id}
                person={person}
                selected={person.id === selectedPersonId}
                onSelect={() => setSelectedPersonId(person.id)}
              />
            ))
          )}
        </div>
        <div className="ejoos-change-detail">
          {selectedPerson ? (
            <PersonDetail
              person={selectedPerson}
              onAccept={() => setDecision(selectedPerson.id, "accepted")}
              onApplyNow={() => void acceptAndApplyPerson(selectedPerson.id)}
              onReject={() => setDecision(selectedPerson.id, "rejected")}
              onClose={() => setSelectedPersonId(null)}
              onPatchPayload={(opId, patch) =>
                patchOpPayload(selectedPerson.id, opId, patch)
              }
              isLoading={isLoading}
            />
          ) : (
            <Box className="ejoos-change-card is-empty">
              <Typography variant="body2" className="ejoos-muted">
                Оберіть людину зі списку, щоб побачити було / стало і дії по
                аркушах.
              </Typography>
            </Box>
          )}
        </div>
      </div>
      <Divider />
    </Stack>
  );
}
