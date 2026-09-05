import { memo, useMemo } from "react";
import {
  Box,
  Button,
  Checkbox,
  Chip,
  MenuItem,
  Select,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  type PersonChange,
  type PersonChangeCategory,
  buildSheetImpacts,
  buildTimesheetPreview,
  personHasWorkbookApplyOps,
  personIsInformationalOnly,
} from "./ejoosPersonDiff";
import { formatTimesheetTransferMark } from "./ejoosExcludedColumns";
import { positionCloseWritesExcluded } from "./ejoosExcludePolicy";
import { buildSheetRowPreviews } from "./ejoosSheetRowPreview";
import { timesheetOpNeedsManualCode } from "./ejoosOpRequirements";
import { EJOOS_TIMESHEET_CODES } from "./ejoosRules";
import { dayFromOrderLabel } from "./ejoosTimesheetText";
import { formatApiDateTime } from "../../shared/format";

export type PersonChangeCardMode = "review" | "history";

export type PersonChangeCardHistoryMeta = {
  version: number;
  appliedAt: string;
  pbFileName?: string;
  timesheetDayLabel?: string;
};

const categoryLabel: Record<PersonChangeCategory, string> = {
  status: "Статус",
  position: "Посада",
  arrival: "Новий",
  data: "Дані",
  error: "Помилка",
  mixed: "Змішане",
};

const timesheetMarkClass = (mark: string) => {
  const text = String(mark ?? "").trim().toLocaleLowerCase("uk-UA");
  if (text === "+") return "present";
  if (text === "-" || text === "—") return "empty";
  if (text.includes("вибув") || text.includes("перев")) return "depart";
  if (text === "зб" || text === "сзч") return "missing";
  if (text === "вп" || text === "лік" || text === "лп") return "med";
  if (text === "від" || text === "вдр") return "leave";
  return "other";
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

const formatAppliedAt = (iso: string) => formatApiDateTime(iso);

export const PersonChangeRow = memo(function PersonChangeRow({
  person,
  selected,
  onSelect,
  mode = "review",
  historyMeta,
  queued,
  onToggleQueue,
  queueDisabled,
}: {
  person: PersonChange;
  selected: boolean;
  onSelect: () => void;
  mode?: PersonChangeCardMode;
  historyMeta?: Pick<PersonChangeCardHistoryMeta, "version" | "appliedAt">;
  queued?: boolean;
  onToggleQueue?: (next: boolean) => void;
  queueDisabled?: boolean;
}) {
  const showQueue = Boolean(onToggleQueue) && mode !== "history";
  return (
    <div
      className={
        [
          "ejoos-change-row",
          selected ? "is-selected" : "",
          `severity-${person.severity}`,
          mode === "history" ? "is-history" : "",
          showQueue ? "has-check" : "",
          queued ? "is-queued" : "",
        ]
          .filter(Boolean)
          .join(" ")
      }
    >
      {showQueue ? (
        <span
          className="ejoos-change-row-check"
          onClick={(event) => event.stopPropagation()}
          onPointerDown={(event) => event.stopPropagation()}
        >
          <Checkbox
            id={`ejoos-queue-${person.id}`}
            className="ejoos-queue-checkbox"
            checked={Boolean(queued)}
            disabled={Boolean(queueDisabled) && !queued}
            onCheckedChange={(value) => onToggleQueue?.(value === true)}
          />
        </span>
      ) : null}
      <button
        type="button"
        className="ejoos-change-row-body"
        onClick={onSelect}
      >
      <div className="ejoos-change-row-main">
        <strong>{person.fullName}</strong>
        <span className="ejoos-change-meta">
          {[person.rank, person.positionIndex || person.personId]
            .filter(Boolean)
            .join(" · ")}
        </span>
        {mode === "history" && historyMeta ? (
          <span className="ejoos-history-applied-meta">
            {formatAppliedAt(historyMeta.appliedAt)}
          </span>
        ) : null}
      </div>
      <div className="ejoos-change-diff">
        <span>{person.summaryBefore}</span>
        <span aria-hidden>→</span>
        <span>{person.summaryAfter}</span>
      </div>
      <div className="ejoos-change-will">
        {mode === "history"
          ? person.ejoosWillDo.slice(0, 3).join(" · ") || "—"
          : person.ejoosWillDo[0] || "—"}
        {person.ejoosWillDo.length > (mode === "history" ? 3 : 1)
          ? ` (+${person.ejoosWillDo.length - (mode === "history" ? 3 : 1)})`
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
        {mode === "history" ? (
          <>
            <Chip
              size="small"
              label={categoryLabel[person.category]}
              variant="outlined"
            />
            <Chip size="small" color="success" label="Застосовано" />
          </>
        ) : (
          <>
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
              <Chip size="small" color="success" label="У черзі" />
            ) : null}
            {person.decision === "rejected" ? (
              <Chip size="small" label="✕" />
            ) : null}
          </>
        )}
        </div>
      </button>
    </div>
  );
});

export function PersonChangeCard({
  person,
  timesheetDay,
  mode = "review",
  historyMeta,
  onAccept,
  onApplyNow,
  onReject,
  onClose,
  onPatchPayload,
  isLoading,
  canQueue,
  applyBlocked,
}: {
  person: PersonChange;
  timesheetDay: number;
  mode?: PersonChangeCardMode;
  historyMeta?: PersonChangeCardHistoryMeta;
  onAccept?: () => void;
  onApplyNow?: () => void;
  onReject?: () => void;
  onClose: () => void;
  onPatchPayload?: (opId: string, patch: Record<string, string>) => void;
  isLoading?: boolean;
  canQueue?: boolean;
  applyBlocked?: boolean;
}) {
  const isHistory = mode === "history";
  const timesheetPreview = useMemo(
    () =>
      person.timesheetPreview ??
      buildTimesheetPreview(person.ops, timesheetDay),
    [person.ops, person.timesheetPreview, timesheetDay],
  );
  const excludeOp = person.ops.find((op) => op.kind === "exclude_transfer");
  const needsDestination = Boolean(
    excludeOp && !excludeOp.payload.destination?.trim(),
  );
  const needsExclusionDetails = Boolean(
    excludeOp &&
      (!excludeOp.payload.excludeDate?.trim() ||
        !excludeOp.payload.orderNumber?.trim() ||
        !excludeOp.payload.orderDate?.trim()),
  );
  const timesheetOps = person.ops.filter(
    (op) =>
      op.kind === "timesheet_day" && op.payload.clearStalePerson !== "1",
  );
  const needsTimesheetCode = timesheetOps.some(timesheetOpNeedsManualCode);
  const manualTimesheetOps = timesheetOps.filter(timesheetOpNeedsManualCode);
  const reviewOnly = personIsInformationalOnly(person.ops);
  const hasWorkbookApply = personHasWorkbookApplyOps(person.ops);
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

  const sheetRowPreviews = useMemo(
    () => buildSheetRowPreviews(person.ops, timesheetDay),
    [person.ops, timesheetDay],
  );

  const timesheetDepartSample = useMemo(() => {
    const phraseFromPreview = timesheetPreview?.departPhrase?.trim();
    const dayFromPreview = timesheetPreview?.departDay || 0;
    const transferOp =
      excludeOp ||
      person.ops.find((op) => op.kind === "move_to_disposition") ||
      person.ops.find(
        (op) =>
          op.kind === "position_change" &&
          positionCloseWritesExcluded(op.payload),
      );
    const phrase =
      phraseFromPreview ||
      (transferOp ? formatTimesheetTransferMark(transferOp.payload) : "");
    if (!phrase) return null;
    const day =
      dayFromPreview ||
      dayFromOrderLabel(
        transferOp?.payload.excludeDate ||
          transferOp?.payload.orderDate ||
          "",
      );
    return { day, phrase };
  }, [excludeOp, person.ops, timesheetPreview]);

  const patch = onPatchPayload ?? (() => undefined);

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
          {isHistory ? (
            <Chip
              size="small"
              color="success"
              label="Застосовано"
              sx={{ mt: 0.5 }}
            />
          ) : person.decision === "accepted" ? (
            <Chip
              size="small"
              color="success"
              label="У черзі"
              sx={{ mt: 0.5 }}
            />
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

      {person.sourceInfluences?.length ? (
        <Box sx={{ mt: 1.5 }}>
          <Typography variant="subtitle2" sx={{ mb: 0.5 }}>
            Що з 1ПБ впливає
          </Typography>
          <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
            {isHistory
              ? "Події Рух і archive, з яких зібрано ці застосовані операції."
              : "Події Рух і archive, з яких зібрано ці операції."}
          </Typography>
          <div className="ejoos-source-list">
            {person.sourceInfluences.map((item) => (
              <div
                key={`${item.source}-${item.event}`}
                className={`ejoos-source-item source-${item.source}`}
              >
                <div className="ejoos-source-item-head">
                  <span className="ejoos-source-badge">{item.sourceLabel}</span>
                  {item.ref ? (
                    <span className="ejoos-source-ref">{item.ref}</span>
                  ) : null}
                </div>
                <strong>{item.event}</strong>
                <p>{item.effect}</p>
              </div>
            ))}
          </div>
        </Box>
      ) : null}

      {timesheetPreview ? (
        <Box sx={{ mt: 1.5 }}>
          <Typography variant="subtitle2" sx={{ mb: 0.5 }}>
            Табель після застосування
          </Typography>
          <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
            {timesheetPreview.runs
              .map((run) =>
                run.from === run.to
                  ? `${String(run.from).padStart(2, "0")}: ${run.mark}`
                  : `${String(run.from).padStart(2, "0")}–${String(run.to).padStart(2, "0")}: ${run.mark}`,
              )
              .join(" · ")}
            {timesheetPreview.note && !timesheetDepartSample
              ? ` · ${timesheetPreview.note}`
              : ""}
          </Typography>
          <div className="ejoos-timesheet-preview">
            {timesheetPreview.days.map((cell) => (
              <div
                key={cell.day}
                className={`ejoos-timesheet-day mark-${timesheetMarkClass(cell.mark)}`}
                title={`${cell.day}: ${cell.title || cell.mark}`}
              >
                <span>{String(cell.day).padStart(2, "0")}</span>
                <strong>
                  {String(cell.mark || "").length > 6
                    ? `${String(cell.mark).slice(0, 5)}…`
                    : cell.mark}
                </strong>
              </div>
            ))}
          </div>
          {timesheetDepartSample ? (
            <div className="ejoos-timesheet-phrase">
              <span className="ejoos-stat-label">
                {timesheetDepartSample.day
                  ? `День ${String(timesheetDepartSample.day).padStart(2, "0")} у Табелі запишеться так`
                  : "У клітинку вибуття Табеля запишеться"}
              </span>
              <strong>{timesheetDepartSample.phrase}</strong>
              <span className="ejoos-muted">
                На сітці день позначено «ПЕРЕВ». У Excel піде саме ця фраза, не
                скорочення.
              </span>
            </div>
          ) : null}
        </Box>
      ) : null}

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
              З ЕЖООС беремо
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
          {excludeOp.payload.basisNumber || excludeOp.payload.basisDate ? (
            <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
              Документ-підстава з РУХ:{" "}
              {excludeOp.payload.basisNumber || "без номера"}
              {excludeOp.payload.basisDate
                ? ` від ${excludeOp.payload.basisDate}`
                : ""}
              . Він не замінює стройовий наказ.
            </Typography>
          ) : null}
          <div className="ejoos-transfer-fields">
            {isHistory ? (
              <>
                <label>
                  Куди вибув / документи
                  <strong>
                    {excludeOp.payload.documentsDest ||
                      excludeOp.payload.destination ||
                      "—"}
                  </strong>
                </label>
                <label>
                  Куди для табеля
                  <strong>
                    {excludeOp.payload.timesheetDestination ||
                      excludeOp.payload.destination ||
                      "—"}
                  </strong>
                </label>
                <label>
                  Дата виключення
                  <strong>{excludeOp.payload.excludeDate || "—"}</strong>
                </label>
                <label>
                  № наказу
                  <strong>{excludeOp.payload.orderNumber || "—"}</strong>
                </label>
                <label>
                  Дата наказу
                  <strong>{excludeOp.payload.orderDate || "—"}</strong>
                </label>
              </>
            ) : (
              <>
                <label>
                  Куди вибув / документи
                  <input
                    value={
                      excludeOp.payload.documentsDest ||
                      excludeOp.payload.destination ||
                      ""
                    }
                    onChange={(event) =>
                      patch(excludeOp.id, {
                        documentsDest: event.target.value,
                      })
                    }
                    placeholder="повний текст із «Яка зміна»"
                  />
                </label>
                <label>
                  Куди для табеля
                  <input
                    value={
                      excludeOp.payload.timesheetDestination ||
                      excludeOp.payload.destination ||
                      ""
                    }
                    onChange={(event) =>
                      patch(excludeOp.id, {
                        timesheetDestination: event.target.value,
                      })
                    }
                    placeholder="вибув до ..."
                  />
                </label>
                <label>
                  Дата виключення
                  <input
                    value={excludeOp.payload.excludeDate || ""}
                    onChange={(event) =>
                      patch(excludeOp.id, {
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
                      patch(excludeOp.id, {
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
                      patch(excludeOp.id, {
                        orderDate: event.target.value,
                      })
                    }
                    placeholder="дд.мм.рррр"
                  />
                </label>
              </>
            )}
          </div>
          {!isHistory && needsDestination ? (
            <Typography variant="body2" color="warning.main" sx={{ mt: 1 }}>
              Заповніть «Куди вибув» — інакше застосувати не можна.
            </Typography>
          ) : null}
          {!isHistory && needsExclusionDetails ? (
            <Typography variant="body2" color="warning.main" sx={{ mt: 1 }}>
              Вкажіть дату виключення, номер і дату стройового наказу. Дані
              розпорядження з РУХ навмисно не підставляються автоматично.
            </Typography>
          ) : null}
        </Box>
      ) : null}

      {!isHistory && manualTimesheetOps.length ? (
        <Box className="ejoos-transfer-box" sx={{ mt: 1.5 }}>
          <Typography variant="subtitle2">Код для Табеля</Typography>
          <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
            Потрібен ручний вибір — статус у sh не зводиться до одного коду
            автоматично.
          </Typography>
          <div className="ejoos-transfer-fields">
            {manualTimesheetOps.map((op) => (
              <label key={op.id}>
                День {op.payload.day || "—"}
                <Select
                  value={op.payload.timesheetCode || ""}
                  onChange={(event) =>
                    patch(op.id, {
                      timesheetCode: event.target.value,
                    })
                  }
                >
                  {EJOOS_TIMESHEET_CODES.map((code) => (
                    <MenuItem key={code} value={code}>
                      {code}
                    </MenuItem>
                  ))}
                </Select>
              </label>
            ))}
          </div>
          {needsTimesheetCode ? (
            <Typography variant="body2" color="warning.main" sx={{ mt: 1 }}>
              Виберіть код перед підтвердженням.
            </Typography>
          ) : null}
        </Box>
      ) : null}

      <Typography variant="subtitle2" sx={{ mt: 1.5, mb: 0.5 }}>
        Карта аркушів ЕЖООС
      </Typography>
      <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
        {isHistory
          ? "Що було записано в журнал для цього бійця."
          : reviewOnly
            ? "Лише перевірка джерел. У журнал при підтвердженні нічого не пишемо."
            : "Що саме зміниться в журналі при застосуванні для цього бійця."}
      </Typography>
      <div className="ejoos-sheet-impact-grid">
        {sheetImpacts.map((item) => (
          <div
            key={item.sheetKey}
            className={`ejoos-sheet-impact effect-${item.effect}`}
          >
            <div className="ejoos-sheet-impact-head">
              <strong>{item.sheetLabel}</strong>
              <span className="ejoos-sheet-impact-badge">
                {isHistory
                  ? item.effect === "append"
                    ? "Додано рядок"
                    : item.effect === "clear"
                      ? "Очищено"
                      : item.effect === "edit"
                        ? "Змінено"
                        : item.effect === "history"
                          ? "Історія + очистка"
                          : item.effect === "skip"
                            ? "Не чіпали"
                            : "Без змін"
                  : item.effectLabel}
              </span>
            </div>
            <p>{item.detail}</p>
            {item.rowHint ? (
              <span className="ejoos-sheet-impact-row">{item.rowHint}</span>
            ) : null}
          </div>
        ))}
      </div>

      {excludeOp?.payload.manualOperation === "1" ? (
        <Box className="ejoos-sheet-block" sx={{ mt: 1.5 }}>
          <Typography variant="subtitle2">
            Технічна перевірка Табеля
          </Typography>
          <Typography variant="body2" className="ejoos-muted">
            action: {excludeOp.payload.timesheetAction || "UNRESOLVED"} ·
            sourceIndex: {excludeOp.payload.timesheetSourceIndex || "—"} ·
            sourceRow: {excludeOp.payload.timesheetSourceRow || "—"} ·
            targetHistoryRow: визначається безпечно в секції під час Apply
          </Typography>
        </Box>
      ) : null}

      {sheetRowPreviews.length ? (
        <Box sx={{ mt: 1.5 }}>
          <Typography variant="subtitle2" sx={{ mb: 0.5 }}>
            Вигляд рядків після застосування
          </Typography>
          <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
            Ключові колонки Excel на аркушах, які зміняться. Порожньо = клітинку
            очищаємо; «лишається» = значення не чіпаємо.
          </Typography>
          <div className="ejoos-row-preview-list">
            {sheetRowPreviews.map((preview) => (
              <div
                key={`${preview.sheetKey}-${preview.role}`}
                className={`ejoos-row-preview sheet-${preview.sheetKey}`}
              >
                <div className="ejoos-row-preview-head">
                  <strong>{preview.sheetLabel}</strong>
                  <span>{preview.role}</span>
                </div>
                {preview.note ? (
                  <p className="ejoos-row-preview-note">{preview.note}</p>
                ) : null}
                {preview.cells.length ? (
                  <div className="ejoos-row-preview-table-wrap">
                    <table className="ejoos-row-preview-table">
                      <thead>
                        <tr>
                          {preview.cells.map((cell, index) => (
                            <th
                              key={`${cell.letter}-${cell.header}-${index}`}
                            >
                              <span>{cell.letter}</span>
                              {cell.header}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          {preview.cells.map((cell, index) => (
                            <td
                              key={`${cell.letter}-${cell.header}-v-${index}`}
                              className={`kind-${cell.kind}`}
                            >
                              {cell.value}
                            </td>
                          ))}
                        </tr>
                      </tbody>
                    </table>
                  </div>
                ) : null}
              </div>
            ))}
          </div>
        </Box>
      ) : null}

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

      {isHistory ? (
        <Stack
          className="ejoos-change-actions"
          direction="row"
          spacing={1}
          flexWrap="wrap"
          alignItems="center"
        >
          <Chip size="small" color="success" label="У журналі" />
          {historyMeta ? (
            <Typography variant="body2" className="ejoos-muted">
              Записано {formatAppliedAt(historyMeta.appliedAt)}
              {historyMeta.timesheetDayLabel
                ? ` · ${historyMeta.timesheetDayLabel}`
                : ""}
              {historyMeta.pbFileName ? ` · ${historyMeta.pbFileName}` : ""}
            </Typography>
          ) : null}
        </Stack>
      ) : (
        <Stack
          className="ejoos-change-actions"
          direction="row"
          spacing={1}
          flexWrap="wrap"
        >
          <Button
            variant="contained"
            disabled={
              Boolean(isLoading) ||
              Boolean(applyBlocked) ||
              person.severity === "conflict" ||
              needsDestination ||
              needsTimesheetCode ||
              (!reviewOnly && !hasWorkbookApply)
            }
            onClick={onApplyNow}
          >
            {reviewOnly
              ? "Підтвердити перегляд"
              : hasWorkbookApply
                ? "Застосувати зараз"
                : "Немає apply"}
          </Button>
          <Button
            variant="outlined"
            disabled={
              Boolean(isLoading) ||
              (person.decision !== "accepted" &&
                (canQueue === false ||
                  person.severity === "conflict" ||
                  needsDestination ||
                  needsTimesheetCode))
            }
            onClick={onAccept}
          >
            {person.decision === "accepted"
              ? "Прибрати з черги"
              : "Додати в чергу"}
          </Button>
          <Button
            variant="outlined"
            onClick={onReject}
            disabled={Boolean(isLoading)}
          >
            {reviewOnly ? "Прибрати" : "Відхилити"}
          </Button>
        </Stack>
      )}
    </Box>
  );
}
