import { useMemo, useState } from "react";
import {
  Box,
  Button,
  Checkbox,
  Dialog,
  DialogActions,
  DialogContent,
  DialogTitle,
  Divider,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  personCanEnterApplyQueue,
  personIsInformationalOnly,
} from "./ejoosPersonDiff";
import {
  PersonChangeCard,
  PersonChangeRow,
} from "./EjoosPersonChangeCard";
import { useEjoosWorkspace } from "./ejoosWorkspaceState";
import { canonicalName, normKey } from "./ejoosIdentity";
import {
  planBlocksWorkbookApply,
  SOURCE_DATE_UNKNOWN_MESSAGE,
} from "./ejoosSyncPlan";

type ChangeFilter =
  | "ALL"
  | "arrival"
  | "status"
  | "position"
  | "data"
  | "error";

type ChangesView = "list" | "queue";

/** Рядки / крапка з комою — список ПІБ або ID для пошуку в операціях. */
export const parsePastedPersonList = (raw: string) =>
  raw
    .split(/[\n;]+/)
    .map((line) => line.replace(/^[-•–—\d.)\s]+/, "").trim())
    .filter(Boolean);

const personMatchesToken = (
  person: { fullName: string; personId: string; positionIndex: string },
  token: string,
) => {
  const needle = canonicalName(token) || normKey(token);
  if (!needle) return false;
  if (normKey(person.personId) === needle) return true;
  if (normKey(person.positionIndex) === needle) return true;
  const name = canonicalName(person.fullName);
  if (!name) return false;
  return name === needle || name.includes(needle) || needle.includes(name);
};

const FILTERS: { id: ChangeFilter; label: string }[] = [
  { id: "ALL", label: "Усі" },
  { id: "arrival", label: "Нові" },
  { id: "status", label: "Статус" },
  { id: "position", label: "Посада" },
  { id: "data", label: "Дані" },
  { id: "error", label: "Помилки" },
];

export function EjoosChangesPanel() {
  const {
    session,
    selectedPersonId,
    setSelectedPersonId,
    setDecision,
    dismissPerson,
    setDecisions,
    patchOpPayload,
    acceptReady,
    applyAccepted,
    acceptAndApplyPerson,
    rebuildOperations,
    setTab,
    isLoading,
  } = useEjoosWorkspace();
  const [view, setView] = useState<ChangesView>("list");
  const [filter, setFilter] = useState<ChangeFilter>("ALL");
  const [query, setQuery] = useState("");
  const [listQuery, setListQuery] = useState("");
  const [applyConfirm, setApplyConfirm] = useState<{
    id: string;
    name: string;
    reviewOnly?: boolean;
  } | null>(null);
  const [bulkApplyOpen, setBulkApplyOpen] = useState(false);

  const people = session?.people ?? [];
  const queuedPeople = useMemo(
    () => people.filter((person) => person.decision === "accepted"),
    [people],
  );
  const pastedNames = useMemo(
    () => parsePastedPersonList(listQuery),
    [listQuery],
  );
  const listHits = useMemo(() => {
    if (!pastedNames.length) return [];
    return people.filter((person) =>
      pastedNames.some((token) => personMatchesToken(person, token)),
    );
  }, [pastedNames, people]);
  const listMissing = useMemo(() => {
    if (!pastedNames.length) return [];
    return pastedNames.filter(
      (token) => !people.some((person) => personMatchesToken(person, token)),
    );
  }, [pastedNames, people]);
  const filtered = useMemo(() => {
    const source = view === "queue" ? queuedPeople : people;
    const q = query.trim().toLowerCase();
    return source.filter((person) => {
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
            kinds.has("position_change") ||
            kinds.has("shpo_occupant") ||
            kinds.has("rank_change") ||
            kinds.has("exclude_transfer");
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
      if (pastedNames.length) {
        if (!pastedNames.some((token) => personMatchesToken(person, token))) {
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
  }, [filter, pastedNames, people, query, queuedPeople, view]);

  const queueableVisible = useMemo(
    () => filtered.filter(personCanEnterApplyQueue),
    [filtered],
  );
  const visibleQueuedCount = queueableVisible.filter(
    (person) => person.decision === "accepted",
  ).length;
  const allVisibleQueued =
    queueableVisible.length > 0 &&
    visibleQueuedCount === queueableVisible.length;
  const someVisibleQueued =
    visibleQueuedCount > 0 && !allVisibleQueued;

  const selectedPerson =
    people.find((p) => p.id === selectedPersonId) ?? null;

  const writableQueuedCount = queuedPeople.filter(
    (person) => personCanEnterApplyQueue(person),
  ).length;

  const toggleQueue = (personId: string, next: boolean) => {
    setDecision(personId, next ? "accepted" : "pending");
  };

  const toggleVisibleQueue = (next: boolean) => {
    setDecisions(
      queueableVisible.map((person) => person.id),
      next ? "accepted" : "pending",
    );
    if (next) setView("queue");
  };

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
            Джерело 1ПБ станом на {session.plan.timesheetDayLabel}. Табель буде
            оновлено лише по {session.plan.timesheetDayLabel}.
          </Typography>
          <Typography variant="body2" className="ejoos-muted">
            {session.pbFileName} · {session.counters.changes} людей ·{" "}
            {session.counters.autoReady} авто · {session.counters.needsReview}{" "}
            перевірити · у черзі {queuedPeople.length}
          </Typography>
          {session.plan.sourceDateUnknown ? (
            <Typography variant="body2" sx={{ mt: 0.5, color: "#f5c16c" }}>
              {SOURCE_DATE_UNKNOWN_MESSAGE}
            </Typography>
          ) : null}
          {session.plan.sourceDateUnknown ? (
            <Stack
              direction="row"
              spacing={1}
              sx={{ mt: 1 }}
              style={{ flexWrap: "wrap", alignItems: "center" }}
            >
              <Typography variant="body2">Станом на</Typography>
              <input
                type="date"
                className="ejoos-search"
                onChange={(event) => {
                  const value = event.target.value;
                  if (!value) return;
                  void rebuildOperations(value);
                }}
              />
            </Stack>
          ) : null}
        </Box>
        <Stack direction="row" spacing={1} style={{ flexWrap: "wrap" }}>
          <Button
            size="small"
            variant="outlined"
            onClick={() => {
              acceptReady();
              setView("queue");
            }}
          >
            Додати всі «авто» в чергу
          </Button>
          <Button
            size="small"
            variant="outlined"
            disabled={isLoading}
            onClick={() => void rebuildOperations()}
          >
            Перебудувати
          </Button>
          <Button
            size="small"
            variant="contained"
            disabled={
              isLoading ||
              planBlocksWorkbookApply(session.plan) ||
              !(writableQueuedCount || queueableVisible.length)
            }
            onClick={() => {
              if (!writableQueuedCount && queueableVisible.length) {
                toggleVisibleQueue(true);
              }
              setBulkApplyOpen(true);
            }}
            sx={{ color: "#1a1a14" }}
          >
            Застосувати{" "}
            {writableQueuedCount || queueableVisible.length
              ? `(${writableQueuedCount || queueableVisible.length})`
              : ""}
          </Button>
        </Stack>
      </Stack>

      <Stack direction="row" spacing={1} style={{ flexWrap: "wrap" }}>
        <Button
          size="small"
          variant={view === "list" ? "contained" : "outlined"}
          onClick={() => setView("list")}
          sx={view === "list" ? { color: "#1a1a14" } : undefined}
        >
          Список
        </Button>
        <Button
          size="small"
          variant={view === "queue" ? "contained" : "outlined"}
          onClick={() => setView("queue")}
          sx={view === "queue" ? { color: "#1a1a14" } : undefined}
        >
          До застосування ({queuedPeople.length})
        </Button>
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
        {view === "list" &&
        (query.trim() || pastedNames.length) &&
        queueableVisible.length > 0 ? (
          <Button
            size="small"
            variant="contained"
            disabled={isLoading || planBlocksWorkbookApply(session.plan)}
            onClick={() => {
              toggleVisibleQueue(true);
              setBulkApplyOpen(true);
            }}
            sx={{ color: "#1a1a14" }}
          >
            Застосувати знайдених ({queueableVisible.length})
          </Button>
        ) : null}
      </Stack>

      <Stack spacing={0.5}>
        <textarea
          className="ejoos-search ejoos-name-list"
          value={listQuery}
          onChange={(event) => setListQuery(event.target.value)}
          rows={5}
          placeholder={
            "Вставте список ПІБ — кожного з нового рядка\nБАСОВСЬКИЙ Юрій Михайлович\nВІТКОВ Віталій Анатолійович"
          }
        />
        {pastedNames.length ? (
          <Typography variant="body2" className="ejoos-muted">
            Знайдено {listHits.length} з {pastedNames.length} у операціях
            {listMissing.length
              ? `. Немає: ${listMissing.join(", ")}`
              : "."}
          </Typography>
        ) : null}
      </Stack>

      <div className="ejoos-changes-layout">
        <div className="ejoos-change-list">
          {filtered.length === 0 ? (
            <Typography variant="body2" className="ejoos-muted" sx={{ p: 2 }}>
              {view === "queue"
                ? "Черга порожня. Знайдіть людей у списку й поставте чекбокс — вони з’являться тут."
                : "Немає змін за фільтром"}
            </Typography>
          ) : (
            <>
              {queueableVisible.length > 0 ? (
                <div className="ejoos-change-list-toolbar">
                  <Checkbox
                    id={`ejoos-queue-visible-${view}`}
                    checked={
                      allVisibleQueued
                        ? true
                        : someVisibleQueued
                          ? "indeterminate"
                          : false
                    }
                    onCheckedChange={(value) =>
                      toggleVisibleQueue(value === true)
                    }
                    label={
                      allVisibleQueued
                        ? `Зняти видимі (${queueableVisible.length})`
                        : `Обрати видимі (${queueableVisible.length})`
                    }
                  />
                </div>
              ) : null}
              {filtered.map((person) => (
                <PersonChangeRow
                  key={person.id}
                  person={person}
                  selected={person.id === selectedPersonId}
                  onSelect={() => setSelectedPersonId(person.id)}
                  queued={person.decision === "accepted"}
                  queueDisabled={!personCanEnterApplyQueue(person)}
                  onToggleQueue={(next) => toggleQueue(person.id, next)}
                />
              ))}
            </>
          )}
        </div>
        <div className="ejoos-change-detail">
          {selectedPerson ? (
            <PersonChangeCard
              person={selectedPerson}
              timesheetDay={session?.plan.timesheetDay ?? 31}
              canQueue={
                personCanEnterApplyQueue(selectedPerson) &&
                !planBlocksWorkbookApply(session.plan)
              }
              applyBlocked={planBlocksWorkbookApply(session.plan)}
              onAccept={() => {
                if (selectedPerson.decision === "accepted") {
                  setDecision(selectedPerson.id, "pending");
                  return;
                }
                setDecision(selectedPerson.id, "accepted");
                setView("queue");
              }}
              onApplyNow={() =>
                setApplyConfirm({
                  id: selectedPerson.id,
                  name: selectedPerson.fullName,
                  reviewOnly: personIsInformationalOnly(selectedPerson.ops),
                })
              }
              onReject={() => {
                if (personIsInformationalOnly(selectedPerson.ops)) {
                  dismissPerson(selectedPerson.id);
                  return;
                }
                setDecision(selectedPerson.id, "rejected");
              }}
              onClose={() => setSelectedPersonId(null)}
              onPatchPayload={(opId, patch) =>
                patchOpPayload(selectedPerson.id, opId, patch)
              }
              isLoading={isLoading}
            />
          ) : (
            <Box className="ejoos-change-card is-empty">
              <Typography variant="body2" className="ejoos-muted">
                {view === "queue"
                  ? "Черга зліва. Кнопка «Застосувати всі» запише всіх у ЕЖООС одним кроком."
                  : "Знайдіть людину, поставте чекбокс — вона перейде у «До застосування»."}
              </Typography>
            </Box>
          )}
        </div>
      </div>
      <Divider />

      <Dialog
        open={Boolean(applyConfirm)}
        onClose={() => setApplyConfirm(null)}
      >
        <DialogTitle>
          {applyConfirm?.reviewOnly ? "Підтвердити перегляд" : "Застосувати зміни"}
        </DialogTitle>
        <DialogContent>
          <Typography variant="body1">
            {applyConfirm?.reviewOnly ? (
              <>
                Позначити перевіреним{" "}
                <strong>{applyConfirm?.name}</strong>? У ЕЖООС нічого не
                зміниться.
              </>
            ) : (
              <>
                Застосувати зміни для{" "}
                <strong>{applyConfirm?.name}</strong> зараз?
              </>
            )}
          </Typography>
          <Typography variant="body2" className="ejoos-muted" sx={{ mt: 1 }}>
            {applyConfirm?.reviewOnly
              ? "Немає автоматичних змін по аркушах. Підтвердження лише прибирає запис з операцій."
              : "Буде створено нову версію ЕЖООС (ШПО, ООС, Табель, Тимч. відсутні — те, що в картці). Файл не качається — експорт з вкладки «Експорт»."}
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button
            variant="outlined"
            disabled={isLoading}
            onClick={() => setApplyConfirm(null)}
          >
            Скасувати
          </Button>
          <Button
            variant="contained"
            disabled={isLoading || !applyConfirm}
            onClick={() => {
              const personId = applyConfirm?.id;
              setApplyConfirm(null);
              if (personId) void acceptAndApplyPerson(personId);
            }}
            sx={{ color: "#1a1a14" }}
          >
            {applyConfirm?.reviewOnly ? "Підтвердити" : "Застосувати"}
          </Button>
        </DialogActions>
      </Dialog>

      <Dialog open={bulkApplyOpen} onClose={() => setBulkApplyOpen(false)}>
        <DialogTitle>Застосувати чергу</DialogTitle>
        <DialogContent>
          <Typography variant="body1">
            Застосувати зміни для{" "}
            <strong>
              {writableQueuedCount || queueableVisible.length}{" "}
              {(writableQueuedCount || queueableVisible.length) === 1
                ? "особи"
                : "осіб"}
            </strong>{" "}
            одним кроком?
          </Typography>
          <Typography variant="body2" className="ejoos-muted" sx={{ mt: 1 }}>
            Буде нова версія ЕЖООС. Після запису операції перерахуються, черга
            очиститься. Файл не качається — експорт з вкладки «Експорт».
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button
            variant="outlined"
            disabled={isLoading}
            onClick={() => setBulkApplyOpen(false)}
          >
            Скасувати
          </Button>
          <Button
            variant="contained"
            disabled={
              isLoading ||
              !(writableQueuedCount || queueableVisible.length)
            }
            onClick={() => {
              const ids = (
                writableQueuedCount ? queuedPeople : queueableVisible
              )
                .filter(personCanEnterApplyQueue)
                .map((person) => person.id);
              setDecisions(ids, "accepted");
              setBulkApplyOpen(false);
              void applyAccepted(ids);
            }}
            sx={{ color: "#1a1a14" }}
          >
            Застосувати всі
          </Button>
        </DialogActions>
      </Dialog>
    </Stack>
  );
}
