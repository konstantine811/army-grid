import { useMemo, useState } from "react";
import {
  Box,
  Button,
  Dialog,
  DialogActions,
  DialogContent,
  DialogTitle,
  Divider,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { personIsInformationalOnly } from "./ejoosPersonDiff";
import {
  PersonChangeCard,
  PersonChangeRow,
} from "./EjoosPersonChangeCard";
import { useEjoosWorkspace } from "./ejoosWorkspaceState";

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
    rebuildOperations,
    setTab,
    isLoading,
  } = useEjoosWorkspace();
  const [filter, setFilter] = useState<ChangeFilter>("ALL");
  const [query, setQuery] = useState("");
  const [applyConfirm, setApplyConfirm] = useState<{
    id: string;
    name: string;
    reviewOnly?: boolean;
  } | null>(null);

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
  const writableAcceptedCount = people.filter(
    (person) =>
      person.decision === "accepted" && !personIsInformationalOnly(person.ops),
  ).length;

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
            variant="outlined"
            disabled={isLoading}
            onClick={() => void rebuildOperations()}
          >
            Перебудувати
          </Button>
          <Button
            size="small"
            variant="contained"
            disabled={!writableAcceptedCount || isLoading}
            onClick={() => void applyAccepted()}
            sx={{ color: "#1a1a14" }}
          >
            Застосувати підтверджені ({writableAcceptedCount || acceptedCount})
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
              <PersonChangeRow
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
            <PersonChangeCard
              person={selectedPerson}
              timesheetDay={session?.plan.timesheetDay ?? 31}
              onAccept={() => setDecision(selectedPerson.id, "accepted")}
              onApplyNow={() =>
                setApplyConfirm({
                  id: selectedPerson.id,
                  name: selectedPerson.fullName,
                  reviewOnly: personIsInformationalOnly(selectedPerson.ops),
                })
              }
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
                Застосувати підтверджені зміни для{" "}
                <strong>{applyConfirm?.name}</strong> зараз?
              </>
            )}
          </Typography>
          <Typography variant="body2" className="ejoos-muted" sx={{ mt: 1 }}>
            {applyConfirm?.reviewOnly
              ? "Це лише позначка ПІБ / ID / звання. Виправлення роблять у джерелах (1ПБ або наказ про присвоєння)."
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
    </Stack>
  );
}
