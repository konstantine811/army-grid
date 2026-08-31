import { useMemo, useState } from "react";
import {
  Box,
  Button,
  Chip,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  filterByQuery,
  findPersonAbsents,
  type EjoosRegisterPerson,
} from "./ejoosLiveViews";
import { FIELD_SOURCE_LABELS, type EjoosFieldAuthority } from "./ejoosRules";
import { readOperatorSettings } from "./ejoosStatusMap";
import { useEjoosWorkspace } from "./ejoosWorkspaceState";
import { useEjoosLiveView } from "./useEjoosLiveView";

type CardTab = "main" | "service" | "absents" | "timesheet" | "history";

const CARD_TABS: { id: CardTab; label: string }[] = [
  { id: "main", label: "Основне" },
  { id: "service", label: "Служба" },
  { id: "absents", label: "Відсутності" },
  { id: "timesheet", label: "Табель" },
  { id: "history", label: "Історія" },
];

export function EjoosRosterPanel() {
  const { session, setTab } = useEjoosWorkspace();
  const { view, hasLiveFile, hasSnapshot } = useEjoosLiveView();
  const [query, setQuery] = useState("");
  const [onlyOccupied, setOnlyOccupied] = useState(true);
  const [selectedKey, setSelectedKey] = useState<string | null>(null);
  const [cardTab, setCardTab] = useState<CardTab>("main");

  const rows = useMemo(() => {
    const base = onlyOccupied
      ? view.roster.filter((row) => !row.isVacant)
      : view.roster;
    return filterByQuery(base, query);
  }, [view.roster, onlyOccupied, query]);

  const selected =
    rows.find((row) => row.key === selectedKey) ||
    view.roster.find((row) => row.key === selectedKey) ||
    null;

  const personAbsents = selected
    ? findPersonAbsents(
        [...view.absentsOpen, ...view.absentsClosed],
        selected,
      )
    : [];

  const historyItems =
    session?.people.filter((person) => {
      if (!selected) return false;
      if (selected.personId && person.personId === selected.personId) return true;
      return (
        person.fullName.trim().toLowerCase() ===
        selected.fullName.trim().toLowerCase()
      );
    }) ?? [];

  if (!hasLiveFile) {
    return (
      <Stack spacing={1.5}>
        <Typography variant="h6">Особовий склад</Typography>
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
    <Stack spacing={1.5} className="ejoos-roster">
      <Box>
        <Typography variant="h6">Особовий склад</Typography>
        <Typography variant="body2" className="ejoos-muted">
          З live ЕЖООС (ШПО + Табель). День: {view.timesheetDayLabel}. Зайнято{" "}
          {view.counts.occupied} / {view.counts.roster}.
          {!hasSnapshot ? " Завантаження файлу…" : ""}
        </Typography>
      </Box>

      <div className="ejoos-roster-toolbar">
        <input
          className="ejoos-search"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder="Пошук: ПІБ, ID, індекс…"
        />
        <label className="ejoos-check-label">
          <input
            type="checkbox"
            checked={onlyOccupied}
            onChange={(event) => setOnlyOccupied(event.target.checked)}
          />
          Лише зайняті
        </label>
      </div>

      <div className="ejoos-changes-layout">
        <div className="ejoos-change-list">
          {rows.map((row) => (
            <button
              key={row.key}
              type="button"
              className={
                selected?.key === row.key
                  ? "ejoos-change-row is-selected"
                  : "ejoos-change-row"
              }
              onClick={() => {
                setSelectedKey(row.key);
                setCardTab("main");
              }}
            >
              <div className="ejoos-change-row-main">
                <strong>{row.fullName || "(вакансія)"}</strong>
                <span className="ejoos-change-meta">
                  {[row.rank, row.positionIndex || row.personId]
                    .filter(Boolean)
                    .join(" · ")}
                </span>
              </div>
              <div className="ejoos-change-will">
                {row.dayCode || "—"}
              </div>
            </button>
          ))}
          {!rows.length ? (
            <Typography variant="body2" className="ejoos-muted">
              Нікого не знайдено.
            </Typography>
          ) : null}
        </div>

        <aside
          className={
            selected ? "ejoos-change-card" : "ejoos-change-card is-empty"
          }
        >
          {!selected ? (
            <Typography variant="body2" className="ejoos-muted">
              Оберіть людину зі списку.
            </Typography>
          ) : (
            <PersonCard
              person={selected}
              cardTab={cardTab}
              setCardTab={setCardTab}
              dayLabel={view.timesheetDayLabel}
              absents={personAbsents}
              historyItems={historyItems.map((item) => ({
                before: item.summaryBefore,
                after: item.summaryAfter,
                will: item.ejoosWillDo.join("; "),
              }))}
            />
          )}
        </aside>
      </div>
    </Stack>
  );
}

function formatAuthorityHint(item?: EjoosFieldAuthority) {
  if (!item) return "—";
  return `статус ← ${FIELD_SOURCE_LABELS[item.primary]}`;
}

function PersonCard({
  person,
  cardTab,
  setCardTab,
  dayLabel,
  absents,
  historyItems,
}: {
  person: EjoosRegisterPerson;
  cardTab: CardTab;
  setCardTab: (tab: CardTab) => void;
  dayLabel: string;
  absents: ReturnType<typeof findPersonAbsents>;
  historyItems: Array<{ before: string; after: string; will: string }>;
}) {
  return (
    <Stack spacing={1.25}>
      <Box>
        <Typography variant="subtitle1">{person.fullName || "(вакансія)"}</Typography>
        <Typography variant="body2" className="ejoos-muted">
          {[person.rank, person.personId].filter(Boolean).join(" · ") || "—"}
        </Typography>
      </Box>

      <div className="ejoos-card-tabs">
        {CARD_TABS.map((item) => (
          <button
            key={item.id}
            type="button"
            className={
              cardTab === item.id
                ? "ejoos-workspace-nav-btn is-active"
                : "ejoos-workspace-nav-btn"
            }
            onClick={() => setCardTab(item.id)}
          >
            {item.label}
          </button>
        ))}
      </div>

      {cardTab === "main" ? (
        <dl className="ejoos-kv">
          <div>
            <dt>ПІБ</dt>
            <dd>{person.fullName || "—"}</dd>
          </div>
          <div>
            <dt>ID</dt>
            <dd>{person.personId || "—"}</dd>
          </div>
          <div>
            <dt>Звання</dt>
            <dd>{person.rank || "—"}</dd>
          </div>
          <div>
            <dt>Статус дня</dt>
            <dd>
              <Chip size="small" label={person.dayCode || "—"} />
            </dd>
          </div>
          <div>
            <dt>Джерела</dt>
            <dd>
              <span className="ejoos-muted">
                {formatAuthorityHint(
                  readOperatorSettings().fieldAuthorities.find(
                    (item) => item.field === "status",
                  ),
                )}
              </span>
            </dd>
          </div>
        </dl>
      ) : null}

      {cardTab === "service" ? (
        <dl className="ejoos-kv">
          <div>
            <dt>Індекс посади</dt>
            <dd>{person.positionIndex || "—"}</dd>
          </div>
          <div>
            <dt>Рядок ШПО/Табель</dt>
            <dd>{person.excelRow}</dd>
          </div>
          <div>
            <dt>Вакансія</dt>
            <dd>{person.isVacant ? "так" : "ні"}</dd>
          </div>
        </dl>
      ) : null}

      {cardTab === "absents" ? (
        absents.length ? (
          <Stack spacing={1}>
            {absents.map((row) => (
              <div key={row.excelRow} className="ejoos-sheet-block">
                <strong>
                  {row.ground || "Відсутність"}
                  {row.actualReturn ? " (закрито)" : " (відкрито)"}
                </strong>
                <span className="ejoos-muted">
                  {row.place || "—"} · з {row.departDate || "?"}
                  {row.actualReturn ? ` → ${row.actualReturn}` : ""}
                </span>
              </div>
            ))}
          </Stack>
        ) : (
          <Typography variant="body2" className="ejoos-muted">
            Немає записів у «Тимч. відсутні».
          </Typography>
        )
      ) : null}

      {cardTab === "timesheet" ? (
        <dl className="ejoos-kv">
          <div>
            <dt>День</dt>
            <dd>{dayLabel}</dd>
          </div>
          <div>
            <dt>Код</dt>
            <dd>{person.dayCode || "—"}</dd>
          </div>
        </dl>
      ) : null}

      {cardTab === "history" ? (
        historyItems.length ? (
          <Stack spacing={1}>
            {historyItems.map((item, index) => (
              <div key={index} className="ejoos-sheet-block">
                <strong>
                  {item.before} → {item.after}
                </strong>
                <span className="ejoos-muted">{item.will}</span>
              </div>
            ))}
          </Stack>
        ) : (
          <Typography variant="body2" className="ejoos-muted">
            Немає змін у поточній сесії 1ПБ для цієї людини.
          </Typography>
        )
      ) : null}
    </Stack>
  );
}
