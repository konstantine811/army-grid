import { useEffect, useMemo, useRef, useState } from "react";
import { TextField } from "@/components/sci/SciPrimitives";
import type { BackendPersonnelRosterLatest } from "../../api";
import { CacheKeys, readDataCache } from "../../data/idbDataCache";
import { loadSharedRosterLatest } from "../../data/sharedAppData";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { loadPersonnelRowsForOverview } from "../overview/overviewPersonnelAssets";
import { overviewNameMatchesQuery } from "../overview/overviewNameSearch";
import {
  buildPersonSummary,
  getPersonDisplayName,
  getPersonFullPositionTitle,
} from "../personnel/personnelUtils";

const rosterRowsFromLatest = (latest: BackendPersonnelRosterLatest | null) => {
  if (!latest?.sheet) return [] as EjournalPreviewRow[];
  return latest.rows.map((row) => ({
    __dbRowId: row.id,
    __rowNumber: row.excelRowNumber,
    ...(row.values && typeof row.values === "object" && !Array.isArray(row.values)
      ? row.values
      : {}),
  })) as EjournalPreviewRow[];
};

const loadPickerPeople = async () => {
  const cached = await readDataCache<BackendPersonnelRosterLatest | null>(
    CacheKeys.rosterLatest,
  );
  const roster = cached ?? (await loadSharedRosterLatest().catch(() => null));
  return loadPersonnelRowsForOverview(rosterRowsFromLatest(roster));
};

export type PersonnelPickerPerson = {
  key: string;
  name: string;
  rank: string;
  position: string;
  row: EjournalPreviewRow;
};

const toPickerPerson = (
  row: EjournalPreviewRow,
  index: number,
): PersonnelPickerPerson | null => {
  const summary = buildPersonSummary(row);
  const name =
    summary.name !== "Особа не вибрана"
      ? summary.name
      : getPersonDisplayName(row);
  if (!name.trim()) return null;
  return {
    key: summary.externalId || String(row.__dbRowId || `row:${index}`),
    name,
    rank: summary.rank || "",
    position: getPersonFullPositionTitle(row),
    row,
  };
};

export function PersonnelNamePicker({
  value,
  onPick,
  disabled,
  placeholder = "Почніть вводити ПІБ з Особового складу",
}: {
  value: string;
  onPick: (row: EjournalPreviewRow) => void;
  disabled?: boolean;
  placeholder?: string;
}) {
  const rootRef = useRef<HTMLDivElement | null>(null);
  const [query, setQuery] = useState(value);
  const [open, setOpen] = useState(false);
  const [people, setPeople] = useState<PersonnelPickerPerson[]>([]);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    setQuery(value);
  }, [value]);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    void loadPickerPeople()
      .then((rows) => {
        if (cancelled) return;
        setPeople(
          rows
            .map(toPickerPerson)
            .filter((item): item is PersonnelPickerPerson => Boolean(item)),
        );
      })
      .catch(() => {
        if (!cancelled) setPeople([]);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    const onPointerDown = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) {
        setOpen(false);
        setQuery(value);
      }
    };
    window.addEventListener("pointerdown", onPointerDown);
    return () => window.removeEventListener("pointerdown", onPointerDown);
  }, [value]);

  const matches = useMemo(() => {
    const text = query.trim();
    if (text.length < 2) return [] as PersonnelPickerPerson[];
    return people
      .filter((person) => overviewNameMatchesQuery(person.name, text))
      .slice(0, 10);
  }, [people, query]);

  return (
    <div className="document-person-picker" ref={rootRef}>
      <TextField
        size="small"
        fullWidth
        disabled={disabled}
        value={query}
        placeholder={placeholder}
        autoComplete="off"
        onFocus={() => setOpen(true)}
        onChange={(event) => {
          setQuery(event.target.value);
          setOpen(true);
        }}
      />
      {open && !disabled ? (
        <div className="document-person-picker-list" role="listbox">
          {loading && query.trim().length >= 2 ? (
            <div className="document-person-picker-empty">Шукаю в Особовому складі…</div>
          ) : null}
          {!loading && query.trim().length >= 2 && !matches.length ? (
            <div className="document-person-picker-empty">
              Немає в Особовому складі — поставте «Ввести розслідувача вручну»
            </div>
          ) : null}
          {matches.map((person) => (
            <button
              key={person.key}
              type="button"
              role="option"
              className="document-person-picker-item"
              onClick={() => {
                onPick(person.row);
                setQuery(person.name);
                setOpen(false);
              }}
            >
              <strong>{person.name}</strong>
              <small>
                {[person.rank, person.position].filter(Boolean).join(" · ") ||
                  "Без звання / посади"}
              </small>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}
