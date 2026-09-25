import {
  forwardRef,
  memo,
  useDeferredValue,
  useImperativeHandle,
  useMemo,
  useState,
} from "react";
import { SearchOutlinedIcon } from "@/components/sci/icons";
import { MagneticTapePreloader } from "@/components/sci/MagneticTapePreloader";
import { PersonnelVirtualList } from "./PersonnelVirtualList";
import type { PersonnelListRecord } from "./personnelListIndex";
import {
  personnelSearchEmptyHint,
  personnelSearchMatchesQuery,
} from "./personnelSearch";
import { normalizePersonnelSearchText } from "./personnelUtils";

export type PersonnelListPanelHandle = {
  setSearch: (value: string) => void;
};

type StaffFilter = "all" | "in" | "archive";

export const PersonnelListPanel = memo(
  forwardRef<
    PersonnelListPanelHandle,
    {
      initialSearch?: string;
      personnelRows: PersonnelListRecord[];
      phonesByExternalId: Record<string, string[]>;
      staffFilter: StaffFilter;
      staffCounts: { all: number; in: number; archive: number };
      questionnaireByExternalId: Record<string, true>;
      questionnairePresenceStatus: "loading" | "ready" | "error";
      isLoading: boolean;
      selectedRowId: string;
      photoByExternalId: Record<string, string>;
      selectedPhotoFullUrl?: string;
      keyboardEnabled: boolean;
      onStaffFilterChange: (value: StaffFilter) => void;
      onNeedPhotos?: (externalIds: string[]) => void;
      onPhotoLoadError?: (externalId: string) => void;
      onSelect: (rowId: string) => void;
    }
  >(function PersonnelListPanel(
    {
      initialSearch = "",
      personnelRows,
      phonesByExternalId,
      staffFilter,
      staffCounts,
      questionnaireByExternalId,
      questionnairePresenceStatus,
      isLoading,
      selectedRowId,
      photoByExternalId,
      selectedPhotoFullUrl = "",
      keyboardEnabled,
      onStaffFilterChange,
      onNeedPhotos,
      onPhotoLoadError,
      onSelect,
    },
    ref,
  ) {
    const [query, setQuery] = useState(initialSearch);
    const deferredQuery = useDeferredValue(query);

    useImperativeHandle(ref, () => ({
      setSearch: (value: string) => setQuery(value),
    }));

    const filteredPersonnel = useMemo(() => {
      const normalizedQuery = normalizePersonnelSearchText(deferredQuery);

      return personnelRows.filter((record) => {
        if (staffFilter === "in" && !record.inStaff) return false;
        if (staffFilter === "archive" && !record.inArchive) return false;
        if (!normalizedQuery) return true;

        const phones = phonesByExternalId[record.summary.externalId];
        const searchableText =
          phones?.length
            ? normalizePersonnelSearchText(
                [record.searchBase, ...phones].filter(Boolean).join(" "),
              )
            : record.searchBase;

        return personnelSearchMatchesQuery(searchableText, normalizedQuery, {
          primaryText: normalizePersonnelSearchText(record.summary.name),
          callSignText: normalizePersonnelSearchText(record.summary.callSign),
        });
      });
    }, [personnelRows, phonesByExternalId, deferredQuery, staffFilter]);

    const questionnaireCounts = useMemo(() => {
      let withQuestionnaire = 0;
      for (const record of filteredPersonnel) {
        const externalId = record.summary.externalId;
        if (externalId && questionnaireByExternalId[externalId]) {
          withQuestionnaire += 1;
        }
      }
      return {
        withQuestionnaire,
        withoutQuestionnaire: filteredPersonnel.length - withQuestionnaire,
      };
    }, [filteredPersonnel, questionnaireByExternalId]);

    return (
      <aside className="analytics-panel personnel-list-panel">
        <div className="panel-heading personnel-list-heading">
          <span>Військовослужбовці · {filteredPersonnel.length}</span>
          <span className="personnel-questionnaire-summary">
            {questionnairePresenceStatus === "ready"
              ? `З анкетами · ${questionnaireCounts.withQuestionnaire}  /  Без анкет · ${questionnaireCounts.withoutQuestionnaire}`
              : questionnairePresenceStatus === "error"
                ? "Анкети · не вдалося завантажити"
                : "Анкети · завантаження…"}
          </span>
        </div>
        <label className="personnel-search">
          <SearchOutlinedIcon fontSize="small" />
          <input
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="ПІБ… А* (на «А»), к?с (Word-маски), Кос"
          />
        </label>
        <div
          aria-label="Фільтр за штаткою"
          className="personnel-questionnaire-filter"
          role="group"
        >
          {(
            [
              ["all", "Усі", staffCounts.all],
              ["in", "У штаті", staffCounts.in],
              ["archive", "Архів", staffCounts.archive],
            ] as const
          ).map(([value, label, count]) => (
            <button
              aria-pressed={staffFilter === value}
              className={staffFilter === value ? "is-active" : undefined}
              key={value}
              title={
                value === "in"
                  ? "Лише особи з актуального «Загального списку» Штатки"
                  : value === "archive"
                    ? "Лише особи, додані з аркуша «Архів»"
                    : "Усі особи"
              }
              onClick={() => onStaffFilterChange(value)}
              type="button"
            >
              {count < 0 ? label : `${label} · ${count}`}
            </button>
          ))}
        </div>
        {isLoading && personnelRows.length === 0 ? (
          <div className="personnel-list-preloader">
            <MagneticTapePreloader
              status="ГОТУЮ СПИСОК У ШТАТІ"
              hint="Спочатку завантажую Штатку, тому проміжний список ООС не показується."
            />
          </div>
        ) : (
          <PersonnelVirtualList
            items={filteredPersonnel}
            selectedRowId={selectedRowId}
            photoByExternalId={photoByExternalId}
            selectedPhotoFullUrl={selectedPhotoFullUrl}
            onNeedPhotos={onNeedPhotos}
            onSelect={onSelect}
            onPhotoLoadError={onPhotoLoadError}
            keyboardEnabled={keyboardEnabled}
          />
        )}
        {filteredPersonnel.length === 0 && query.trim() ? (
          <p className="personnel-empty-hint">
            {personnelSearchEmptyHint(query) ||
              (staffFilter === "in"
                ? "Нічого у «У штаті». Спробуйте «Усі» або «Імпорт Штатки в БД»."
                : staffFilter === "archive"
                  ? "Нічого в «Архіві». Імпортуйте Штатку з аркушем «Архів»."
                  : "Нічого не знайдено. Спробуйте «Імпорт Штатки в БД».")}
          </p>
        ) : null}
      </aside>
    );
  }),
);
