import { useEffect, useMemo, useRef, useState } from "react";
import { useVirtualizer } from "@tanstack/react-virtual";
import {
  Alert,
  Box,
  Button,
  Chip,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import type { BackendEjournalLiveVersion } from "../../api";
import {
  collectAppliedHistoryEntries,
  readAppliedHistoryFromWorkbook,
  type StoredApplyRecord,
} from "./ejoosAppliedHistory";
import {
  PersonChangeCard,
  PersonChangeRow,
} from "./EjoosPersonChangeCard";
import { ejoosDownloadFileName } from "./ejoosWorkbookKind";
import { useEjoosWorkspace } from "./ejoosWorkspaceState";

/** Версії після доповнення з анкет через xlsx-populate часто биті в Excel. */
const looksLikeAnketaFillVersion = (version: BackendEjournalLiveVersion) => {
  const notes = String(version.notes ?? "");
  const protocol = version.changeProtocol as { kind?: string } | null;
  const kind = String(protocol?.kind ?? "");
  return (
    /анкет/i.test(notes) ||
    /anketa/i.test(kind) ||
    /ООС ← анкет|Виключені ← анкет/i.test(notes)
  );
};

export function EjoosHistoryPanel() {
  const {
    live,
    ejoosSnapshot,
    ensureEjoosLoaded,
    downloadVersion,
    downloadVersionProtocol,
    rollback,
    restoreStylesFromHistory,
    isLoading,
    setTab,
  } = useEjoosWorkspace();

  const versions = [...(live?.versions ?? [])].sort(
    (a, b) => b.version - a.version,
  );
  const current = live?.current;
  const currentId = current?.id;
  const currentLooksBad =
    Boolean(current) && looksLikeAnketaFillVersion(current!);

  const safeRollbackTarget = versions.find(
    (version) =>
      version.id !== currentId &&
      (!current || version.version < current.version) &&
      !looksLikeAnketaFillVersion(version),
  );

  const currentVersionNumber =
    current?.version ||
    Math.max(0, ...versions.map((version) => version.version), 0);

  const [fileApplies, setFileApplies] = useState<StoredApplyRecord[] | null>(
    null,
  );

  useEffect(() => {
    if (ejoosSnapshot || !live?.current) return;
    void ensureEjoosLoaded();
  }, [ejoosSnapshot, live?.current, ensureEjoosLoaded]);

  useEffect(() => {
    const file = ejoosSnapshot?.file;
    if (!file) {
      setFileApplies(null);
      return;
    }
    let cancelled = false;
    void readAppliedHistoryFromWorkbook(file)
      .then((applies) => {
        if (!cancelled) setFileApplies(applies);
      })
      .catch(() => {
        if (!cancelled) setFileApplies([]);
      });
    return () => {
      cancelled = true;
    };
  }, [ejoosSnapshot?.file, ejoosSnapshot?.fileName]);

  const entries = useMemo(
    () =>
      collectAppliedHistoryEntries(
        live?.versions ?? [],
        currentVersionNumber,
        fileApplies,
      ),
    [live?.versions, currentVersionNumber, fileApplies],
  );

  const [query, setQuery] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return entries.filter((entry) => {
      if (!q) return true;
      const person = entry.person;
      return (
        person.fullName.toLowerCase().includes(q) ||
        person.personId.toLowerCase().includes(q) ||
        person.positionIndex.toLowerCase().includes(q) ||
        entry.pbFileName.toLowerCase().includes(q)
      );
    });
  }, [entries, query]);

  const selected =
    filtered.find((entry) => entry.id === selectedId) ??
    entries.find((entry) => entry.id === selectedId) ??
    null;

  const historyViewportRef = useRef<HTMLDivElement>(null);
  const historyVirtualizer = useVirtualizer({
    count: versions.length,
    getScrollElement: () => historyViewportRef.current,
    estimateSize: () => 112,
    overscan: 6,
  });

  return (
    <Stack spacing={1.5} className="ejoos-history ejoos-changes">
      <Stack
        direction={{ xs: "column", md: "row" }}
        justifyContent="space-between"
        alignItems={{ xs: "stretch", md: "center" }}
        spacing={1}
      >
        <Box>
          <Typography variant="h6">Застосовані зміни</Typography>
          <Typography variant="body2" className="ejoos-muted">
            Журнал цього файлу ЕЖООС: що вже підтверджено і записано в нього з
            останнього імпорту. Не історія всіх версій у БД.
            {entries.length ? ` · ${entries.length} записів` : ""}
            {current ? ` · поточний файл v${current.version}` : ""}
          </Typography>
        </Box>
        <Button size="small" variant="outlined" onClick={() => setTab("changes")}>
          До операцій
        </Button>
      </Stack>

      {currentLooksBad ? (
        <Alert severity="warning" variant="outlined">
          Поточна <strong>v{current?.version}</strong> ймовірно пошкоджена для
          Excel (доповнення з анкет). Стилі можна повернути з попередньої
          робочої версії — дані мерджу залишаться.
          {safeRollbackTarget ? (
            <Stack
              direction="row"
              spacing={1}
              sx={{ mt: 1.5 }}
              style={{ flexWrap: "wrap" }}
            >
              <Button
                size="small"
                variant="contained"
                disabled={isLoading}
                onClick={() =>
                  void restoreStylesFromHistory(safeRollbackTarget.id)
                }
                sx={{ color: "#1a1a14" }}
              >
                Повернути стилі з v{safeRollbackTarget.version}
              </Button>
              <Button
                size="small"
                variant="outlined"
                disabled={isLoading}
                onClick={() => void rollback(safeRollbackTarget.id)}
              >
                Відкотити на v{safeRollbackTarget.version}
              </Button>
            </Stack>
          ) : null}
        </Alert>
      ) : null}

      <Stack
        direction="row"
        spacing={1}
        style={{ flexWrap: "wrap", alignItems: "center" }}
      >
        <input
          className="ejoos-search"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder="Пошук ПІБ / ID / індекс"
        />
      </Stack>

      {entries.length === 0 ? (
        <Box className="ejoos-change-card is-empty">
          <Typography variant="body2" className="ejoos-muted">
            Ще немає застосованих змін у цьому файлі. Після «Підтвердити і
            застосувати» на вкладці «Операції» сюди потраплять записи, які
            реально записались у поточний ЕЖООС.
          </Typography>
        </Box>
      ) : (
        <div className="ejoos-changes-layout">
          <div className="ejoos-change-list">
            {filtered.length === 0 ? (
              <Typography variant="body2" className="ejoos-muted" sx={{ p: 2 }}>
                Немає записів за пошуком
              </Typography>
            ) : (
              filtered.map((entry) => (
                <PersonChangeRow
                  key={entry.id}
                  person={entry.person}
                  selected={entry.id === selectedId}
                  onSelect={() => setSelectedId(entry.id)}
                  mode="history"
                  historyMeta={{
                    version: entry.version,
                    appliedAt: entry.appliedAt,
                  }}
                />
              ))
            )}
          </div>
          <div className="ejoos-change-detail">
            {selected ? (
              <PersonChangeCard
                person={selected.person}
                timesheetDay={selected.timesheetDay}
                mode="history"
                historyMeta={{
                  version: selected.version,
                  appliedAt: selected.appliedAt,
                  pbFileName: selected.pbFileName,
                  timesheetDayLabel: selected.timesheetDayLabel,
                }}
                onClose={() => setSelectedId(null)}
              />
            ) : (
              <Box className="ejoos-change-card is-empty">
                <Typography variant="body2" className="ejoos-muted">
                  Оберіть людину, щоб побачити що вже записано в ЕЖООС: було /
                  стало, події 1ПБ, табель і карта аркушів.
                </Typography>
              </Box>
            )}
          </div>
        </div>
      )}

      <details className="ejoos-legacy-details">
        <summary>Версії файлу ЕЖООС · скачати / відкотити</summary>
        <Typography variant="body2" className="ejoos-muted" sx={{ mb: 1 }}>
          Версії не перезаписуються. Відкат створює нову копію обраної версії.
        </Typography>
        {versions.length === 0 ? (
          <Typography variant="body2" className="ejoos-muted">
            Поки немає збережених версій.
          </Typography>
        ) : (
          <div ref={historyViewportRef} className="ejoos-history-viewport">
            <div
              className="ejoos-history-virtual-list"
              style={{ height: historyVirtualizer.getTotalSize() }}
            >
              {historyVirtualizer.getVirtualItems().map((virtualRow) => {
                const version = versions[virtualRow.index];
                const isCurrent = version.id === currentId;
                const isBad = looksLikeAnketaFillVersion(version);
                return (
                  <div
                    key={version.id}
                    ref={historyVirtualizer.measureElement}
                    data-index={virtualRow.index}
                    className="ejoos-history-virtual-row"
                    style={{ transform: `translateY(${virtualRow.start}px)` }}
                  >
                    <div className="ejoos-history-row">
                      <div className="ejoos-history-meta">
                        <Stack direction="row" spacing={0.5} alignItems="center">
                          <Chip
                            size="small"
                            label={`v${version.version}`}
                            color={isCurrent ? "primary" : "default"}
                          />
                          {isBad ? (
                            <Chip
                              size="small"
                              color="warning"
                              label="ризик Excel"
                            />
                          ) : null}
                        </Stack>
                        <Typography variant="body2">
                          {new Date(version.createdAt).toLocaleString("uk-UA")}
                          {version.sourceFileName
                            ? ` · ${version.sourceFileName}`
                            : ""}
                        </Typography>
                        {version.sourcePbFileName ? (
                          <Typography variant="caption" className="ejoos-muted">
                            1ПБ: {version.sourcePbFileName}
                          </Typography>
                        ) : null}
                        {version.notes ? (
                          <Typography variant="caption" className="ejoos-muted">
                            {version.notes}
                          </Typography>
                        ) : null}
                      </div>
                      <Stack
                        direction="row"
                        spacing={0.5}
                        style={{ flexWrap: "wrap" }}
                      >
                        <Button
                          size="small"
                          disabled={isLoading}
                          onClick={() =>
                            void downloadVersion(
                              version.id,
                              ejoosDownloadFileName(
                                version.version,
                                version.asOfDate,
                              ),
                            )
                          }
                        >
                          Файл
                        </Button>
                        <Button
                          size="small"
                          disabled={isLoading}
                          onClick={() => downloadVersionProtocol(version)}
                        >
                          Протокол
                        </Button>
                        {!isCurrent ? (
                          <Button
                            size="small"
                            disabled={isLoading}
                            onClick={() => void rollback(version.id)}
                          >
                            Відкотити
                          </Button>
                        ) : null}
                      </Stack>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </details>
    </Stack>
  );
}
