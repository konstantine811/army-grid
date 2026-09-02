import { memo, useEffect, useMemo, useRef, useState } from "react";
import {
  Alert,
  Box,
  Button,
  Chip,
  Divider,
  LinearProgress,
  MenuItem,
  Select,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import {
  api,
  type BackendEjournalLiveState,
  type BackendEjournalLiveVersion,
} from "../../api";
import { formatApiDateTime } from "../../shared/format";
import {
  type ExcelWorkbookSnapshot,
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../../excelRoundTrip";
import { readEjoosWorkbookSnapshot } from "./ejoosTimesheetPersonRows";
import {
  applyConfirmedEjoosOps,
  base64ToFile,
  blobToBase64,
  downloadBlobFile,
  downloadTextFile,
} from "./ejoosSyncApply";
import {
  buildConfirmSummary,
  buildProtocolText,
  collectProcessedMovementKeys,
  planBlocksWorkbookApply,
  SOURCE_DATE_UNKNOWN_MESSAGE,
  workbookApplyBlockMessage,
  type EjoosSyncOp,
  type EjoosSyncPlan,
} from "./ejoosSyncPlan";
import {
  timesheetOpBlocksApply,
  timesheetOpNeedsManualCode,
  excludeTransferOpBlocksApply,
  ambiguousTransferOpBlocksApply,
  contradictoryStatusOpsBlockApply,
} from "./ejoosOpRequirements";
import {
  assertEjoosWorkbook,
  assertPbWorkbook,
  detectWorkbookKindFromFile,
  ejoosDownloadFileName,
} from "./ejoosWorkbookKind";
import { readOperatorSettings } from "./ejoosStatusMap";
import { runHeavyJob } from "../../workers/runHeavyJob";

type FilterClass = "ALL" | "ready" | "needs_input" | "conflict";

const classLabel: Record<Exclude<FilterClass, "ALL">, string> = {
  ready: "Готово до застосування",
  needs_input: "Потрібна твоя дія",
  conflict: "Конфлікт / перевірка",
};

export function EjoosLiveSyncPanel() {
  const [live, setLive] = useState<BackendEjournalLiveState | null>(null);
  const [ejoosSnapshot, setEjoosSnapshot] =
    useState<ExcelWorkbookSnapshot | null>(null);
  const [pbSnapshot, setPbSnapshot] = useState<ExcelWorkbookSnapshot | null>(
    null,
  );
  const [plan, setPlan] = useState<EjoosSyncPlan | null>(null);
  const [sourceAsOf, setSourceAsOf] = useState("");
  const [checkedIds, setCheckedIds] = useState<Set<string>>(new Set());
  const [manualCodes, setManualCodes] = useState<Record<string, string>>({});
  const [manualFields, setManualFields] = useState<
    Record<string, { destination?: string; place?: string; departDate?: string }>
  >({});
  const [filter, setFilter] = useState<FilterClass>("ALL");
  const [sheetFilter, setSheetFilter] = useState("ALL");
  const [nameQuery, setNameQuery] = useState("");
  const [page, setPage] = useState(0);
  const PAGE_SIZE = 40;
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const planGenRef = useRef(0);
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [lastProtocol, setLastProtocol] = useState("");
  const [lastResultFiles, setLastResultFiles] = useState<{
    ejoosName: string;
    protocolText: string;
    version: number;
  } | null>(null);
  const buildPlanOffThread = (
    ejoos: ExcelWorkbookSnapshot,
    pb: ExcelWorkbookSnapshot,
    sourceAsOfDate?: string,
  ) =>
    runHeavyJob({
      type: "ejoosSyncPlan",
      ejoos,
      pb,
      processedMovementKeys: [
        ...collectProcessedMovementKeys(live?.versions),
      ],
      sourceAsOfDate: sourceAsOfDate || undefined,
      statusRules: readOperatorSettings().statusRules,
    });

  const refreshLive = async () => {
    const state = await api.getEjournalLive("1ПБ");
    setLive(state);
    return state;
  };

  useEffect(() => {
    void refreshLive().catch((err) => {
      setError(err instanceof Error ? err.message : "Не вдалося завантажити ЕЖООС з БД");
    });
  }, []);

  const loadCurrentEjoosFromDb = async (version?: BackendEjournalLiveVersion | null) => {
    const meta = version ?? live?.current;
    if (!meta) return null;
    const full = await api.getEjournalLiveFile(meta.id, "1ПБ");
    if (!full.fileBase64) throw new Error("У версії немає файлу");
    const file = base64ToFile(
      full.fileBase64,
      full.sourceFileName || `ЕЖООС_v${full.version}.xlsx`,
    );
    const snapshot = await readEjoosWorkbookSnapshot(file);
    setEjoosSnapshot(snapshot);
    return snapshot;
  };

  const seedEjoos = async (file: File | undefined) => {
    if (!file) return;
    setIsLoading(true);
    setError("");
    try {
      const snapshot = await readEjoosWorkbookSnapshot(file);
      assertEjoosWorkbook(snapshot);
      const fileBase64 = await blobToBase64(file);
      const created = await api.seedEjournalLive({
        fileBase64,
        sourceFileName: file.name,
        unitLabel: "1ПБ",
        notes: "Початкове завантаження канонічного ЕЖООС",
      });
      await refreshLive();
      setEjoosSnapshot(snapshot);
      setMessage(`Збережено канонічний ЕЖООС v${created.version} · sha ${created.sha256.slice(0, 8)}…`);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося зберегти ЕЖООС");
    } finally {
      setIsLoading(false);
    }
  };

  const openPb = async (file: File | undefined) => {
    if (!file) return;
    setIsLoading(true);
    setError("");
    try {
      let ejoos = ejoosSnapshot;
      if (!ejoos) {
        ejoos = await loadCurrentEjoosFromDb();
      }
      if (!ejoos) {
        throw new Error("Спочатку завантажте канонічний ЕЖООС у БД");
      }
      assertEjoosWorkbook(ejoos);
      const pb = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
      assertPbWorkbook(pb);
      setPbSnapshot(pb);
      const nextPlan = await buildPlanOffThread(ejoos, pb, sourceAsOf);
      setPlan(nextPlan);
      setPage(0);
      setCheckedIds(
        new Set(
          nextPlan.ops.filter((op) => op.checkedDefault).map((op) => op.id),
        ),
      );
      setManualCodes({});
      setManualFields({});
      setSheetFilter("ALL");
      setNameQuery("");
      setFilter("ALL");
      setMessage(
        `План змін з ${file.name}: готово ${nextPlan.summary.ready}, потрібна дія ${nextPlan.summary.needsInput}, конфліктів ${nextPlan.summary.conflict}. Нічого ще не записано.` +
          (nextPlan.limitsNote ? ` ${nextPlan.limitsNote}` : ""),
      );
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося розібрати 1ПБ");
    } finally {
      setIsLoading(false);
    }
  };

  const sheetOptions = useMemo(() => {
    if (!plan) return [] as string[];
    return [...new Set(plan.ops.map((op) => op.sheet))].sort();
  }, [plan]);

  const visibleOps = useMemo(() => {
    if (!plan) return [];
    const q = nameQuery.trim().toLowerCase();
    return plan.ops.filter((op) => {
      if (filter !== "ALL" && op.class !== filter) return false;
      if (sheetFilter !== "ALL" && op.sheet !== sheetFilter) return false;
      if (!q) return true;
      return (
        op.fullName.toLowerCase().includes(q) ||
        op.personId.toLowerCase().includes(q) ||
        op.positionIndex.toLowerCase().includes(q)
      );
    });
  }, [filter, nameQuery, plan, sheetFilter]);

  const pageCount = Math.max(1, Math.ceil(visibleOps.length / PAGE_SIZE));
  const pagedOps = useMemo(() => {
    const safePage = Math.min(page, pageCount - 1);
    const start = safePage * PAGE_SIZE;
    return visibleOps.slice(start, start + PAGE_SIZE);
  }, [PAGE_SIZE, page, pageCount, visibleOps]);

  useEffect(() => {
    setPage(0);
  }, [filter, sheetFilter, nameQuery]);

  const selectedOps = useMemo(() => {
    if (!plan) return [];
    return plan.ops
      .filter((op) => checkedIds.has(op.id))
      .map((op) => {
        const manual = manualCodes[op.id];
        const fields = manualFields[op.id];
        let next = op;
        if (manual) {
          next = {
            ...next,
            after: manual,
            payload: { ...next.payload, timesheetCode: manual },
            class: "ready",
          };
        }
        if (fields?.destination) {
          next = {
            ...next,
            after: `виключити: ${next.payload.type || ""} → ${fields.destination}`,
            payload: { ...next.payload, destination: fields.destination },
            class: "ready",
          };
        }
        if (fields?.place || fields?.departDate) {
          next = {
            ...next,
            after: `${next.payload.absenceType || "?"} → ${fields.place || next.payload.place || "?"} з ${fields.departDate || next.payload.departDate || "?"}`,
            payload: {
              ...next.payload,
              place: fields.place || next.payload.place || "",
              departDate: fields.departDate || next.payload.departDate || "",
            },
            class:
              (fields.place || next.payload.place) &&
              (fields.departDate || next.payload.departDate)
                ? "ready"
                : next.class,
          };
        }
        return next;
      });
  }, [checkedIds, manualCodes, manualFields, plan]);

  const confirmSummary = useMemo(
    () => buildConfirmSummary(selectedOps),
    [selectedOps],
  );

  const blockedSelected = selectedOps.filter((op) => {
    if (op.class === "conflict") return true;
    if (timesheetOpBlocksApply(op)) return true;
    if (excludeTransferOpBlocksApply(op)) return true;
    if (ambiguousTransferOpBlocksApply(op)) return true;
    if (op.kind === "absent_upsert") {
      return !(op.payload.absenceType && op.payload.departDate);
    }
    if (op.class === "needs_input") return true;
    return false;
  });
  if (contradictoryStatusOpsBlockApply(selectedOps)) {
    blockedSelected.push(
      ...selectedOps.filter((op) => op.kind === "position_change"),
    );
  }

  const toggleOp = (id: string) => {
    setCheckedIds((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const selectVisibleReady = () => {
    setCheckedIds((current) => {
      const next = new Set(current);
      visibleOps.forEach((op) => {
        if (op.class === "ready") next.add(op.id);
      });
      return next;
    });
  };

  const clearVisibleChecks = () => {
    setCheckedIds((current) => {
      const next = new Set(current);
      visibleOps.forEach((op) => next.delete(op.id));
      return next;
    });
  };

  const downloadDraftProtocol = () => {
    if (!plan) return;
    const text = buildProtocolText(plan, selectedOps, {
      actor: "draft",
      at: new Date().toLocaleString("uk-UA"),
    });
    downloadTextFile(
      `протокол_чернетка_${plan.timesheetDayLabel.replaceAll(".", "-")}.txt`,
      text,
    );
  };

  const applySelected = async () => {
    if (!plan || !ejoosSnapshot || !live?.current) return;
    if (planBlocksWorkbookApply(plan)) {
      setError(workbookApplyBlockMessage(plan));
      return;
    }
    if (!selectedOps.length) {
      setError("Не вибрано жодної зміни");
      return;
    }
    if (blockedSelected.length) {
      setError(
        `Серед вибраних є ${blockedSelected.length} змін, які потребують дозаповнення або є конфліктом`,
      );
      return;
    }
    setConfirmOpen(false);
    setIsLoading(true);
    setError("");
    try {
      assertEjoosWorkbook(ejoosSnapshot);
      const result = await applyConfirmedEjoosOps({
        ejoos: ejoosSnapshot,
        plan,
        ops: selectedOps,
      });
      const resultBlob = result.blob;
      const protocolWithVersion = buildProtocolText(plan, selectedOps, {
        actor: "operator",
        at: new Date().toLocaleString("uk-UA"),
        version: (live.current.version || 0) + 1,
      });
      const changeProtocol = {
        ...result.changeProtocol,
        protocolText: protocolWithVersion,
      };
      const fileBase64 = await blobToBase64(resultBlob);
      const saved = await api.applyEjournalLive({
        baseVersionId: live.current.id,
        fileBase64,
        sourceFileName: result.fileName,
        asOfDate: plan.timesheetDayLabel,
        unitLabel: "1ПБ",
        sourcePbFileName: pbSnapshot?.fileName,
        changeProtocol,
        notes: `Застосовано ${selectedOps.length} змін з ${plan.pbName}`,
      });
      downloadBlobFile(result.fileName, resultBlob);
      downloadTextFile(
        `протокол_ЕЖООС_v${saved.version}.txt`,
        protocolWithVersion,
      );
      setLastProtocol(protocolWithVersion);
      setLastResultFiles({
        ejoosName: result.fileName,
        protocolText: protocolWithVersion,
        version: saved.version,
      });
      const state = await refreshLive();
      if (state.current) {
        await loadCurrentEjoosFromDb(state.current);
      }
      setPlan(null);
      setPbSnapshot(null);
      setCheckedIds(new Set());
      setMessage(
        `Записано ЕЖООС v${saved.version}. Застосовано ${selectedOps.length} змін. Стара версія збережена.`,
      );
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Не вдалося застосувати зміни",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const downloadVersionFile = async (
    versionId: string,
    _fileNameHint?: string,
  ) => {
    setIsLoading(true);
    try {
      const full = await api.getEjournalLiveFile(versionId);
      if (!full.fileBase64) throw new Error("Немає файлу");
      const downloadName = ejoosDownloadFileName(
        full.version,
        full.asOfDate,
        full.sourceFileName,
      );
      const file = base64ToFile(full.fileBase64, downloadName);
      if ((await detectWorkbookKindFromFile(file)) === "pb_1pb") {
        throw new Error(
          "У цій версії збережено 1ПБ (sh/Рух/archive), а не ЕЖООС. Імпортуйте канонічний ЕЖООС і застосуйте зміни знову.",
        );
      }
      downloadBlobFile(downloadName, file);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося скачати");
    } finally {
      setIsLoading(false);
    }
  };

  const downloadCurrent = async () => {
    if (!live?.current) return;
    await downloadVersionFile(live.current.id);
  };

  const downloadVersionProtocol = (version: BackendEjournalLiveVersion) => {
    const protocol = version.changeProtocol as { protocolText?: string } | null;
    const text =
      (protocol && typeof protocol.protocolText === "string"
        ? protocol.protocolText
        : "") ||
      `Протокол v${version.version}\nНемає збереженого тексту протоколу.`;
    downloadTextFile(`протокол_ЕЖООС_v${version.version}.txt`, text);
  };

  const rollbackTo = async (versionId: string) => {
    if (!window.confirm("Зробити цю версію поточною? Буде створено нову версію-копію.")) {
      return;
    }
    setIsLoading(true);
    try {
      const targetMeta = live?.versions?.find((version) => version.id === versionId);
      if (!targetMeta) throw new Error("Версію для відкату не знайдено");

      const targetFile = await api.getEjournalLiveFile(versionId, "1ПБ");
      if (!targetFile.fileBase64) {
        throw new Error("Файл цільової версії не повернувся з БД");
      }
      const targetBlob = base64ToFile(
        targetFile.fileBase64,
        targetFile.sourceFileName || `ЕЖООС_v${targetMeta.version}.xlsx`,
      );
      const fileBase64 = await blobToBase64(targetBlob);

      const saved = await api.rollbackEjournalLive({
        targetVersionId: versionId,
        unitLabel: "1ПБ",
        fileBase64,
      });
      await refreshLive();
      await loadCurrentEjoosFromDb(saved);
      setMessage(`Відкат: поточна тепер v${saved.version} (копія обраної)`);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося відкотити");
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <Box className="panel-card" sx={{ mb: 2, p: 2 }}>
      <Stack spacing={1.5}>
        <Box>
          <Typography variant="h6">Ведення ЕЖООС з 1ПБ</Typography>
          <Typography variant="body2" color="text.secondary">
            Система лише пропонує зміни. Запис у файл і БД — тільки після твого
            перегляду й підтвердження. Кожна версія зберігається.
          </Typography>
        </Box>

        {isLoading ? <LinearProgress /> : null}
        {message ? <Alert severity="success">{message}</Alert> : null}
        {error ? <Alert severity="error">{error}</Alert> : null}

        <Stack direction={{ xs: "column", md: "row" }} spacing={1} alignItems="flex-start">
          <Box sx={{ flex: 1 }}>
            <Typography variant="subtitle2">1. Канонічний ЕЖООС</Typography>
            {live?.current ? (
              <Stack spacing={0.5}>
                <Typography variant="body2">
                  Поточна <strong>v{live.current.version}</strong> ·{" "}
                  {live.current.sourceFileName || "без імені"} ·{" "}
                  {Math.round(live.current.byteSize / 1024)} КБ · sha{" "}
                  {live.current.sha256.slice(0, 10)}…
                </Typography>
                <Stack direction="row" spacing={1} flexWrap="wrap">
                  <Button
                    size="small"
                    variant="outlined"
                    startIcon={<FileDownloadOutlinedIcon />}
                    onClick={() => void downloadCurrent()}
                  >
                    Завантажити поточний
                  </Button>
                  <Button
                    size="small"
                    variant="outlined"
                    onClick={() => void loadCurrentEjoosFromDb()}
                  >
                    Відкрити з БД для аналізу
                  </Button>
                </Stack>
                {live.versions.length > 1 ? (
                  <Box sx={{ mt: 1 }}>
                    <Typography variant="caption" color="text.secondary">
                      Історія версій
                    </Typography>
                    <Stack spacing={0.5} sx={{ mt: 0.5 }}>
                      {live.versions.slice(0, 8).map((version) => (
                        <Stack
                          key={version.id}
                          direction="row"
                          spacing={1}
                          alignItems="center"
                        >
                          <Chip
                            size="small"
                            label={`v${version.version}`}
                            color={
                              version.id === live.current?.id
                                ? "primary"
                                : "default"
                            }
                          />
                          <Typography variant="caption">
                            {formatApiDateTime(version.createdAt)}
                            {version.sourcePbFileName
                              ? ` · ${version.sourcePbFileName}`
                              : ""}
                          </Typography>
                          {version.id !== live.current?.id ? (
                            <Button
                              size="small"
                              onClick={() => void rollbackTo(version.id)}
                            >
                              Відкотити
                            </Button>
                          ) : null}
                          <Button
                            size="small"
                            onClick={() =>
                              void downloadVersionFile(
                                version.id,
                                version.sourceFileName || undefined,
                              )
                            }
                          >
                            Файл
                          </Button>
                          <Button
                            size="small"
                            onClick={() => downloadVersionProtocol(version)}
                          >
                            Протокол
                          </Button>
                        </Stack>
                      ))}
                    </Stack>
                  </Box>
                ) : null}
              </Stack>
            ) : (
              <Button
                component="label"
                variant="contained"
                startIcon={<CloudUploadOutlinedIcon />}
              >
                Завантажити ЕЖООС у БД (один раз)
                <input
                  hidden
                  type="file"
                  accept=".xlsx"
                  onChange={(event) => void seedEjoos(event.target.files?.[0])}
                />
              </Button>
            )}
          </Box>

          <Divider orientation="vertical" flexItem />

          <Box sx={{ flex: 1 }}>
            <Typography variant="subtitle2">2. Новий 1ПБ (sh / Рух / archive)</Typography>
            <Button
              component="label"
              variant="outlined"
              startIcon={<CloudUploadOutlinedIcon />}
              disabled={!live?.current}
            >
              Завантажити 1ПБ і побудувати план
              <input
                hidden
                type="file"
                accept=".xlsx"
                onChange={(event) => void openPb(event.target.files?.[0])}
              />
            </Button>
            {pbSnapshot ? (
              <Typography variant="body2" sx={{ mt: 1 }}>
                1ПБ: {pbSnapshot.fileName}
              </Typography>
            ) : null}
          </Box>
        </Stack>

        {plan ? (
          <>
            <Divider />
            <Stack
              direction={{ xs: "column", md: "row" }}
              spacing={1}
              alignItems="center"
              justifyContent="space-between"
            >
              <Typography variant="subtitle2">
                3. Огляд змін ·{" "}
                {plan.sourceDateUnknown
                  ? "дата 1ПБ невідома"
                  : `джерело 1ПБ станом на ${plan.timesheetDayLabel} (табель лише по цей день)`}{" "}
                · всього {plan.ops.length}
              </Typography>
              <Stack direction="row" spacing={1} alignItems="center" flexWrap="wrap">
                <Select
                  size="small"
                  value={filter}
                  onChange={(event) =>
                    setFilter(event.target.value as FilterClass)
                  }
                >
                  <MenuItem value="ALL">Усі класи</MenuItem>
                  <MenuItem value="ready">Готово</MenuItem>
                  <MenuItem value="needs_input">Потрібна дія</MenuItem>
                  <MenuItem value="conflict">Конфлікти</MenuItem>
                </Select>
                <Select
                  size="small"
                  value={sheetFilter}
                  onChange={(event) => setSheetFilter(String(event.target.value))}
                >
                  <MenuItem value="ALL">Усі аркуші</MenuItem>
                  {sheetOptions.map((sheet) => (
                    <MenuItem key={sheet} value={sheet}>
                      {sheet}
                    </MenuItem>
                  ))}
                </Select>
                <input
                  value={nameQuery}
                  onChange={(event) => setNameQuery(event.target.value)}
                  placeholder="Пошук ПІБ / ID / індекс"
                  style={{
                    minWidth: 180,
                    padding: "6px 10px",
                    borderRadius: 6,
                    border: "1px solid #ccc",
                  }}
                />
                <Button size="small" onClick={selectVisibleReady}>
                  Вибрати видимі «готово»
                </Button>
                <Button size="small" onClick={clearVisibleChecks}>
                  Зняти видимі
                </Button>
                <Button size="small" onClick={downloadDraftProtocol}>
                  Чернетка протоколу
                </Button>
              </Stack>
            </Stack>

            <Stack direction="row" spacing={1} flexWrap="wrap">
              <Chip label={`Готово: ${plan.summary.ready}`} color="success" />
              <Chip
                label={`Потрібна дія: ${plan.summary.needsInput}`}
                color="warning"
              />
              <Chip
                label={`Конфлікт: ${plan.summary.conflict}`}
                color="error"
              />
              <Chip label={`Вибрано: ${selectedOps.length}`} />
            </Stack>

            <Box
              sx={{
                maxHeight: 420,
                overflow: "auto",
                border: "1px solid",
                borderColor: "divider",
                borderRadius: 1,
              }}
            >
              <table className="data-table" style={{ width: "100%" }}>
                <thead>
                  <tr>
                    <th style={{ width: 36 }} />
                    <th>Клас</th>
                    <th>Людина</th>
                    <th>Аркуш</th>
                    <th>Було</th>
                    <th>Стане</th>
                    <th>Джерело / чому</th>
                  </tr>
                </thead>
                <tbody>
                  {pagedOps.map((op) => (
                    <OpRow
                      key={op.id}
                      op={op}
                      checked={checkedIds.has(op.id)}
                      manualCode={manualCodes[op.id] || ""}
                      manualDestination={manualFields[op.id]?.destination || ""}
                      manualPlace={manualFields[op.id]?.place || ""}
                      manualDepartDate={manualFields[op.id]?.departDate || ""}
                      onToggle={() => toggleOp(op.id)}
                      onManualCode={(value) =>
                        setManualCodes((current) => ({
                          ...current,
                          [op.id]: value,
                        }))
                      }
                      onManualField={(field, value) =>
                        setManualFields((current) => ({
                          ...current,
                          [op.id]: { ...current[op.id], [field]: value },
                        }))
                      }
                    />
                  ))}
                </tbody>
              </table>
            </Box>

            <Stack direction="row" spacing={1} alignItems="center" flexWrap="wrap">
              <Typography variant="caption">
                Показано {pagedOps.length} з {visibleOps.length} (стор.{" "}
                {Math.min(page, pageCount - 1) + 1}/{pageCount})
              </Typography>
              <Button
                size="small"
                disabled={page <= 0}
                onClick={() => setPage((p) => Math.max(0, p - 1))}
              >
                ← Назад
              </Button>
              <Button
                size="small"
                disabled={page >= pageCount - 1}
                onClick={() => setPage((p) => Math.min(pageCount - 1, p + 1))}
              >
                Далі →
              </Button>
            </Stack>

            {plan.sourceDateUnknown ? (
              <Alert severity="warning" variant="outlined">
                {SOURCE_DATE_UNKNOWN_MESSAGE}
                <Stack direction="row" spacing={1} sx={{ mt: 1 }} alignItems="center">
                  <input
                    type="date"
                    value={sourceAsOf}
                    onChange={(event) => {
                      const value = event.target.value;
                      setSourceAsOf(value);
                      if (!value || !ejoosSnapshot || !pbSnapshot) return;
                      const gen = (planGenRef.current += 1);
                      setIsLoading(true);
                      void buildPlanOffThread(ejoosSnapshot, pbSnapshot, value)
                        .then((nextPlan) => {
                          if (gen !== planGenRef.current) return;
                          setPlan(nextPlan);
                        })
                        .catch((err) => {
                          if (gen !== planGenRef.current) return;
                          setError(
                            err instanceof Error
                              ? err.message
                              : "Не вдалося перерахувати план",
                          );
                        })
                        .finally(() => {
                          if (gen === planGenRef.current) setIsLoading(false);
                        });
                    }}
                  />
                </Stack>
              </Alert>
            ) : null}
            {plan.timesheetMonthHeaderUnknown &&
            !plan.monthRolloverRequired ? (
              <Typography variant="body2" className="ejoos-muted">
                Заголовок місяця в I2 не прочитано — місяць беремо з дати 1ПБ (
                {plan.timesheetDayLabel}).
              </Typography>
            ) : null}

            {plan.limitsNote ? (
              <Typography variant="caption" color="text.secondary">
                {plan.limitsNote}
              </Typography>
            ) : null}

            <Stack direction="row" spacing={1} alignItems="center">
              <Button
                variant="contained"
                color="primary"
                disabled={
                  !selectedOps.length ||
                  isLoading ||
                  planBlocksWorkbookApply(plan)
                }
                onClick={() => setConfirmOpen(true)}
              >
                Застосувати вибрані ({selectedOps.length})
              </Button>
              {blockedSelected.length ? (
                <Typography variant="caption" color="error">
                  Заблоковано серед вибраних: {blockedSelected.length}
                </Typography>
              ) : null}
            </Stack>

            {confirmOpen ? (
              <Alert
                severity="warning"
                action={
                  <Stack direction="row" spacing={1}>
                    <Button size="small" onClick={() => setConfirmOpen(false)}>
                      Скасувати
                    </Button>
                    <Button
                      size="small"
                      variant="contained"
                      color="warning"
                      disabled={Boolean(blockedSelected.length)}
                      onClick={() => void applySelected()}
                    >
                      Підтверджую запис
                    </Button>
                  </Stack>
                }
              >
                <Typography variant="subtitle2" sx={{ mb: 1 }}>
                  Буде створено нову версію ЕЖООС на базі v{live?.current?.version}.
                  Попередня версія залишиться в історії.
                </Typography>
                <Typography variant="body2" component="div">
                  Застосується <strong>{confirmSummary.total}</strong> змін:
                </Typography>
                <ul style={{ margin: "6px 0", paddingLeft: 18 }}>
                  {confirmSummary.byKind.map((item) => (
                    <li key={item.kind}>
                      {item.label}: {item.count}
                    </li>
                  ))}
                </ul>
                <Typography variant="caption" component="div">
                  ПІБ (перші 20):{" "}
                  {confirmSummary.names.slice(0, 20).join("; ") || "—"}
                  {confirmSummary.names.length > 20
                    ? ` … ще ${confirmSummary.names.length - 20}`
                    : ""}
                </Typography>
                {blockedSelected.length ? (
                  <Typography variant="body2" color="error" sx={{ mt: 1 }}>
                    Серед вибраних є блокери ({blockedSelected.length}) — зніміть
                    їх або дозаповніть.
                  </Typography>
                ) : null}
              </Alert>
            ) : null}
          </>
        ) : null}

        {lastResultFiles ? (
          <Alert severity="info">
            Останній результат: ЕЖООС v{lastResultFiles.version}. Файли вже
            завантажені. Можна знову скачати поточний ЕЖООС або протокол з історії
            версій.
            {lastProtocol ? (
              <Button
                size="small"
                sx={{ ml: 1 }}
                onClick={() =>
                  downloadTextFile(
                    `протокол_ЕЖООС_v${lastResultFiles.version}.txt`,
                    lastProtocol,
                  )
                }
              >
                Ще раз протокол
              </Button>
            ) : null}
          </Alert>
        ) : null}
      </Stack>
    </Box>
  );
}

const OpRow = memo(function OpRow({
  op,
  checked,
  manualCode,
  manualDestination,
  manualPlace,
  manualDepartDate,
  onToggle,
  onManualCode,
  onManualField,
}: {
  op: EjoosSyncOp;
  checked: boolean;
  manualCode: string;
  manualDestination: string;
  manualPlace: string;
  manualDepartDate: string;
  onToggle: () => void;
  onManualCode: (value: string) => void;
  onManualField: (
    field: "destination" | "place" | "departDate",
    value: string,
  ) => void;
}) {
  return (
    <tr>
      <td>
        <input
          type="checkbox"
          checked={checked}
          disabled={op.class === "conflict"}
          onChange={onToggle}
        />
      </td>
      <td>
        <Chip
          size="small"
          label={classLabel[op.class]}
          color={
            op.class === "ready"
              ? "success"
              : op.class === "conflict"
                ? "error"
                : "warning"
          }
        />
      </td>
      <td>
        <div>
          <strong>{op.fullName || "—"}</strong>
        </div>
        <Typography variant="caption" component="div">
          ID {op.personId || "—"} · індекс {op.positionIndex || "—"} ·{" "}
          {op.rank || "—"}
        </Typography>
      </td>
      <td>{op.sheet}</td>
      <td>
        <code>{op.before}</code>
      </td>
      <td>
        {op.kind === "timesheet_day" &&
        op.payload.clearStalePerson !== "1" &&
        timesheetOpNeedsManualCode(op) ? (
          <Select
            size="small"
            value={manualCode || op.payload.timesheetCode || ""}
            displayEmpty
            onChange={(event) => onManualCode(String(event.target.value))}
          >
            <MenuItem value="">код…</MenuItem>
            {["+", "вдр", "від", "ВП", "лік", "ЛП", "СЗЧ", "ЗБ", "пол"].map(
              (code) => (
                <MenuItem key={code} value={code}>
                  {code}
                </MenuItem>
              ),
            )}
          </Select>
        ) : op.kind === "exclude_transfer" && !op.payload.destination ? (
          <input
            value={manualDestination}
            placeholder="Куди вибув…"
            onChange={(event) =>
              onManualField("destination", event.target.value)
            }
            style={{ width: "100%", minWidth: 140 }}
          />
        ) : op.kind === "absent_upsert" &&
          !(op.payload.absenceType && op.payload.departDate) ? (
          <Stack spacing={0.5}>
            <code>{op.after}</code>
            <input
              value={manualPlace || op.payload.place || ""}
              placeholder="Місце…"
              onChange={(event) => onManualField("place", event.target.value)}
            />
            <input
              value={manualDepartDate || op.payload.departDate || ""}
              placeholder="Дата вибуття…"
              onChange={(event) =>
                onManualField("departDate", event.target.value)
              }
            />
          </Stack>
        ) : (
          <code>{manualCode || manualDestination || op.after}</code>
        )}
      </td>
      <td>
        <Typography variant="caption" component="div">
          {op.sourceRef}
        </Typography>
        <Typography variant="caption" color="text.secondary" component="div">
          {op.why}
        </Typography>
      </td>
    </tr>
  );
});
