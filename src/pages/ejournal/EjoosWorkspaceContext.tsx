import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useRef,
  useState,
  type ReactNode,
} from "react";
import {
  api,
  type BackendEjournalLiveState,
  type BackendEjournalLiveVersion,
  type BackendEjournalPbState,
  type BackendEjournalPbSource,
} from "../../api";
import {
  type ExcelWorkbookSnapshot,
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  applyConfirmedEjoosOps,
  base64ToFile,
  blobToBase64,
  downloadBlobFile,
  downloadTextFile,
  fileToBase64,
} from "./ejoosSyncApply";
import {
  acceptAllReady,
  collectedAcceptedOps,
  groupOpsIntoPersonChanges,
  patchPersonOpPayload,
  setPersonDecision,
  type EjoosDiffSession,
  type PersonChangeDecision,
} from "./ejoosPersonDiff";
import { buildEjoosSyncPlan, buildProtocolText, parseTimesheetDayFromPbName } from "./ejoosSyncPlan";
import { buildNormalizedSnapshotFromWorkbook } from "./ejoosNormalized";
import {
  getInitialEjoosTab,
  replaceEjoosTabInUrl,
  type EjoosWorkspaceTab,
} from "../../app/navigation";
import {
  assertEjoosWorkbook,
  assertPbWorkbook,
  detectWorkbookKind,
  ejoosDownloadFileName,
} from "./ejoosWorkbookKind";

export type { EjoosWorkspaceTab };

type EjoosWorkspaceContextValue = {
  tab: EjoosWorkspaceTab;
  setTab: (tab: EjoosWorkspaceTab) => void;
  live: BackendEjournalLiveState | null;
  refreshLive: () => Promise<BackendEjournalLiveState>;
  pbSources: BackendEjournalPbState | null;
  refreshPbSources: () => Promise<BackendEjournalPbState>;
  ejoosSnapshot: ExcelWorkbookSnapshot | null;
  pbSnapshot: ExcelWorkbookSnapshot | null;
  session: EjoosDiffSession | null;
  selectedPersonId: string | null;
  setSelectedPersonId: (id: string | null) => void;
  message: string;
  error: string;
  isLoading: boolean;
  seedEjoos: (file: File) => Promise<void>;
  /** Seed якщо немає версії; інакше нова версія з файлу (старі лишаються). */
  importEjoos: (file: File) => Promise<void>;
  loadEjoosFromDb: () => Promise<void>;
  ensureEjoosLoaded: () => Promise<ExcelWorkbookSnapshot | null>;
  /** Завантажити 1ПБ (sh / Рух / archive) і побудувати план операцій, якщо ЕЖООС доступний. */
  loadPb: (file: File) => Promise<void>;
  /** Зберегти поточний 1ПБ у БД. */
  savePbToDb: () => Promise<BackendEjournalPbSource | null>;
  /** Відкрити 1ПБ з БД і побудувати план операцій, якщо ЕЖООС доступний. */
  loadPbFromDb: (id?: string) => Promise<void>;
  analyzePb: (file: File) => Promise<void>;
  rebuildOperations: () => Promise<void>;
  setDecision: (personChangeId: string, decision: PersonChangeDecision) => void;
  patchOpPayload: (
    personChangeId: string,
    opId: string,
    payloadPatch: Record<string, string>,
  ) => void;
  acceptReady: () => void;
  applyAccepted: () => Promise<void>;
  /** Підтвердити одну людину і одразу записати її ops у нову версію ЕЖООС. */
  acceptAndApplyPerson: (personChangeId: string) => Promise<void>;
  downloadCurrentEjoos: () => Promise<void>;
  downloadVersion: (versionId: string, fileName?: string) => Promise<void>;
  downloadVersionProtocol: (version: BackendEjournalLiveVersion) => void;
  rollback: (versionId: string) => Promise<void>;
};

const EjoosWorkspaceContext = createContext<EjoosWorkspaceContextValue | null>(
  null,
);

export function EjoosWorkspaceProvider({ children }: { children: ReactNode }) {
  const [tab, setTabState] = useState<EjoosWorkspaceTab>(getInitialEjoosTab);
  const [live, setLive] = useState<BackendEjournalLiveState | null>(null);
  const [pbSources, setPbSources] = useState<BackendEjournalPbState | null>(
    null,
  );
  const [ejoosSnapshot, setEjoosSnapshot] =
    useState<ExcelWorkbookSnapshot | null>(null);
  const [pbSnapshot, setPbSnapshot] = useState<ExcelWorkbookSnapshot | null>(
    null,
  );
  const [session, setSession] = useState<EjoosDiffSession | null>(null);
  const [selectedPersonId, setSelectedPersonId] = useState<string | null>(null);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const pbFileBase64Ref = useRef<string | null>(null);

  const refreshLive = useCallback(async () => {
    const state = await api.getEjournalLive("1ПБ");
    setLive(state);
    return state;
  }, []);

  const refreshPbSources = useCallback(async () => {
    const state = await api.getEjournalPbSources("1ПБ");
    setPbSources(state);
    return state;
  }, []);

  useEffect(() => {
    void refreshLive().catch(() => {
      /* live ЕЖООС у БД опційний */
    });
    void refreshPbSources().catch(() => {
      /* 1ПБ у БД опційний */
    });
  }, [refreshLive, refreshPbSources]);

  const syncNormalizedFromSnapshot = useCallback(
    async (
      snapshot: ExcelWorkbookSnapshot,
      meta?: { versionId?: string; asOfDate?: string | null },
    ) => {
      const payload = buildNormalizedSnapshotFromWorkbook({
        workbook: snapshot,
        unitLabel: "1ПБ",
        versionId: meta?.versionId,
        asOfDate: meta?.asOfDate,
      });
      await api.syncEjournalNormalized(payload);
    },
    [],
  );

  const setTab = useCallback((next: EjoosWorkspaceTab) => {
    setTabState(next);
  }, []);

  useEffect(() => {
    replaceEjoosTabInUrl(tab);
  }, [tab]);

  useEffect(() => {
    const onPopState = () => {
      setTabState(getInitialEjoosTab());
    };
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, []);

  const ensureEjoosSnapshot = useCallback(async () => {
    if (ejoosSnapshot) return ejoosSnapshot;
    const state = live ?? (await refreshLive());
    if (!state.current) throw new Error("Спочатку завантажте канонічний ЕЖООС");
    const full = await api.getEjournalLiveFile(state.current.id, "1ПБ");
    if (!full.fileBase64) throw new Error("У версії немає файлу");
    const file = base64ToFile(
      full.fileBase64,
      full.sourceFileName || `ЕЖООС_v${full.version}.xlsx`,
    );
    const snapshot = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
    setEjoosSnapshot(snapshot);
    return snapshot;
  }, [ejoosSnapshot, live, refreshLive]);

  const seedEjoos = async (file: File) => {
    setIsLoading(true);
    setError("");
    try {
      const snapshot = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
      assertEjoosWorkbook(snapshot);
      setEjoosSnapshot(snapshot);
      setSession(null);

      const dayInfo = parseTimesheetDayFromPbName(file.name);
      let dbNote = "";
      try {
        const fileBase64 = await fileToBase64(file);
        const created = await api.seedEjournalLive({
          fileBase64,
          sourceFileName: file.name,
          asOfDate: dayInfo.label,
          unitLabel: "1ПБ",
          notes: "Початкове завантаження канонічного ЕЖООС",
        });
        await refreshLive();
        try {
          await syncNormalizedFromSnapshot(snapshot, {
            versionId: created.id,
            asOfDate: created.asOfDate ?? dayInfo.label,
          });
        } catch (syncErr) {
          dbNote += ` Нормалізовані таблиці не синхронізовано: ${
            syncErr instanceof Error ? syncErr.message : "невідома помилка"
          }.`;
        }
        dbNote = ` Збережено в БД як v${created.version}.` + dbNote;
      } catch (saveErr) {
        dbNote = ` Файл відкрито локально; БД: ${
          saveErr instanceof Error ? saveErr.message : "недоступна"
        }.`;
      }

      setMessage(`ЕЖООС завантажено: ${file.name}.${dbNote}`);
      setTab("shpo");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося зберегти ЕЖООС");
    } finally {
      setIsLoading(false);
    }
  };

  const importEjoos = async (file: File) => {
    const current = live?.current ?? (await refreshLive()).current;
    if (!current) {
      await seedEjoos(file);
      return;
    }

    const ok = window.confirm(
      `Завантажити «${file.name}» як нову версію ЕЖООС?\n\n` +
        `Поточна v${current.version} лишиться в історії.`,
    );
    if (!ok) return;

    setIsLoading(true);
    setError("");
    try {
      const snapshot = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
      assertEjoosWorkbook(snapshot);
      setEjoosSnapshot(snapshot);
      setSession(null);
      setSelectedPersonId(null);
      setPbSnapshot(null);

      const dayInfo = parseTimesheetDayFromPbName(file.name);
      let dbNote = "";
      try {
        const fileBase64 = await fileToBase64(file);
        const saved = await api.applyEjournalLive({
          baseVersionId: current.id,
          fileBase64,
          sourceFileName: file.name,
          asOfDate: dayInfo.label,
          unitLabel: "1ПБ",
          changeProtocol: {
            kind: "import",
            ops: [],
            summary: `Імпорт оновленого ЕЖООС з файлу «${file.name}»`,
            protocolText: [
              `Імпорт ЕЖООС`,
              `Файл: ${file.name}`,
              `База: v${current.version}`,
              `Станом на: ${dayInfo.label}`,
              `Час: ${new Date().toLocaleString("uk-UA")}`,
            ].join("\n"),
          },
          notes: `Імпорт повного файлу ЕЖООС: ${file.name}`,
        });
        await refreshLive();
        try {
          await syncNormalizedFromSnapshot(snapshot, {
            versionId: saved.id,
            asOfDate: saved.asOfDate ?? dayInfo.label,
          });
        } catch (syncErr) {
          dbNote += ` Нормалізовані таблиці не синхронізовано: ${
            syncErr instanceof Error ? syncErr.message : "невідома помилка"
          }.`;
        }
        dbNote =
          ` Збережено в БД як v${saved.version} (було v${current.version}).` +
          dbNote;
      } catch (saveErr) {
        dbNote = ` Файл відкрито локально; БД: ${
          saveErr instanceof Error ? saveErr.message : "недоступна"
        }.`;
      }

      setMessage(`ЕЖООС завантажено: ${file.name}.${dbNote}`);
      setTab("shpo");
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Не вдалося імпортувати ЕЖООС",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const loadEjoosFromDb = async () => {
    setIsLoading(true);
    setError("");
    try {
      await ensureEjoosSnapshot();
      setMessage("ЕЖООС відкрито з БД для аналізу");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося відкрити ЕЖООС");
    } finally {
      setIsLoading(false);
    }
  };

  const ensureEjoosLoaded = async () => {
    try {
      return await ensureEjoosSnapshot();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося відкрити ЕЖООС");
      return null;
    }
  };

  // Автоматично підтягнути файл ЕЖООС з БД для вкладок-таблиць.
  useEffect(() => {
    if (!live?.current || ejoosSnapshot) return;
    void ensureEjoosSnapshot().catch((err) => {
      setError(
        err instanceof Error ? err.message : "Не вдалося відкрити ЕЖООС з БД",
      );
    });
  }, [live, ejoosSnapshot, ensureEjoosSnapshot]);

  const buildSessionFromPb = useCallback(
    async (pb: ExcelWorkbookSnapshot) => {
      const ejoos = await ensureEjoosSnapshot();
      assertEjoosWorkbook(ejoos);
      const plan = buildEjoosSyncPlan(ejoos, pb);
      const nextSession = groupOpsIntoPersonChanges(plan, pb);
      setSession(nextSession);
      setSelectedPersonId(null);
      return nextSession;
    },
    [ensureEjoosSnapshot],
  );

  const loadPb = async (file: File) => {
    setIsLoading(true);
    setError("");
    try {
      const pb = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
      assertPbWorkbook(pb);
      setPbSnapshot(pb);
      setSelectedPersonId(null);

      let savedNote = "";
      try {
        const fileBase64 = await fileToBase64(file);
        pbFileBase64Ref.current = fileBase64;
        const dayInfo = parseTimesheetDayFromPbName(file.name);
        const saved = await api.uploadEjournalPb({
          fileBase64,
          sourceFileName: file.name,
          asOfDate: dayInfo.label,
          unitLabel: "1ПБ",
          notes: "1ПБ (sh / Рух / archive)",
        });
        await refreshPbSources();
        savedNote = ` Збережено в БД (${saved.sourceFileName || file.name}).`;
      } catch (saveErr) {
        savedNote = ` Файл у памʼяті, але не збережено в БД: ${
          saveErr instanceof Error ? saveErr.message : "помилка"
        }.`;
      }

      let analysisNote = "";
      let nextTab: EjoosWorkspaceTab = "pb";
      try {
        const nextSession = await buildSessionFromPb(pb);
        analysisNote =
          ` Операції: ${nextSession.counters.changes}, авто ` +
          `${nextSession.counters.autoReady}, перевірити ` +
          `${nextSession.counters.needsReview}, конфлікти ` +
          `${nextSession.counters.errors}.`;
      } catch (analysisErr) {
        setSession(null);
        analysisNote = ` Аналіз не виконано: ${
          analysisErr instanceof Error
            ? analysisErr.message
            : "ЕЖООС недоступний"
        }.`;
        nextTab = "pb";
      }

      setMessage(
        `1ПБ завантажено: ${file.name} · аркуші ${pb.sheets
          .map((sheet) => sheet.sheetName)
          .join(", ")}.${savedNote}${analysisNote}`,
      );
      setTab(nextTab);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося завантажити 1ПБ");
    } finally {
      setIsLoading(false);
    }
  };

  const analyzePb = async (file: File) => {
    await loadPb(file);
  };

  const savePbToDb = async () => {
    const base64 = pbFileBase64Ref.current;
    const fileName = pbSnapshot?.fileName || session?.pbFileName;
    if (!base64 || !fileName) {
      setError(
        "Немає 1ПБ у памʼяті — спочатку завантажте файл або відкрийте з БД",
      );
      return null;
    }
    setIsLoading(true);
    setError("");
    try {
      const saved = await api.uploadEjournalPb({
        fileBase64: base64,
        sourceFileName: fileName,
        asOfDate: session?.plan.timesheetDayLabel,
        unitLabel: "1ПБ",
        notes: "1ПБ збережено вручну",
      });
      await refreshPbSources();
      setMessage(`1ПБ збережено в БД: ${saved.sourceFileName || fileName}`);
      return saved;
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося зберегти 1ПБ");
      return null;
    } finally {
      setIsLoading(false);
    }
  };

  const rebuildOperations = async () => {
    if (!pbSnapshot) {
      setError("Немає відкритого 1ПБ для аналізу");
      return;
    }
    setIsLoading(true);
    setError("");
    try {
      const state = await refreshLive();
      const current = state.current;
      if (!current) throw new Error("Спочатку завантажте або відкрийте ЕЖООС");
      const full = await api.getEjournalLiveFile(current.id, "1ПБ");
      if (!full.fileBase64) throw new Error("У поточній версії ЕЖООС немає файлу");
      const file = base64ToFile(
        full.fileBase64,
        full.sourceFileName || `ЕЖООС_v${full.version}.xlsx`,
      );
      const snapshot = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
      assertEjoosWorkbook(snapshot);
      setEjoosSnapshot(snapshot);
      const plan = buildEjoosSyncPlan(snapshot, pbSnapshot);
      const nextSession = groupOpsIntoPersonChanges(plan, pbSnapshot);
      setSession(nextSession);
      setSelectedPersonId(null);
      setMessage(
        `Операції перебудовано з поточного ЕЖООС v${current.version}: ` +
          `${nextSession.counters.changes}, авто ${nextSession.counters.autoReady}, ` +
          `перевірити ${nextSession.counters.needsReview}, конфлікти ${nextSession.counters.errors}.`,
      );
      setTab("changes");
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Не вдалося перебудувати операції",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const loadPbFromDb = async (id?: string) => {
    setIsLoading(true);
    setError("");
    try {
      const remote = await api.getEjournalPbFile(id, "1ПБ");
      if (!remote.fileBase64) {
        throw new Error("Сервер не повернув файл 1ПБ");
      }
      const fileName = remote.sourceFileName || "1PB.xlsx";
      pbFileBase64Ref.current = remote.fileBase64;
      const file = base64ToFile(remote.fileBase64, fileName);
      const pb = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
      assertPbWorkbook(pb);
      setPbSnapshot(pb);
      setSelectedPersonId(null);
      await refreshPbSources();

      let analysisNote = "";
      try {
        const nextSession = await buildSessionFromPb(pb);
        analysisNote =
          ` Операції: ${nextSession.counters.changes}, авто ` +
          `${nextSession.counters.autoReady}, перевірити ` +
          `${nextSession.counters.needsReview}, конфлікти ` +
          `${nextSession.counters.errors}.`;
      } catch (analysisErr) {
        setSession(null);
        analysisNote = ` Аналіз не виконано: ${
          analysisErr instanceof Error
            ? analysisErr.message
            : "ЕЖООС недоступний"
        }.`;
      }

      setMessage(
        `Відкрито 1ПБ з БД: ${fileName} · аркуші ${pb.sheets
          .map((sheet) => sheet.sheetName)
          .join(", ")}.${analysisNote}`,
      );
      setTab("pb");
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Не вдалося відкрити 1ПБ з БД",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const setDecision = (personChangeId: string, decision: PersonChangeDecision) => {
    setSession((current) =>
      current ? setPersonDecision(current, personChangeId, decision) : current,
    );
  };

  const patchOpPayload = (
    personChangeId: string,
    opId: string,
    payloadPatch: Record<string, string>,
  ) => {
    setSession((current) =>
      current
        ? patchPersonOpPayload(current, personChangeId, opId, payloadPatch)
        : current,
    );
  };

  const acceptReady = () => {
    setSession((current) => (current ? acceptAllReady(current) : current));
  };

  const runApplyOps = async (
    workingSession: EjoosDiffSession,
    ops: import("./ejoosSyncPlan").EjoosSyncOp[],
    note: string,
  ) => {
    if (!ejoosSnapshot || !live?.current) {
      throw new Error("Немає сесії змін або канонічного ЕЖООС");
    }
    assertEjoosWorkbook(ejoosSnapshot);
    if (!ops.length) throw new Error("Немає підтверджених змін");
    const result = await applyConfirmedEjoosOps({
      ejoos: ejoosSnapshot,
      plan: workingSession.plan,
      ops,
    });
    const fileBase64 = await blobToBase64(result.blob);
    const protocolText = buildProtocolText(workingSession.plan, ops, {
      actor: "operator",
      at: new Date().toLocaleString("uk-UA"),
      version: live.current.version + 1,
    });
    const saved = await api.applyEjournalLive({
      baseVersionId: live.current.id,
      fileBase64,
      sourceFileName: result.fileName,
      asOfDate: workingSession.plan.timesheetDayLabel,
      unitLabel: "1ПБ",
      sourcePbFileName: pbSnapshot?.fileName || workingSession.pbFileName,
      sourcePbSha256: pbSources?.current?.sha256,
      changeProtocol: { ...result.changeProtocol, protocolText },
      notes: note,
    });
    // Не качаємо файл на кожне застосування — експорт лише з вкладки «Експорт».
    const state = await refreshLive();
    let normalizedWarning = "";
    if (state.current) {
      const full = await api.getEjournalLiveFile(state.current.id, "1ПБ");
      if (full.fileBase64) {
        const file = base64ToFile(
          full.fileBase64,
          full.sourceFileName || `ЕЖООС_v${full.version}.xlsx`,
        );
        const nextSnap = await readWorkbookSnapshot(
          file,
          EJOOS_SYNC_READ_OPTIONS,
        );
        setEjoosSnapshot(nextSnap);
        try {
          await syncNormalizedFromSnapshot(nextSnap, {
            versionId: saved.id,
            asOfDate: saved.asOfDate,
          });
        } catch (syncErr) {
          normalizedWarning = ` Нормалізовані таблиці не синхронізовано: ${
            syncErr instanceof Error ? syncErr.message : "невідома помилка"
          }.`;
        }
      }
    }
    return { saved, normalizedWarning };
  };

  const applyAccepted = async () => {
    if (!session || !ejoosSnapshot || !live?.current) {
      setError("Немає сесії змін або канонічного ЕЖООС");
      return;
    }
    const ops = collectedAcceptedOps(session);
    if (!ops.length) {
      setError("Немає підтверджених змін");
      return;
    }
    setIsLoading(true);
    setError("");
    try {
      const { saved, normalizedWarning } = await runApplyOps(
        session,
        ops,
        `Застосовано ${ops.length} ops / ${session.people.filter((p) => p.decision === "accepted").length} людей`,
      );
      setSession(null);
      setMessage(
        `Записано ЕЖООС v${saved.version}. Застосовано ${ops.length} змін.${normalizedWarning}`,
      );
      setTab("import");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося застосувати");
    } finally {
      setIsLoading(false);
    }
  };

  const acceptAndApplyPerson = async (personChangeId: string) => {
    if (!session || !ejoosSnapshot || !live?.current) {
      setError("Немає сесії змін або канонічного ЕЖООС");
      return;
    }
    const person = session.people.find((item) => item.id === personChangeId);
    if (!person) {
      setError("Людину не знайдено в сесії");
      return;
    }
    const excludeOp = person.ops.find((op) => op.kind === "exclude_transfer");
    if (excludeOp && !excludeOp.payload.destination?.trim()) {
      setError("Вкажіть «куди вибув» перед застосуванням переведення");
      return;
    }
    if (person.severity === "conflict") {
      setError("Конфлікт — спочатку розберіть вручну");
      return;
    }

    const nextSession = setPersonDecision(session, personChangeId, "accepted");
    setSession(nextSession);
    const ops = nextSession.people
      .find((item) => item.id === personChangeId)
      ?.ops.filter((op) => op.class !== "conflict");
    if (!ops?.length) {
      setError("Немає ops для застосування");
      return;
    }

    const ok = window.confirm(
      `Застосувати переведення для ${person.fullName} зараз?\nБуде створено нову версію ЕЖООС.`,
    );
    if (!ok) return;

    setIsLoading(true);
    setError("");
    try {
      const { saved, normalizedWarning } = await runApplyOps(
        nextSession,
        ops,
        `Переведення: ${person.fullName} · ${ops.length} ops`,
      );
      // Прибираємо цю людину з сесії, решту змін лишаємо
      setSession({
        ...nextSession,
        people: nextSession.people.filter((item) => item.id !== personChangeId),
        counters: {
          ...nextSession.counters,
          changes: Math.max(0, nextSession.counters.changes - 1),
        },
      });
      setSelectedPersonId(null);
      setMessage(
        `Переведення «${person.fullName}» записано в ЕЖООС v${saved.version}. Файл не качається — експорт з вкладки «Експорт», коли закінчите всі зміни.${normalizedWarning}`,
      );
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося застосувати");
    } finally {
      setIsLoading(false);
    }
  };

  const downloadCurrentEjoos = async () => {
    if (!live?.current) return;
    setIsLoading(true);
    setError("");
    try {
      const full = await api.getEjournalLiveFile(live.current.id);
      if (!full.fileBase64) throw new Error("Немає файлу");
      const downloadName = ejoosDownloadFileName(
        full.version,
        full.asOfDate,
        full.sourceFileName,
      );
      const file = base64ToFile(full.fileBase64, downloadName);
      const snapshot = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
      const kind = detectWorkbookKind(snapshot);
      if (kind === "pb_1pb") {
        throw new Error(
          `У БД збережено файл 1ПБ (sh / Рух / archive), а не ЕЖООС.\n` +
            `Імпортуйте канонічний ЕЖООС (ШПО, ООС, Виключені, Табель), ` +
            `потім проаналізуйте 1ПБ і застосуйте зміни — скачається оновлений ЕЖООС.`,
        );
      }
      if (kind !== "ejoos") {
        throw new Error(
          `Файл у БД не схожий на ЕЖООС (немає ШПО / ООС / Табель). Перезавантажте канонічний журнал.`,
        );
      }
      downloadBlobFile(downloadName, file);
      setMessage(`Скачано ${downloadName} (ЕЖООС v${full.version})`);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося скачати");
    } finally {
      setIsLoading(false);
    }
  };

  const downloadVersion = async (versionId: string, fileName?: string) => {
    setIsLoading(true);
    setError("");
    try {
      const full = await api.getEjournalLiveFile(versionId);
      if (!full.fileBase64) throw new Error("Немає файлу");
      const downloadName =
        fileName ||
        ejoosDownloadFileName(full.version, full.asOfDate, full.sourceFileName);
      const file = base64ToFile(full.fileBase64, downloadName);
      downloadBlobFile(downloadName, file);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося скачати");
    } finally {
      setIsLoading(false);
    }
  };

  const downloadVersionProtocol = (version: BackendEjournalLiveVersion) => {
    const protocol = version.changeProtocol as { protocolText?: string } | null;
    const text =
      (protocol && typeof protocol.protocolText === "string"
        ? protocol.protocolText
        : "") || `Протокол v${version.version}\nНемає тексту.`;
    downloadTextFile(`протокол_ЕЖООС_v${version.version}.txt`, text);
  };

  const rollback = async (versionId: string) => {
    if (
      !window.confirm(
        "Зробити цю версію поточною? Буде створено нову версію-копію.",
      )
    ) {
      return;
    }
    setIsLoading(true);
    setError("");
    try {
      const saved = await api.rollbackEjournalLive({
        targetVersionId: versionId,
        unitLabel: "1ПБ",
      });
      const state = await refreshLive();
      const current = state.current ?? saved;
      const full = await api.getEjournalLiveFile(current.id, "1ПБ");
      if (!full.fileBase64) throw new Error("Відкат виконано, але файл не повернувся з БД");
      const file = base64ToFile(
        full.fileBase64,
        full.sourceFileName || `ЕЖООС_v${full.version}.xlsx`,
      );
      const snapshot = await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS);
      assertEjoosWorkbook(snapshot);
      setEjoosSnapshot(snapshot);

      let analysisNote = "";
      let nextTab: EjoosWorkspaceTab = "history";
      if (pbSnapshot) {
        const plan = buildEjoosSyncPlan(snapshot, pbSnapshot);
        const nextSession = groupOpsIntoPersonChanges(plan, pbSnapshot);
        setSession(nextSession);
        setSelectedPersonId(null);
        analysisNote =
          ` Операції перебудовано: ${nextSession.counters.changes}, авто ` +
          `${nextSession.counters.autoReady}, перевірити ` +
          `${nextSession.counters.needsReview}, конфлікти ` +
          `${nextSession.counters.errors}.`;
        nextTab = "changes";
      } else {
        setSession(null);
      }

      setMessage(`Відкат: поточна v${saved.version}.${analysisNote}`);
      setTab(nextTab);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося відкотити");
    } finally {
      setIsLoading(false);
    }
  };

  const value: EjoosWorkspaceContextValue = {
    tab,
    setTab,
    live,
    refreshLive,
    pbSources,
    refreshPbSources,
    ejoosSnapshot,
    pbSnapshot,
    session,
    selectedPersonId,
    setSelectedPersonId,
    message,
    error,
    isLoading,
    seedEjoos,
    importEjoos,
    loadEjoosFromDb,
    ensureEjoosLoaded,
    loadPb,
    analyzePb,
    rebuildOperations,
    savePbToDb,
    loadPbFromDb,
    setDecision,
    patchOpPayload,
    acceptReady,
    applyAccepted,
    acceptAndApplyPerson,
    downloadCurrentEjoos,
    downloadVersion,
    downloadVersionProtocol,
    rollback,
  };

  return (
    <EjoosWorkspaceContext.Provider value={value}>
      {children}
    </EjoosWorkspaceContext.Provider>
  );
}

export function useEjoosWorkspace() {
  const ctx = useContext(EjoosWorkspaceContext);
  if (!ctx) {
    throw new Error("useEjoosWorkspace must be used within EjoosWorkspaceProvider");
  }
  return ctx;
}
