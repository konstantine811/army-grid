import {
  useCallback,
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
} from "../../api";
import {
  type ExcelWorkbookSnapshot,
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../../excelRoundTrip";
import { readEjoosWorkbookSnapshot } from "./ejoosTimesheetPersonRows";
import { formatApplyErrorWithTimesheetDump } from "./ejoosTimesheetDebugDump";
import {
  applyConfirmedEjoosOps,
  base64ToFile,
  blobToBase64,
  downloadBlobFile,
  downloadTextFile,
  fileToBase64,
} from "./ejoosSyncApply";
import {
  graftWorkbookStyles,
  sanitizeEjoosWorkbookBlob,
} from "./ejoosWorkbookSanitize";
import {
  fillEjoosSheetFromAnketa,
  formatEjoosAnketaFillReport,
  type EjoosAnketaFillMode,
  type EjoosAnketaFillTarget,
} from "./ejoosAnketaFill";
import {
  acceptAllReady,
  collectedAcceptedOps,
  collectedWritableAcceptedOps,
  personIsInformationalOnly,
  personCanEnterApplyQueue,
  isWorkbookApplyOp,
  mergePersonDecisions,
  patchPersonOpPayload,
  dismissPersonFromSession,
  setPersonDecision,
  setPersonDecisions,
  type EjoosDiffSession,
  type PersonChangeDecision,
} from "./ejoosPersonDiff";
import { personOpsBlockApply } from "./ejoosOpRequirements";
import {
  buildProtocolText,
  collectProcessedMovementKeys,
  planBlocksWorkbookApply,
  resolveJournalTimesheetDay,
  workbookApplyBlockMessage,
} from "./ejoosSyncPlan";
import { buildNormalizedSnapshotFromWorkbook } from "./ejoosNormalized";
import {
  getInitialEjoosTab,
  replaceEjoosTabInUrl,
  type EjoosWorkspaceTab,
} from "../../app/navigation";
import { readOperatorSettings } from "./ejoosStatusMap";
import { runHeavyJob } from "../../workers/runHeavyJob";
import {
  assertEjoosWorkbook,
  assertPbWorkbook,
  ejoosDownloadFileName,
} from "./ejoosWorkbookKind";
import {
  EjoosWorkspaceContext,
  type EjoosWorkspaceContextValue,
} from "./ejoosWorkspaceState";

export type { EjoosWorkspaceTab };

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
  const pbLoadPromiseRef =
    useRef<Promise<ExcelWorkbookSnapshot | null> | null>(null);

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
    const snapshot = await readEjoosWorkbookSnapshot(file);
    setEjoosSnapshot(snapshot);
    return snapshot;
  }, [ejoosSnapshot, live, refreshLive]);

  const ensurePbSnapshot = useCallback(
    async (id?: string) => {
      if (!id && pbSnapshot) return pbSnapshot;
      if (!id && pbLoadPromiseRef.current) return pbLoadPromiseRef.current;

      const load = async () => {
        const state = id ? pbSources : pbSources ?? (await refreshPbSources());
        const sourceId = id || state?.current?.id;
        if (!sourceId) return null;

        const remote = await api.getEjournalPbFile(sourceId, "1ПБ");
        if (!remote.fileBase64) {
          throw new Error("Сервер не повернув файл 1ПБ");
        }
        const fileName = remote.sourceFileName || "1PB.xlsx";
        pbFileBase64Ref.current = remote.fileBase64;
        const file = base64ToFile(remote.fileBase64, fileName);
        const snapshot = await readWorkbookSnapshot(
          file,
          EJOOS_SYNC_READ_OPTIONS,
        );
        assertPbWorkbook(snapshot);
        setPbSnapshot(snapshot);
        setSelectedPersonId(null);
        return snapshot;
      };

      if (id) return load();
      const pending = load().finally(() => {
        if (pbLoadPromiseRef.current === pending) {
          pbLoadPromiseRef.current = null;
        }
      });
      pbLoadPromiseRef.current = pending;
      return pending;
    },
    [pbSnapshot, pbSources, refreshPbSources],
  );

  const seedEjoos = async (file: File) => {
    setIsLoading(true);
    setError("");
    try {
      const snapshot = await readEjoosWorkbookSnapshot(file);
      assertEjoosWorkbook(snapshot);
      setEjoosSnapshot(snapshot);
      setSession(null);

      const dayInfo = resolveJournalTimesheetDay(file.name);
      let dbNote = "";
      try {
        const fileBase64 = await blobToBase64(file);
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
      const snapshot = await readEjoosWorkbookSnapshot(file);
      assertEjoosWorkbook(snapshot);
      setEjoosSnapshot(snapshot);
      setSession(null);
      setSelectedPersonId(null);
      setPbSnapshot(null);

      const dayInfo = resolveJournalTimesheetDay(file.name);
      let dbNote = "";
      try {
        const fileBase64 = await blobToBase64(file);
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

  const sourceAsOfOverrideRef = useRef("");

  const rebuildSessionFromSnapshots = useCallback(
    async (
      ejoos: ExcelWorkbookSnapshot,
      pb: ExcelWorkbookSnapshot,
      versions?: BackendEjournalLiveVersion[] | null,
      sourceAsOfDate?: string,
    ) => {
      const asOf = sourceAsOfDate ?? sourceAsOfOverrideRef.current;
      return runHeavyJob({
        type: "ejoosSession",
        ejoos,
        pb,
        processedMovementKeys: [
          ...collectProcessedMovementKeys(versions ?? undefined),
        ],
        sourceAsOfDate: asOf || undefined,
        statusRules: readOperatorSettings().statusRules,
      });
    },
    [],
  );

  const buildSessionFromPb = useCallback(
    async (pb: ExcelWorkbookSnapshot) => {
      const ejoos = await ensureEjoosSnapshot();
      assertEjoosWorkbook(ejoos);
      const nextSession = await rebuildSessionFromSnapshots(
        ejoos,
        pb,
        live?.versions,
      );
      setSession(nextSession);
      setSelectedPersonId(null);
      return nextSession;
    },
    [ensureEjoosSnapshot, live?.versions, rebuildSessionFromSnapshots],
  );

  // 1ПБ і план операцій потрібні лише на «Зміни» / «1ПБ».
  // На Експорті / Історії / аркушах це валить вкладку (xlsx-populate усієї книги).
  const needsPbSession = tab === "changes" || tab === "pb";

  useEffect(() => {
    if (!needsPbSession) return;
    if (!pbSources?.current || pbSnapshot) return;
    setIsLoading(true);
    setError("");
    setMessage("Завантажую 1ПБ з БД та будую операції…");
    void ensurePbSnapshot()
      .then(async (pb) => {
        if (!pb) return;
        try {
          const nextSession = await buildSessionFromPb(pb);
          setMessage(
            `1ПБ автоматично відкрито з БД: ${pb.fileName}. Операції: ${nextSession.counters.changes}.`,
          );
        } catch (analysisErr) {
          setSession(null);
          setMessage(
            `1ПБ автоматично відкрито з БД: ${pb.fileName}. Аналіз не виконано: ${
              analysisErr instanceof Error
                ? analysisErr.message
                : "ЕЖООС недоступний"
            }.`,
          );
        }
      })
      .catch((err) => {
        setError(
          err instanceof Error
            ? err.message
            : "Не вдалося автоматично відкрити 1ПБ з БД",
        );
      })
      .finally(() => {
        setIsLoading(false);
      });
  }, [
    needsPbSession,
    pbSources,
    pbSnapshot,
    ensurePbSnapshot,
    buildSessionFromPb,
  ]);

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
        const dayInfo = resolveJournalTimesheetDay(file.name);
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

  const rebuildOperations = async (sourceAsOfDate?: string) => {
    if (sourceAsOfDate !== undefined) {
      sourceAsOfOverrideRef.current = sourceAsOfDate;
    }
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
      const snapshot = await readEjoosWorkbookSnapshot(file);
      assertEjoosWorkbook(snapshot);
      setEjoosSnapshot(snapshot);
      const nextSession = await rebuildSessionFromSnapshots(
        snapshot,
        pbSnapshot,
        state.versions,
      );
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

  const dismissPerson = (personChangeId: string) => {
    setSession((current) =>
      current ? dismissPersonFromSession(current, personChangeId) : current,
    );
    setSelectedPersonId((current) =>
      current === personChangeId ? null : current,
    );
    setError("");
    setMessage("Запис прибрано з операцій. У ЕЖООС нічого не змінювалось.");
  };

  const setDecisions = (
    personChangeIds: string[],
    decision: PersonChangeDecision,
  ) => {
    setSession((current) =>
      current
        ? setPersonDecisions(current, personChangeIds, decision)
        : current,
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
    const protocolText = buildProtocolText(workingSession.plan, ops, {
      actor: "operator",
      at: new Date().toLocaleString("uk-UA"),
      version: live.current.version + 1,
    });
    const changeProtocol = { ...result.changeProtocol, protocolText };
    const fileBase64 = await blobToBase64(result.blob);
    const saved = await api.applyEjournalLive({
      baseVersionId: live.current.id,
      fileBase64,
      sourceFileName: result.fileName,
      asOfDate: workingSession.plan.timesheetDayLabel,
      unitLabel: "1ПБ",
      sourcePbFileName: pbSnapshot?.fileName || workingSession.pbFileName,
      sourcePbSha256: pbSources?.current?.sha256,
      changeProtocol,
      notes: note,
    });
    // Не качаємо файл на кожне застосування — експорт лише з вкладки «Експорт».
    const state = await refreshLive();
    let normalizedWarning = "";
    const localAppliedFile = base64ToFile(fileBase64, result.fileName);
    const nextEjoosSnapshot = await readEjoosWorkbookSnapshot(localAppliedFile);
    setEjoosSnapshot(nextEjoosSnapshot);
    try {
      await syncNormalizedFromSnapshot(nextEjoosSnapshot, {
        versionId: saved.id,
        asOfDate: saved.asOfDate,
      });
    } catch (syncErr) {
      normalizedWarning = ` Нормалізовані таблиці не синхронізовано: ${
        syncErr instanceof Error ? syncErr.message : "невідома помилка"
      }.`;
    }
    const refreshedVersions = [...(state.versions ?? [])];
    const savedIndex = refreshedVersions.findIndex(
      (version) => version.id === saved.id || version.version === saved.version,
    );
    if (savedIndex >= 0) {
      refreshedVersions[savedIndex] = {
        ...refreshedVersions[savedIndex],
        changeProtocol,
      };
    } else {
      refreshedVersions.push({
        ...saved,
        changeProtocol,
      });
    }
    return {
      saved,
      normalizedWarning,
      nextEjoosSnapshot,
      liveVersions: refreshedVersions,
    };
  };

  const applyAccepted = async (personIds?: string[]) => {
    if (!session || !ejoosSnapshot || !live?.current) {
      setError("Немає сесії змін або канонічного ЕЖООС");
      return;
    }
    if (planBlocksWorkbookApply(session.plan)) {
      setError(workbookApplyBlockMessage(session.plan));
      return;
    }
    const forceIds = new Set(personIds ?? []);
    const queuedAcceptedPeople = session.people.filter(
      (person) =>
        person.decision === "accepted" || forceIds.has(person.id),
    );
    const applicablePeople = queuedAcceptedPeople.filter(personCanEnterApplyQueue);
    const skippedBlockedPeople = queuedAcceptedPeople.filter(
      (person) =>
        !personCanEnterApplyQueue(person) &&
        !personIsInformationalOnly(person.ops),
    );
    const applicableIds = new Set(applicablePeople.map((person) => person.id));
    const applicableSession: EjoosDiffSession = {
      ...session,
      people: session.people.map((person) =>
        applicableIds.has(person.id)
          ? { ...person, decision: "accepted" as const }
          : { ...person, decision: "pending" as const },
      ),
    };
    const accepted = collectedAcceptedOps(applicableSession);
    const ops = collectedWritableAcceptedOps(applicableSession);
    if (!queuedAcceptedPeople.length) {
      setError("Немає людей у черзі застосування");
      return;
    }
    if (!accepted.length) {
      setError(
        `Масове застосування заблоковано, доки не уточнено дані: ${skippedBlockedPeople
          .slice(0, 4)
          .map((person) => person.fullName)
          .join(", ")}${skippedBlockedPeople.length > 4 ? "…" : ""}`,
      );
      return;
    }
    if (!ops.length) {
      setError("");
      setMessage(
        "У черзі немає змін для журналу. Позначки ПІБ / ID / звання виправляють у джерелах вручну — у ЕЖООС нічого не пишемо.",
      );
      return;
    }
    const missingDest = applicableSession.people.filter(
      (person) =>
        person.decision === "accepted" &&
        person.ops.some(
          (op) =>
            op.kind === "exclude_transfer" && !op.payload.destination?.trim(),
        ),
    );
    if (missingDest.length) {
      setError(
        `Спочатку вкажіть «куди вибув»: ${missingDest
          .slice(0, 4)
          .map((person) => person.fullName)
          .join(", ")}${missingDest.length > 4 ? "…" : ""}`,
      );
      return;
    }
    if (personOpsBlockApply(ops)) {
      setError(
        "У черзі є неповне переведення, немає наказу/«куди вибув», або відкритий СЗЧ суперечить новій постановці.",
      );
      return;
    }
    setIsLoading(true);
    setError("");
    try {
      const queuedPeople = applicablePeople.length;
      const skippedNotes = accepted.length - ops.length;
      const { saved, normalizedWarning, nextEjoosSnapshot, liveVersions } =
        await runApplyOps(
          session,
          ops,
          `Застосовано ${ops.length} ops / ${queuedPeople} людей з черги` +
            (skippedNotes
              ? ` · пропущено ${skippedNotes} позначок даних`
              : ""),
      );
      if (pbSnapshot && nextEjoosSnapshot) {
        const completedIds = new Set([
          ...applicablePeople.map((person) => person.id),
          ...queuedAcceptedPeople
            .filter(personIsInformationalOnly)
            .map((person) => person.id),
        ]);
        const previousForMerge: EjoosDiffSession = {
          ...session,
          people: session.people.map((person) =>
            completedIds.has(person.id)
              ? { ...person, decision: "pending" as const }
              : person,
          ),
        };
        const rebuilt = mergePersonDecisions(
          await rebuildSessionFromSnapshots(
            nextEjoosSnapshot,
            pbSnapshot,
            liveVersions,
          ),
          previousForMerge,
        );
        setSession(rebuilt);
      } else {
        setSession(null);
      }
      setSelectedPersonId(null);
      setMessage(
        `Записано ЕЖООС v${saved.version}. Застосовано ${ops.length} змін з черги.` +
          (skippedNotes
            ? ` Позначки даних (${skippedNotes}) пропущено — їх виправляють у джерелах.`
            : "") +
          (skippedBlockedPeople.length
            ? ` Не застосовано ${skippedBlockedPeople.length} неготових: ${skippedBlockedPeople
                .slice(0, 4)
                .map((person) => person.fullName)
                .join(", ")}${skippedBlockedPeople.length > 4 ? "…" : ""}.`
            : "") +
          (pbSnapshot && nextEjoosSnapshot ? " Операції перераховано." : "") +
          ` Файл не качається — експорт з вкладки «Експорт».${normalizedWarning}`,
      );
    } catch (err) {
      setError(formatApplyErrorWithTimesheetDump(ejoosSnapshot, err));
    } finally {
      setIsLoading(false);
    }
  };

  const acceptAndApplyPerson = async (personChangeId: string) => {
    if (!session || !ejoosSnapshot || !live?.current) {
      setError("Немає сесії змін або канонічного ЕЖООС");
      return;
    }
    if (planBlocksWorkbookApply(session.plan)) {
      setError(workbookApplyBlockMessage(session.plan));
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

    if (personIsInformationalOnly(ops)) {
      setError("");
      setSession(dismissPersonFromSession(nextSession, personChangeId));
      setSelectedPersonId(null);
      const cancelNotInSh = ops.every(
        (op) =>
          op.payload.type === "TRANSFER_CANCELLED" &&
          op.payload.reviewReason === "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH",
      );
      setMessage(
        cancelNotInSh
          ? `Перегляд «${person.fullName}» підтверджено. У ЕЖООС нічого не змінюється — людини немає в актуальній sh, аркуші не чіпаємо.`
          : `Перегляд «${person.fullName}» підтверджено. У ЕЖООС нічого не змінюється — звання / ПІБ / ID виправляють у джерелах.`,
      );
      return;
    }

    if (personOpsBlockApply(ops)) {
      setError(
        "У цій картці є неповне переведення або суперечливий статус. Спочатку уточніть дані.",
      );
      return;
    }
    const applyOps = ops.filter(isWorkbookApplyOp);
    if (!applyOps.length) {
      setError(
        `Немає змін для запису в ЕЖООС по «${person.fullName}». Нотатки без apply не підтверджуються окремо.`,
      );
      return;
    }

    setIsLoading(true);
    setError("");
    try {
      const { saved, normalizedWarning, nextEjoosSnapshot, liveVersions } =
        await runApplyOps(
          nextSession,
          applyOps,
          `${person.fullName} · ${applyOps.length} ops`,
        );
      // План залежить від стану ЕЖООС: після змін по ХУБАЄВУ конфлікт АТРАХОВА
      // має зникнути одразу, без перезавантаження сторінки.
      if (pbSnapshot && nextEjoosSnapshot) {
        const rebuilt = mergePersonDecisions(
          await rebuildSessionFromSnapshots(
            nextEjoosSnapshot,
            pbSnapshot,
            liveVersions,
          ),
          nextSession,
        );
        setSession(rebuilt);
      } else {
        setSession({
          ...nextSession,
          people: nextSession.people.filter(
            (item) => item.id !== personChangeId,
          ),
          counters: {
            ...nextSession.counters,
            changes: Math.max(0, nextSession.counters.changes - 1),
          },
        });
      }
      setSelectedPersonId(null);
      setMessage(
        `Зміни «${person.fullName}» записано в ЕЖООС v${saved.version}.` +
              (pbSnapshot && nextEjoosSnapshot
                ? " Операції перераховано з оновленого файлу."
                : "") +
              ` Файл не качається — експорт з вкладки «Експорт», коли закінчите всі зміни.${normalizedWarning}`,
      );
    } catch (err) {
      setError(formatApplyErrorWithTimesheetDump(ejoosSnapshot, err));
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
      const cleaned = await sanitizeEjoosWorkbookBlob(file);
      downloadBlobFile(downloadName, cleaned);
      const warnCorrupt =
        full.version >= 26
          ? " Якщо Excel скаржиться на файл — у «Історії» відкотіться на версію до v26 і скачайте її."
          : "";
      setMessage(`Скачано ${downloadName} (ЕЖООС v${full.version}).${warnCorrupt}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося скачати");
    } finally {
      setIsLoading(false);
    }
  };

  const downloadVersion = async (versionId: string, _fileName?: string) => {
    setIsLoading(true);
    setError("");
    try {
      const full = await api.getEjournalLiveFile(versionId);
      if (!full.fileBase64) throw new Error("Немає файлу");
      const downloadName = ejoosDownloadFileName(
        full.version,
        full.asOfDate,
        full.sourceFileName,
      );
      const file = base64ToFile(full.fileBase64, downloadName);
      const cleaned = await sanitizeEjoosWorkbookBlob(file);
      downloadBlobFile(downloadName, cleaned);
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

  const looksLikeAnketaFill = (version: BackendEjournalLiveVersion) => {
    const notes = String(version.notes ?? "");
    const protocol = version.changeProtocol as { kind?: string } | null;
    return /анкет/i.test(notes) || /anketa/i.test(String(protocol?.kind ?? ""));
  };

  const restoreStylesFromHistory = async (sourceVersionId?: string) => {
    if (!live?.current) {
      setError("Немає поточної версії ЕЖООС");
      return;
    }
    const current = live.current;
    const versions = live.versions ?? [];
    const source =
      (sourceVersionId
        ? versions.find((version) => version.id === sourceVersionId)
        : null) ||
      [...versions]
        .filter(
          (version) =>
            version.id !== current.id &&
            version.version < current.version &&
            !looksLikeAnketaFill(version),
        )
        .sort((a, b) => b.version - a.version)[0];
    if (!source) {
      setError(
        "У Історії немає версії зі стилями. Завантажте резервну копію .xlsx з нормальним оформленням.",
      );
      return;
    }
    if (
      !window.confirm(
        `Повернути стилі з v${source.version} у поточний файл?\nДані (в тому числі мердж анкет) залишаться. Буде створено нову версію.`,
      )
    ) {
      return;
    }
    setIsLoading(true);
    setError("");
    setMessage(`Беру стилі з v${source.version}…`);
    try {
      const [currentFull, sourceFull] = await Promise.all([
        api.getEjournalLiveFile(current.id, "1ПБ"),
        api.getEjournalLiveFile(source.id, "1ПБ"),
      ]);
      if (!currentFull.fileBase64 || !sourceFull.fileBase64) {
        throw new Error("Не вдалося завантажити файли з БД");
      }
      const currentFile = base64ToFile(
        currentFull.fileBase64,
        currentFull.sourceFileName || `ЕЖООС_v${current.version}.xlsx`,
      );
      const sourceFile = base64ToFile(
        sourceFull.fileBase64,
        sourceFull.sourceFileName || `ЕЖООС_v${source.version}.xlsx`,
      );
      const blob = await graftWorkbookStyles(currentFile, sourceFile);
      const fileName = `ЄЖООС_стилі_з_v${source.version}.xlsx`;
      const fileBase64 = await blobToBase64(blob);
      const saved = await api.applyEjournalLive({
        baseVersionId: current.id,
        fileBase64,
        sourceFileName: fileName,
        asOfDate: current.asOfDate || undefined,
        unitLabel: "1ПБ",
        changeProtocol: {
          kind: "styles-restore",
          fromVersion: source.version,
          protocolText: `Стилі повернуто з v${source.version}. Дані аркушів не змінювались.`,
        },
        notes: `Стилі ← v${source.version}`,
      });
      await refreshLive();
      const localFile = base64ToFile(fileBase64, fileName);
      if (ejoosSnapshot) {
        setEjoosSnapshot({
          ...ejoosSnapshot,
          file: localFile,
          fileName,
        });
      }
      setMessage(
        `Стилі повернуто з v${source.version} · збережено як v${saved.version}. Перевірте файл у Excel.`,
      );
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Не вдалося повернути стилі з історії",
      );
    } finally {
      setIsLoading(false);
    }
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
      const state = await refreshLive();
      const current = state.current ?? saved;
      const full = await api.getEjournalLiveFile(current.id, "1ПБ");
      if (!full.fileBase64) throw new Error("Відкат виконано, але файл не повернувся з БД");
      const file = base64ToFile(
        full.fileBase64,
        full.sourceFileName || `ЕЖООС_v${full.version}.xlsx`,
      );
      const snapshot = await readEjoosWorkbookSnapshot(file);
      assertEjoosWorkbook(snapshot);
      setEjoosSnapshot(snapshot);

      let analysisNote = "";
      let nextTab: EjoosWorkspaceTab = "history";
      if (pbSnapshot) {
        const nextSession = await rebuildSessionFromSnapshots(
          snapshot,
          pbSnapshot,
          state.versions,
        );
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

  const fillSheetFromAnketa = async (
    target: EjoosAnketaFillTarget,
    mode: EjoosAnketaFillMode = "fill",
  ) => {
    const sheetLabel = target === "excluded" ? "Виключені" : "ООС";
    const isMerge = mode === "merge";
    setIsLoading(true);
    setError("");
    setMessage(
      isMerge
        ? `Мердж ${sheetLabel}: завантажую анкетні дані…`
        : "Завантажую анкетні дані…",
    );
    try {
      const ejoos = await ensureEjoosLoaded();
      if (!ejoos) throw new Error("Немає канонічного ЕЖООС у БД");
      if (!live?.current) throw new Error("Немає поточної версії ЕЖООС");

      const { blob, report, fileName } = await fillEjoosSheetFromAnketa({
        ejoos,
        target,
        mode,
        onProgress: (done, total) => {
          setMessage(
            isMerge
              ? `Мердж ${sheetLabel} за ПІБ / ID… ${done}/${total}`
              : `Шукаю пропуски в ${sheetLabel}… ${done}/${total}`,
          );
        },
        onStatus: (text) => setMessage(text),
      });

      const stagingNote = "";

      if (report.fieldCount === 0 && !report.styledCells) {
        setMessage(
          isMerge
            ? `Мердж ${sheetLabel}: немає порожніх або менш повних комірок · ${formatEjoosAnketaFillReport(report)}.`
            : `Пропусків для доповнення в ${sheetLabel} немає · ${formatEjoosAnketaFillReport(report)}.${stagingNote}`,
        );
        return;
      }

      setMessage("Зберігаю нову версію ЕЖООС…");
      const fileBase64 = await blobToBase64(blob);
      const saved = await api.applyEjournalLive({
        baseVersionId: live.current.id,
        fileBase64,
        sourceFileName: fileName,
        asOfDate: live.current.asOfDate || undefined,
        unitLabel: "1ПБ",
        changeProtocol: {
          kind: isMerge ? `anketa-${target}-merge` : `anketa-${target}-fill`,
          report,
          protocolText: isMerge
            ? `Мердж ${sheetLabel} з анкет (колонки пропусків, ПІБ/ID)\n${formatEjoosAnketaFillReport(report)}`
            : `Доповнення ${sheetLabel} з анкетних даних\n${formatEjoosAnketaFillReport(report)}`,
        },
        notes: isMerge
          ? `${sheetLabel} ← мердж анкет · ${formatEjoosAnketaFillReport(report)}`
          : `${sheetLabel} ← анкетні дані · ${formatEjoosAnketaFillReport(report)}`,
      });
      await refreshLive();
      const localFile = base64ToFile(fileBase64, fileName);
      if (ejoosSnapshot) {
        setEjoosSnapshot({
          ...ejoosSnapshot,
          file: localFile,
          fileName,
        });
      }
      setMessage(
        isMerge
          ? `${sheetLabel} змерджено з анкет · v${saved.version} · ${formatEjoosAnketaFillReport(report)}.`
          : `${sheetLabel} доповнено з анкет · v${saved.version} · ${formatEjoosAnketaFillReport(report)}.${stagingNote}`,
      );
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : isMerge
            ? `Не вдалося змерджити ${sheetLabel} з анкетних даних`
            : `Не вдалося доповнити ${sheetLabel} з анкетних даних`,
      );
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
    dismissPerson,
    setDecisions,
    patchOpPayload,
    acceptReady,
    applyAccepted,
    acceptAndApplyPerson,
    downloadCurrentEjoos,
    downloadVersion,
    downloadVersionProtocol,
    rollback,
    restoreStylesFromHistory,
    fillSheetFromAnketa,
  };

  return (
    <EjoosWorkspaceContext.Provider value={value}>
      {children}
    </EjoosWorkspaceContext.Provider>
  );
}
