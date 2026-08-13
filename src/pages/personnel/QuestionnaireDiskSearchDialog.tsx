import { useEffect, useMemo, useRef, useState } from "react";
import { useVirtualizer } from "@tanstack/react-virtual";
import {
  Alert,
  Box,
  Button,
  Checkbox,
  Chip,
  LinearProgress,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  api,
  type DiskQuestionnaireMatchLevel,
  type DiskQuestionnaireSearchPersonResult,
} from "../../api";
import { extractPassportPhotoFromPdf } from "./autoPassportPhoto";
import { FloatingWindow } from "./FloatingWindow";
import { dataUrlToFile } from "./personnelUtils";

type SearchPersonInput = {
  rowId: string;
  externalId: string;
  fullName: string;
  callSign?: string;
  missingQuestionnaire: boolean;
  missingPhoto: boolean;
};

type RowState = DiskQuestionnaireSearchPersonResult & {
  rowId: string;
  selectedPath: string;
  confirmed: boolean;
  confirming: boolean;
  autoStatus?: string;
  autoPhotoStatus?: string;
};

const MATCH_LABEL: Record<DiskQuestionnaireMatchLevel, string> = {
  fio: "ПІБ",
  fi: "ПІ",
  surname: "Прізвище",
  callsign: "Позивний",
};

const normalizeNamePart = (value: string) =>
  String(value ?? "")
    .normalize("NFC")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[''`ʼ´]/g, "")
    .replace(/[^a-zа-яіїєґ0-9\s-]/gi, " ")
    .replace(/-/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const getNormalizedNameTokens = (value: string) => {
  const withoutExtension = String(value ?? "").replace(/\.pdf$/i, "");
  const withoutCallsign = withoutExtension.replace(/\([^)]*\)/g, " ");
  return normalizeNamePart(withoutCallsign).split(" ").filter(Boolean);
};

const parseStrictFio = (value: string) => {
  const tokens = getNormalizedNameTokens(value);
  return {
    surname: tokens[0] ?? "",
    firstName: tokens[1] ?? "",
    patronymic: tokens[2] ?? "",
  };
};

const isExactFioFileNameMatch = (fullName: string, fileName: string) => {
  const person = parseStrictFio(fullName);
  if (!person.surname || !person.firstName || !person.patronymic) return false;

  const fileTokens = getNormalizedNameTokens(fileName);
  const personKey = [person.surname, person.firstName, person.patronymic].join("|");
  for (let index = 0; index <= fileTokens.length - 3; index += 1) {
    const fileKey = fileTokens.slice(index, index + 3).join("|");
    if (fileKey === personKey) return true;
  }

  return false;
};

const normalizeDuplicateFileName = (fileName: string) =>
  String(fileName ?? "")
    .normalize("NFC")
    .trim()
    .toLocaleLowerCase("uk-UA");

const getAutoMatch = (row: RowState) => {
  if (!row.matches.length) return null;
  const fioMatches = row.matches.filter(
    (match) =>
      match.matchLevel === "fio" &&
      isExactFioFileNameMatch(row.fullName, match.fileName),
  );
  if (fioMatches.length === 1) return fioMatches[0];
  if (fioMatches.length <= 1) return null;

  const firstName = normalizeDuplicateFileName(fioMatches[0]?.fileName ?? "");
  const allSameName = fioMatches.every(
    (match) => normalizeDuplicateFileName(match.fileName) === firstName,
  );
  return allSameName ? fioMatches[0] : null;
};

export function QuestionnaireDiskSearchDialog({
  open,
  people,
  onClose,
  onConfirmed,
  onAutoPhotoSaved,
  onPreviewQuestionnaire,
  onCropPhoto,
}: {
  open: boolean;
  people: SearchPersonInput[];
  onClose: () => void;
  onConfirmed: (externalId: string) => void;
  onAutoPhotoSaved: (externalId: string, photoData: string) => void;
  onPreviewQuestionnaire: (file: File, title: string, externalId: string) => void;
  onCropPhoto: (file: File, externalId: string) => void;
}) {
  const [isSearching, setIsSearching] = useState(false);
  const [error, setError] = useState("");
  const [summary, setSummary] = useState("");
  const [rows, setRows] = useState<RowState[]>([]);
  const [filterMatchedOnly, setFilterMatchedOnly] = useState(true);

  const visibleRows = useMemo(
    () =>
      filterMatchedOnly
        ? rows.filter(
            (row) =>
              row.matches.length > 0 ||
              Boolean(row.autoStatus || row.autoPhotoStatus) ||
              (row.missingPhoto && !row.missingQuestionnaire),
          )
        : rows,
    [filterMatchedOnly, rows],
  );

  const listRef = useRef<HTMLDivElement | null>(null);
  const rowVirtualizer = useVirtualizer({
    count: visibleRows.length,
    getScrollElement: () => listRef.current,
    estimateSize: (index) => {
      const matchCount = visibleRows[index]?.matches.length ?? 0;
      return 148 + Math.min(matchCount, 10) * 34;
    },
    overscan: 8,
    gap: 12,
  });

  const patchRow = (rowId: string, patch: Partial<RowState>) => {
    setRows((current) =>
      current.map((item) => (item.rowId === rowId ? { ...item, ...patch } : item)),
    );
  };

  const loadMatchFile = async (relativePath: string, fileName: string) => {
    const blob = await api.getDiskQuestionnaireFile(relativePath);
    return new File([blob], fileName || "questionnaire.pdf", {
      type: "application/pdf",
    });
  };

  const shouldAutoConfirm = (row: RowState) =>
    row.missingQuestionnaire && Boolean(getAutoMatch(row));

  const shouldAutoExtractPhoto = (row: RowState) =>
    row.missingPhoto &&
    (!row.missingQuestionnaire || Boolean(getAutoMatch(row)));

  const loadAutoPhotoFile = async (row: RowState) => {
    if (!row.missingQuestionnaire) {
      const questionnaire = await api.getPersonQuestionnaire(row.externalId);
      if (!questionnaire?.fileData) {
        throw new Error("Анкета є в статусі, але PDF не вдалося прочитати з БД.");
      }
      if (
        questionnaire.fileName &&
        !isExactFioFileNameMatch(row.fullName, questionnaire.fileName)
      ) {
        throw new Error(
          "Автофото пропущено: ПІБ у назві збереженої анкети не збігається з карткою.",
        );
      }
      return dataUrlToFile(
        questionnaire.fileData,
        questionnaire.fileName || "questionnaire.pdf",
      );
    }

    const match = getAutoMatch(row);
    if (!match) throw new Error("Немає PDF для автообробки фото.");
    return loadMatchFile(match.relativePath, match.fileName);
  };

  const autoProcessExactMatches = async (nextRows: RowState[]) => {
    const exactRows = nextRows.filter(
      (row) => shouldAutoConfirm(row) || shouldAutoExtractPhoto(row),
    );
    if (!exactRows.length) return;

    let savedQuestionnaires = 0;
    let savedPhotos = 0;

    for (const row of exactRows) {
      const match = getAutoMatch(row);
      if (!match) continue;
      patchRow(row.rowId, {
        confirming: shouldAutoConfirm(row),
        autoStatus: shouldAutoConfirm(row)
          ? "Автозбереження анкети…"
          : "Анкета вже є. Автообробка фото…",
        autoPhotoStatus: row.missingPhoto ? "Пошук обличчя в PDF…" : undefined,
      });

      if (shouldAutoConfirm(row)) {
        try {
          await api.confirmDiskQuestionnaire(row.externalId, match.relativePath);
          savedQuestionnaires += 1;
          patchRow(row.rowId, {
            confirmed: true,
            confirming: false,
            missingQuestionnaire: false,
            autoStatus:
              row.matches.length === 1
                ? "Анкету збережено автоматично: повний ПІБ, один збіг."
                : "Анкету збережено автоматично: повний ПІБ, дублікати з однаковою назвою.",
          });
          onConfirmed(row.externalId);
        } catch (err) {
          patchRow(row.rowId, {
            confirming: false,
            confirmed: false,
            autoStatus:
              err instanceof Error
                ? `Автозбереження анкети не вдалося: ${err.message}`
                : "Автозбереження анкети не вдалося.",
          });
          continue;
        }
      }

      if (!row.missingPhoto) continue;

      try {
        const file = await loadAutoPhotoFile(row);
        const photo = await extractPassportPhotoFromPdf(file);
        if (!photo) {
          patchRow(row.rowId, {
            autoPhotoStatus:
              "Автофото не знайдено. Можна вирізати вручну з preview.",
          });
          continue;
        }

        const savedPhoto = await api.upsertPersonPhoto(row.externalId, {
          photoData: photo.dataUrl,
          fileName: file.name,
          mimeType: "image/jpeg",
          crop: photo.crop,
        });
        savedPhotos += 1;
        onAutoPhotoSaved(row.externalId, savedPhoto?.photoData || photo.dataUrl);
        patchRow(row.rowId, {
          missingPhoto: false,
          autoPhotoStatus:
            "Фото знайдено, обрізано під паспорт і додано до preview.",
        });
      } catch (err) {
        patchRow(row.rowId, {
          autoPhotoStatus:
            err instanceof Error
              ? `Автофото не вдалося: ${err.message}`
              : "Автофото не вдалося. Можна вирізати вручну з preview.",
        });
      }
    }

    setSummary((current) => {
      const suffix = `Авто: анкет ${savedQuestionnaires}, фото ${savedPhotos}.`;
      return current ? `${current} · ${suffix}` : suffix;
    });
  };

  useEffect(() => {
    if (!open) return;
    listRef.current?.scrollTo({ top: 0 });
  }, [open, filterMatchedOnly, isSearching]);

  useEffect(() => {
    if (!open) return;

    let cancelled = false;
    const snapshot = people;
    const run = async () => {
      setIsSearching(true);
      setError("");
      setSummary("");
      setRows([]);
      try {
        const result = await api.searchQuestionnairesOnDisk({
          people: snapshot.map(({ externalId, fullName, callSign, missingQuestionnaire, missingPhoto }) => ({
            externalId,
            fullName,
            callSign,
            missingQuestionnaire,
            missingPhoto,
          })),
          refreshIndex: true,
        });
        if (cancelled) return;
        setSummary(
          `Проскановано ${result.scannedFiles} PDF · збіги у ${result.matchedPeople} з ${result.people.length} осіб`,
        );
        const nextRows = result.people.map((person, index) => ({
            ...person,
            rowId: snapshot[index]?.rowId || `${person.externalId}-${index}`,
            selectedPath: person.matches[0]?.relativePath ?? "",
            confirmed: false,
            confirming: false,
          }));
        setRows(nextRows);
        void autoProcessExactMatches(nextRows);
      } catch (err) {
        if (cancelled) return;
        setError(
          err instanceof Error
            ? err.message
            : "Не вдалося виконати пошук анкет на диску.",
        );
      } finally {
        if (!cancelled) setIsSearching(false);
      }
    };

    void run();
    return () => {
      cancelled = true;
    };
    // Snapshot people only when dialog opens — avoid re-scan on parent re-renders.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const handlePreview = async (row: RowState) => {
    if (!row.selectedPath) return;
    const match = row.matches.find((item) => item.relativePath === row.selectedPath);
    try {
      const file = await loadMatchFile(
        row.selectedPath,
        match?.fileName || "questionnaire.pdf",
      );
      onPreviewQuestionnaire(
        file,
        `${row.fullName} · ${match?.fileName || ""}`,
        row.externalId,
      );
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Не вдалося відкрити PDF з диска.",
      );
    }
  };

  const handleCrop = async (row: RowState) => {
    if (!row.selectedPath) return;
    const match = row.matches.find((item) => item.relativePath === row.selectedPath);
    try {
      const file = await loadMatchFile(
        row.selectedPath,
        match?.fileName || "questionnaire.pdf",
      );
      onCropPhoto(file, row.externalId);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Не вдалося відкрити PDF для вирізання фото.",
      );
    }
  };

  const handleConfirm = async (row: RowState, checked: boolean) => {
    if (!checked) {
      patchRow(row.rowId, { confirmed: false });
      return;
    }
    if (!row.selectedPath || row.confirming) return;

    patchRow(row.rowId, { confirming: true });
    setError("");

    try {
      await api.confirmDiskQuestionnaire(row.externalId, row.selectedPath);
      patchRow(row.rowId, {
        confirmed: true,
        confirming: false,
        missingQuestionnaire: false,
      });
      onConfirmed(row.externalId);
    } catch (err) {
      patchRow(row.rowId, { confirming: false, confirmed: false });
      setError(
        err instanceof Error
          ? err.message
          : "Не вдалося зберегти підтверджену анкету.",
      );
    }
  };

  return (
    <FloatingWindow
      open={open}
      title="Пошук анкет на диску"
      onClose={onClose}
      placement="center"
      defaultWidth={920}
      defaultHeight={720}
      minWidth={520}
      minHeight={420}
      className="questionnaire-disk-search-floating"
      bodyClassName="questionnaire-disk-search-content"
      footer={
        <Button size="small" variant="outlined" onClick={onClose}>
          Закрити
        </Button>
      }
    >
      <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
        Kingston/army_work · ПІБ → ПІ → прізвище; позивний для
        уточнення дублікатів
      </Typography>
      {isSearching && <LinearProgress color="primary" sx={{ mb: 1.5 }} />}
      {summary && (
        <Alert severity="info" variant="outlined" sx={{ mb: 1.5 }}>
          {summary}
        </Alert>
      )}
      {error && (
        <Alert severity="error" variant="outlined" sx={{ mb: 1.5 }}>
          {error}
        </Alert>
      )}
      <div className="questionnaire-disk-filter">
        <Checkbox
          id="questionnaire-disk-filter-matched"
          label="ПОКАЗУВАТИ ЛИШЕ ЗІ ЗНАЙДЕНИМИ АНКЕТАМИ"
          checked={filterMatchedOnly}
          onCheckedChange={(value) => setFilterMatchedOnly(value === true)}
        />
      </div>

      <div className="questionnaire-disk-list" ref={listRef}>
        {!isSearching && visibleRows.length === 0 ? (
          <Typography variant="body2" color="text.secondary">
            Немає записів для показу.
          </Typography>
        ) : (
          <div
            className="questionnaire-disk-list-virtual"
            style={{ height: `${rowVirtualizer.getTotalSize()}px` }}
          >
            {rowVirtualizer.getVirtualItems().map((virtualRow) => {
              const row = visibleRows[virtualRow.index];
              if (!row) return null;

              return (
                <article
                  key={row.rowId}
                  className="questionnaire-disk-row"
                  data-index={virtualRow.index}
                  ref={rowVirtualizer.measureElement}
                  style={{ transform: `translateY(${virtualRow.start}px)` }}
                >
                  <header className="questionnaire-disk-row-head">
                    <Box>
                      <Typography component="strong" variant="subtitle1">
                        {row.fullName}
                      </Typography>
                      <Typography variant="body2" color="text.secondary">
                        ID: {row.externalId}
                        {row.callSign ? ` · позивний ${row.callSign}` : ""}
                      </Typography>
                    </Box>
                    <Stack
                      direction="row"
                      spacing={0.75}
                      style={{ flexWrap: "wrap" }}
                    >
                      {row.missingQuestionnaire && (
                        <Chip size="small" label="Немає анкети" color="warning" />
                      )}
                      {row.missingPhoto && (
                        <Chip size="small" label="Немає фото" color="warning" />
                      )}
                      {row.confirmed && (
                        <Chip size="small" label="Підтверджено" color="success" />
                      )}
                    </Stack>
                  </header>

                  {row.matches.length === 0 ? (
                    <Typography variant="body2" color="text.secondary">
                      Збігів не знайдено
                    </Typography>
                  ) : (
                    <ul className="questionnaire-disk-matches">
                      {row.matches.map((match, matchIndex) => {
                        const isSelected =
                          row.selectedPath === match.relativePath;
                        const matchId = `match-${row.rowId}-${matchIndex}`;
                        return (
                          <li
                            key={`${row.rowId}-${match.relativePath}-${matchIndex}`}
                          >
                            <div className="questionnaire-disk-match-row">
                              <Checkbox
                                id={matchId}
                                checked={isSelected}
                                onCheckedChange={(value) =>
                                  patchRow(row.rowId, {
                                    selectedPath:
                                      value === true
                                        ? match.relativePath
                                        : row.selectedPath ===
                                            match.relativePath
                                          ? ""
                                          : row.selectedPath,
                                    confirmed: false,
                                  })
                                }
                              />
                              <label
                                className="questionnaire-disk-match-meta"
                                htmlFor={matchId}
                              >
                                <span className="questionnaire-disk-match-name">
                                  {match.fileName}
                                </span>
                                <span>
                                  {MATCH_LABEL[match.matchLevel]}
                                  {match.callSign
                                    ? ` · ${match.callSign}`
                                    : ""}
                                </span>
                              </label>
                            </div>
                          </li>
                        );
                      })}
                    </ul>
                  )}

                  {row.autoStatus || row.autoPhotoStatus ? (
                    <Typography variant="body2" color="text.secondary">
                      {[row.autoStatus, row.autoPhotoStatus]
                        .filter(Boolean)
                        .join(" · ")}
                    </Typography>
                  ) : null}

                  <Stack
                    direction="row"
                    spacing={1}
                    style={{ flexWrap: "wrap" }}
                  >
                    <Button
                      size="small"
                      variant="outlined"
                      disabled={!row.selectedPath}
                      onClick={() => void handlePreview(row)}
                    >
                      Preview анкети
                    </Button>
                    <Button
                      size="small"
                      variant="outlined"
                      disabled={!row.selectedPath}
                      onClick={() => void handleCrop(row)}
                    >
                      Preview фото / вирізати
                    </Button>
                    <Checkbox
                      id={`confirm-${row.rowId}`}
                      label={
                        row.confirming
                          ? "ЗБЕРЕЖЕННЯ…"
                          : "ТАК, ЦЕ ВОНА (ЗБЕРЕГТИ АНКЕТУ)"
                      }
                      checked={row.confirmed}
                      disabled={!row.selectedPath || row.confirming}
                      onCheckedChange={(value) =>
                        void handleConfirm(row, value === true)
                      }
                    />
                  </Stack>
                </article>
              );
            })}
          </div>
        )}
      </div>
    </FloatingWindow>
  );
}
