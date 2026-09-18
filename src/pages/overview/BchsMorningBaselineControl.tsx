import { morningUnitMatches } from "./overviewMorningUnit";
import { useEffect, useRef, useState } from "react";
import { Button, Chip, Stack, Typography } from "@/components/sci/SciPrimitives";
import { UploadFileOutlinedIcon } from "@/components/sci/icons";
import {
  extractBchsMorningRotaNumber,
  readManualBchsMorningBaseline,
  saveManualBchsMorningBaseline,
  clearManualBchsMorningBaseline,
  type BchsMorningManualBaseline,
  type MorningComparisonBaselineKind,
} from "./overviewRotaBchsMorningSnapshot";
import { parseMorningComparisonBaselineFile } from "./overviewMorningComparisonBaselineImport";

const formatBaselineLabel = (baseline: BchsMorningManualBaseline) => {
  const dateLabel = baseline.snapshot.date.split("-").reverse().join(".");
  const rota = extractBchsMorningRotaNumber(baseline.snapshot.unitLabel);
  const rotaLabel = rota ? `${rota} рота · ` : "";
  const kindLabel = baseline.baselineKind === "staff" ? "Штатка" : "БЧС";
  const missionCount = baseline.snapshot.people.filter((person) => person.bucket === "mission").length;
  const awolCount = baseline.snapshot.people.filter((person) => person.bucket === "awol").length;
  return `${rotaLabel}${kindLabel}: ${baseline.fileName} · ${dateLabel} · ${baseline.snapshot.people.length} осіб · На виконанні: ${missionCount} · СЗЧ: ${awolCount}`;
};

export function BchsMorningBaselineControl({
  unitLabel,
  onMessage,
  onBaselineChange,
  compact = false,
}: {
  unitLabel?: string;
  onMessage?: (message: string) => void;
  onBaselineChange?: (baseline: BchsMorningManualBaseline | null) => void;
  compact?: boolean;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const loadSeqRef = useRef(0);
  const [baseline, setBaseline] = useState<BchsMorningManualBaseline | null>(null);
  const [isBusy, setIsBusy] = useState(false);

  const applyBaseline = (next: BchsMorningManualBaseline | null) => {
    setBaseline(next);
    onBaselineChange?.(next);
  };

  useEffect(() => {
    const seq = ++loadSeqRef.current;
    const lookupUnit = unitLabel?.trim();
    applyBaseline(null);
    if (!lookupUnit) return;
    void readManualBchsMorningBaseline(lookupUnit).then((entry) => {
      if (seq !== loadSeqRef.current) return;
      applyBaseline(entry);
    });
  }, [unitLabel]);

  const handlePickFile = () => {
    inputRef.current?.click();
  };

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    const seq = ++loadSeqRef.current;
    setIsBusy(true);
    try {
      const parsed = await parseMorningComparisonBaselineFile(file, unitLabel);
      const storageUnit = unitLabel?.trim() || parsed.unitLabel;
      const baselineKind: MorningComparisonBaselineKind =
        parsed.baselineKind ?? "bchs";
      const nextBaseline: BchsMorningManualBaseline = {
        fileName: file.name,
        uploadedAt: new Date().toISOString(),
        baselineKind,
        snapshot: {
          ...parsed,
          unitLabel: storageUnit,
          source: "manual",
          baselineKind,
          sourceLabel: file.name,
        },
      };
      await saveManualBchsMorningBaseline(nextBaseline);
      const verified =
        (await readManualBchsMorningBaseline(storageUnit)) ?? nextBaseline;
      if (seq !== loadSeqRef.current) return;
      applyBaseline(verified);
      const missionCount = verified.snapshot.people.filter(
        (person) => person.bucket === "mission",
      ).length;
      onMessage?.(
        `Базу для порівняння збережено: ${formatBaselineLabel(verified)}` +
          (missionCount ? ` · на виконанні: ${missionCount}` : "") +
          ".",
      );
    } catch (error) {
      onMessage?.(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати файл для порівняння.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const handleClear = async () => {
    const lookupUnit = unitLabel?.trim() || baseline?.snapshot.unitLabel;
    if (!lookupUnit) return;
    const seq = ++loadSeqRef.current;
    setIsBusy(true);
    try {
      await clearManualBchsMorningBaseline(lookupUnit);
      if (seq !== loadSeqRef.current) return;
      applyBaseline(null);
      onMessage?.("Базу для порівняння очищено.");
    } finally {
      setIsBusy(false);
    }
  };

  const baselineMatchesUnit =
    !unitLabel?.trim() ||
    !baseline ||
    morningUnitMatches(unitLabel, baseline.snapshot.unitLabel);

  return (
    <Stack
      direction="row"
      spacing={compact ? 0.75 : 1}
      alignItems="center"
      className={
        compact
          ? "overview-bchs-baseline-control is-compact"
          : "overview-bchs-baseline-control"
      }
    >
      {!compact ? (
        <Typography variant="body2" color="text.secondary">
          База «Зміна статусу»:
        </Typography>
      ) : null}
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        hidden
        onChange={(event) => void handleFileChange(event)}
      />
      {compact ? (
        <button
          type="button"
          className="sci-data-table-export overview-bchs-baseline-upload"
          disabled={isBusy}
          title="Завантажити базу для «Зміна статусу» (попередній БЧС або Штатка)"
          onClick={handlePickFile}
        >
          <UploadFileOutlinedIcon fontSize="inherit" />
          База
        </button>
      ) : (
        <Button
          variant="outlined"
          size="small"
          startIcon={<UploadFileOutlinedIcon />}
          disabled={isBusy}
          onClick={handlePickFile}
        >
          Завантажити базу
        </Button>
      )}
      {baseline ? (
        <Chip
          size="small"
          color={baselineMatchesUnit ? "default" : "warning"}
          label={
            baselineMatchesUnit
              ? formatBaselineLabel(baseline)
              : `${formatBaselineLabel(baseline)} · інша рота у фільтрі`
          }
          onDelete={isBusy ? undefined : () => void handleClear()}
        />
      ) : compact ? (
        <Typography variant="caption" color="text.secondary" noWrap>
          {unitLabel
            ? "немає бази"
            : "оберіть роту у «Підрозділ»"}
        </Typography>
      ) : (
        <Typography variant="caption" color="text.secondary">
          {unitLabel
            ? "немає — завантажте попередній БЧС або Штатку"
            : "завантажте БЧС/Штатку або оберіть роту у фільтрі «Підрозділ»"}
        </Typography>
      )}
    </Stack>
  );
}
