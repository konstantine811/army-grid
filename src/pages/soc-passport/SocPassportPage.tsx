import { useMemo, useState } from "react";
import { Alert, Box, Stack, Typography } from "@/components/sci/SciPrimitives";
import {
  CloudUploadOutlinedIcon,
  FileDownloadOutlinedIcon,
  TableChartOutlinedIcon,
} from "@/components/sci/icons";
import { Button as SciButton } from "../../components/ui/button/button";
import { readWorkbookSnapshot, type ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { buildSocPassportResult } from "./socPassportCalc";
import {
  exportSocPassportExitsWorkbook,
  exportSocPassportNoExitsWorkbook,
  exportSocPassportRussiaWorkbook,
  exportSocPassportWorkbook,
} from "./socPassportExport";
import { exportSocPassportDeparturesWorkbook } from "./socPassportDeparturesExport";
import {
  buildArrivalsFromPb,
  buildCombatLossesFromPb,
  buildDeparturesFromEjoos,
  buildDispositionFromArchive,
  buildSzchFromRuh,
  withMorningArrivalSources,
  type ArrivalsMonthResult,
  type CombatLossesResult,
  type DeparturesResult,
  type DispositionArchiveResult,
  type SzchRuhResult,
} from "./socPassportDepartures";
import { assertEjoosWorkbook, assertPbWorkbook } from "../ejournal/ejoosWorkbookKind";
import {
  buildHousingIdpStats,
  housingIdpRemainingNameSet,
  type HousingIdpStatsResult,
} from "./housingIdpStats";
import {
  countsInNoExitsList,
  normalizePersonName,
} from "./socPassportFields";
import { parseSocPassportSources } from "./socPassportParse";
import {
  STATIC_COMBAT_EXIT_OVERRIDE_SOURCE,
  STATIC_COMBAT_EXIT_OVERRIDES,
  findStaticCombatExitOverride,
} from "./staticCombatExitOverrides";
import type { PassportTableRow, SocPassportResult } from "./socPassportTypes";
import {
  parseUbdRegistryStats,
  type UbdRegistryStatsResult,
} from "./ubdRegistryStats";

const formatFileLabel = (fileName?: string) => fileName || "не завантажено";

const shortFileName = (name?: string) => {
  if (!name) return "";
  return name.length > 36 ? `${name.slice(0, 34)}…` : name;
};

const FileButton = ({
  label,
  loadedName,
  onFile,
  disabled,
}: {
  label: string;
  loadedName?: string;
  onFile: (file: File) => void;
  disabled?: boolean;
}) => (
  <SciButton asChild variant="OUTLINE" disabled={disabled}>
    <label>
      <CloudUploadOutlinedIcon fontSize="small" />
      {loadedName ? `${label}: ${shortFileName(loadedName)}` : label}
      <input
        hidden
        type="file"
        accept=".xlsx,.xlsm,.xls"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onFile(file);
          event.target.value = "";
        }}
      />
    </label>
  </SciButton>
);

const CountsCells = ({ row }: { row: PassportTableRow }) => {
  if (row.kind === "section" || !row.counts) {
    return (
      <>
        <td colSpan={7} />
      </>
    );
  }
  return (
    <>
      <td>{row.counts.officerMobilized || ""}</td>
      <td>{row.counts.officerContract || ""}</td>
      <td>{row.counts.sergeantMobilized || ""}</td>
      <td>{row.counts.sergeantContract || ""}</td>
      <td>{row.counts.soldierMobilized || ""}</td>
      <td>{row.counts.soldierContract || ""}</td>
      <td>{row.total || ""}</td>
    </>
  );
};

const ArrivalsPreview = ({
  arrivals,
  titleHint,
  headingPrefix = "Прибули",
}: {
  arrivals: ArrivalsMonthResult;
  titleHint: string;
  headingPrefix?: string;
}) => (
  <>
    <div className="panel-heading" style={{ marginTop: 16 }}>
      {headingPrefix} · {arrivals.monthLabel} · {arrivals.sourceSheet}
    </div>
    <Typography variant="body2" sx={{ mb: 1.5, lineHeight: 1.7 }}>
      {titleHint}
    </Typography>
    <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
      <table className="bchs-analytics-table">
        <thead>
          <tr>
            <th>№</th>
            <th>За джерелом прибуття</th>
            <th>Офіцери</th>
            <th>Сержанти</th>
            <th>Солдати</th>
            <th>Разом</th>
          </tr>
        </thead>
        <tbody>
          {arrivals.sourceSummary.map((row, index) => (
            <tr key={row.source}>
              <td>{index + 1}</td>
              <td>{row.label}</td>
              <td>{row.byRank.officer || ""}</td>
              <td>{row.byRank.sergeant || ""}</td>
              <td>{row.byRank.soldier || ""}</td>
              <td>
                <strong>{row.count}</strong>
              </td>
            </tr>
          ))}
          <tr>
            <td />
            <td>
              <strong>Разом</strong>
            </td>
            <td>
              <strong>{arrivals.byRank.officer || ""}</strong>
            </td>
            <td>
              <strong>{arrivals.byRank.sergeant || ""}</strong>
            </td>
            <td>
              <strong>{arrivals.byRank.soldier || ""}</strong>
            </td>
            <td>
              <strong>{arrivals.total}</strong>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </>
);

const DispositionPreview = ({
  disposition,
}: {
  disposition: DispositionArchiveResult;
}) => (
  <>
    <div className="panel-heading" style={{ marginTop: 16 }}>
      Виведені у розпорядження · {disposition.sourceSheet}
      {disposition.periodFrom ? ` · з ${disposition.periodFrom}` : ""}
    </div>
    <Typography variant="body2" sx={{ mb: 1.5, lineHeight: 1.7 }}>
      З 1ПБ archive: рядки з маркером розпорядження, відкриті періоди
      {disposition.periodFromLabel ? ` ${disposition.periodFromLabel}` : ""}.
      У archive всього {disposition.totalArchiveRows}
      {disposition.skippedNotDisposition ||
      disposition.skippedBeforePeriod ||
      disposition.skippedReturned
        ? ` (не розпорядження ${disposition.skippedNotDisposition}, до періоду ${disposition.skippedBeforePeriod}, повернулись ${disposition.skippedReturned})`
        : ""}
      .
    </Typography>
    <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
      <table className="bchs-analytics-table">
        <thead>
          <tr>
            <th>№</th>
            <th>Підстава</th>
            <th>Офіцери</th>
            <th>Сержанти</th>
            <th>Солдати</th>
            <th>Разом</th>
          </tr>
        </thead>
        <tbody>
          {disposition.summary.map((row, index) => (
            <tr key={row.reason}>
              <td>{index + 1}</td>
              <td>{row.label}</td>
              <td>{row.byRank.officer || ""}</td>
              <td>{row.byRank.sergeant || ""}</td>
              <td>{row.byRank.soldier || ""}</td>
              <td>
                <strong>{row.count}</strong>
              </td>
            </tr>
          ))}
          <tr>
            <td />
            <td>
              <strong>Разом у розпорядженні</strong>
            </td>
            <td>
              <strong>{disposition.totals.byRank.officer}</strong>
            </td>
            <td>
              <strong>{disposition.totals.byRank.sergeant}</strong>
            </td>
            <td>
              <strong>{disposition.totals.byRank.soldier}</strong>
            </td>
            <td>
              <strong>{disposition.totals.all}</strong>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </>
);

const SzchPreview = ({ szch }: { szch: SzchRuhResult }) => (
  <>
    <div className="panel-heading" style={{ marginTop: 16 }}>
      СЗЧ · {szch.sourceSheet}
      {szch.periodFrom ? ` · з ${szch.periodFrom}` : ""}
    </div>
    <Typography variant="body2" sx={{ mb: 1.5, lineHeight: 1.7 }}>
      З 1ПБ Рух: рядки з маркером СЗЧ / самовільне залишення
      {szch.periodFromLabel ? ` ${szch.periodFromLabel}` : ""}. У Рух всього{" "}
      {szch.totalMovements}
      {szch.skippedBeforePeriod || szch.skippedNoDate
        ? ` (до періоду ${szch.skippedBeforePeriod}, без дати ${szch.skippedNoDate})`
        : ""}
      .
    </Typography>
    <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
      <table className="bchs-analytics-table">
        <thead>
          <tr>
            <th>№</th>
            <th>Категорія</th>
            <th>Кількість</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>1</td>
            <td>Офіцери</td>
            <td>
              <strong>{szch.totals.byRank.officer}</strong>
            </td>
          </tr>
          <tr>
            <td>2</td>
            <td>Сержанти</td>
            <td>
              <strong>{szch.totals.byRank.sergeant}</strong>
            </td>
          </tr>
          <tr>
            <td>3</td>
            <td>Солдати</td>
            <td>
              <strong>{szch.totals.byRank.soldier}</strong>
            </td>
          </tr>
          <tr>
            <td />
            <td>
              <strong>Разом СЗЧ</strong>
            </td>
            <td>
              <strong>{szch.totals.all}</strong>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </>
);

const CombatLossesPreview = ({ losses }: { losses: CombatLossesResult }) => (
  <>
    <div className="panel-heading" style={{ marginTop: 16 }}>
      Втрати · {losses.sourceSheet}
      {losses.periodFrom ? ` · з ${losses.periodFrom}` : ""}
    </div>
    <Typography variant="body2" sx={{ mb: 1.5, lineHeight: 1.7 }}>
      З 1ПБ Рух + archive: загиблі / безвісти / у бою / інші обставини
      {losses.periodFromLabel ? ` ${losses.periodFromLabel}` : ""}.
      {losses.skippedBeforePeriod || losses.skippedNoDate
        ? ` Відфільтровано: до періоду ${losses.skippedBeforePeriod}, без дати ${losses.skippedNoDate}.`
        : ""}
    </Typography>
    <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
      <table className="bchs-analytics-table">
        <thead>
          <tr>
            <th>№</th>
            <th>Категорія</th>
            <th>Офіцери</th>
            <th>Сержанти</th>
            <th>Солдати</th>
            <th>Разом</th>
          </tr>
        </thead>
        <tbody>
          {losses.summary.map((row, index) => (
            <tr key={row.reason}>
              <td>{index + 1}</td>
              <td>{row.label}</td>
              <td>{row.byRank.officer || ""}</td>
              <td>{row.byRank.sergeant || ""}</td>
              <td>{row.byRank.soldier || ""}</td>
              <td>
                <strong>{row.count}</strong>
              </td>
            </tr>
          ))}
          <tr>
            <td />
            <td>
              <strong>Разом втрати</strong>
            </td>
            <td>
              <strong>{losses.totals.byRank.officer}</strong>
            </td>
            <td>
              <strong>{losses.totals.byRank.sergeant}</strong>
            </td>
            <td>
              <strong>{losses.totals.byRank.soldier}</strong>
            </td>
            <td>
              <strong>{losses.totals.all}</strong>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </>
);

export function SocPassportPage() {
  const [ejoos, setEjoos] = useState<ExcelWorkbookSnapshot | null>(null);
  const [pb, setPb] = useState<ExcelWorkbookSnapshot | null>(null);
  const [morning, setMorning] = useState<ExcelWorkbookSnapshot | null>(null);
  const [jbdExits, setJbdExits] = useState<ExcelWorkbookSnapshot | null>(null);
  const [bplaExits, setBplaExits] = useState<ExcelWorkbookSnapshot | null>(null);
  const [ubdRoster, setUbdRoster] = useState<ExcelWorkbookSnapshot | null>(null);
  const [housingIdp, setHousingIdp] = useState<ExcelWorkbookSnapshot | null>(
    null,
  );
  const [housingIdpStats, setHousingIdpStats] =
    useState<HousingIdpStatsResult | null>(null);
  const [result, setResult] = useState<SocPassportResult | null>(null);
  const [departures, setDepartures] = useState<DeparturesResult | null>(null);
  const [pbArrivals, setPbArrivals] = useState<ArrivalsMonthResult | null>(null);
  const [pbDisposition, setPbDisposition] =
    useState<DispositionArchiveResult | null>(null);
  const [pbSzch, setPbSzch] = useState<SzchRuhResult | null>(null);
  const [pbCombatLosses, setPbCombatLosses] =
    useState<CombatLossesResult | null>(null);
  const [ubdStats, setUbdStats] = useState<UbdRegistryStatsResult | null>(null);
  const [message, setMessage] = useState(
    "Завантажте ЕЖООС (ШПО + ООС) і за бажанням ранковий звіт — потім натисніть «Порахувати».",
  );
  const [isBusy, setIsBusy] = useState(false);
  const [query, setQuery] = useState("");
  const [ubdYearFilter, setUbdYearFilter] = useState<number | "unknown" | null>(
    null,
  );
  const [batchListSheet, setBatchListSheet] = useState<string | null>(null);
  const filteredPeople = useMemo(() => {
    if (!result) return [];
    const needle = query.trim().toLocaleLowerCase("uk-UA");
    if (!needle) return result.people;
    return result.people.filter((person) =>
      `${person.name} ${person.position} ${person.regionLabel} ${person.relativesRaw}`
        .toLocaleLowerCase("uk-UA")
        .includes(needle),
    );
  }, [query, result]);

  const noExitsPeople = useMemo(() => {
    if (!result) return [];
    return result.people
      .filter(
        (person) =>
          person.exitBand === "none" &&
          countsInNoExitsList(person) &&
          !person.ubdRosterStatus &&
          !person.staticCombatExitOverride,
      )
      .sort((left, right) => left.name.localeCompare(right.name, "uk"));
  }, [result]);

  const staticCombatExitRows = useMemo(() => {
    const people = result?.people ?? [];
    return STATIC_COMBAT_EXIT_OVERRIDES.map((row) => {
      const matched = people.find(
        (person) => findStaticCombatExitOverride(person.name)?.name === row.name,
      );
      return {
        ...row,
        inMorning: Boolean(matched),
        exitCount: matched?.exitCount ?? null,
        morningStatus: matched?.morningStatus ?? "",
      };
    });
  }, [result]);

  const staticCombatMatchedCount = staticCombatExitRows.filter(
    (row) => row.inMorning,
  ).length;

  const loadWorkbook = async (
    file: File | undefined,
    kind: "ejoos" | "pb" | "morning" | "jbd" | "bpla" | "ubd" | "housing",
  ) => {
    if (!file) return;
    setIsBusy(true);
    try {
      const snapshot = await readWorkbookSnapshot(file);
      if (kind === "ejoos") {
        assertEjoosWorkbook(snapshot);
        setEjoos(snapshot);
        try {
          const nextDepartures = buildDeparturesFromEjoos(snapshot);
          let nextPbArrivals = pbArrivals;
          if (pb) {
            try {
              nextPbArrivals = buildArrivalsFromPb(
                pb,
                8,
                new Date().getFullYear(),
                { ejoos: snapshot },
              );
              setPbArrivals(nextPbArrivals);
            } catch {
              /* лишаємо попередній розрахунок 1ПБ */
            }
          }
          const withPb = {
            ...nextDepartures,
            ...(nextPbArrivals ? { arrivalsAugustPb: nextPbArrivals } : {}),
            ...(pbDisposition
              ? { dispositionFromArchive: pbDisposition }
              : {}),
            ...(pbSzch ? { szchFromRuh: pbSzch } : {}),
            ...(pbCombatLosses
              ? { combatLossesFromPb: pbCombatLosses }
              : {}),
          };
          const withMorning = morning
            ? withMorningArrivalSources(withPb, snapshot, morning)
            : withPb;
          setDepartures(withMorning);
          const arrivals = withMorning.arrivalsAugust;
          const arrivalsPart = arrivals
            ? ` Прибули ООС (${arrivals.monthLabel}): ${arrivals.total} · ТЦК ${arrivals.bySource.tck} / НЦ ${arrivals.bySource.trainingCenter} / перевед. ${arrivals.bySource.unitTransfer}.`
            : "";
          const morningArrivals = withMorning.arrivalsFromMorning;
          const morningPart = morningArrivals
            ? ` Звідки (Штатка): ${morningArrivals.total} · БРЕЗ ${morningArrivals.bySource.brez} / ТЦК ${morningArrivals.bySource.tck} / перевед. ${morningArrivals.bySource.unitTransfer}.`
            : morning
              ? ""
              : " (для «Звідки» завантажте Штатку / ранковий).";
          const pbPart = nextPbArrivals
            ? ` Прибули 1ПБ: ${nextPbArrivals.total} · ТЦК ${nextPbArrivals.bySource.tck} / НЦ ${nextPbArrivals.bySource.trainingCenter} / перевед. ${nextPbArrivals.bySource.unitTransfer}.`
            : "";
          setMessage(
            `ЕЖООС: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}. «Вибули» з «${withMorning.sourceSheet}»${withMorning.periodFrom ? ` з ${withMorning.periodFrom}` : ""}: ${withMorning.totals.all} з ${withMorning.totalUnfiltered} (звільнення ${withMorning.totals.discharges}, переведені ${withMorning.totals.transfers}).${arrivalsPart}${morningPart}${pbPart}`,
          );
        } catch (departuresError) {
          setDepartures(null);
          setMessage(
            `ЕЖООС: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}. ${
              departuresError instanceof Error
                ? departuresError.message
                : "Не вдалося порахувати «Вибули»."
            }`,
          );
        }
      } else if (kind === "pb") {
        assertPbWorkbook(snapshot);
        setPb(snapshot);
        const month = 8;
        const year = new Date().getFullYear();
        try {
          const nextPbArrivals = buildArrivalsFromPb(snapshot, month, year, {
            ejoos,
          });
          let nextDisposition: DispositionArchiveResult | null = null;
          let nextSzch: SzchRuhResult | null = null;
          let nextCombatLosses: CombatLossesResult | null = null;
          try {
            nextDisposition = buildDispositionFromArchive(snapshot);
          } catch {
            nextDisposition = null;
          }
          try {
            nextSzch = buildSzchFromRuh(snapshot);
          } catch {
            nextSzch = null;
          }
          try {
            nextCombatLosses = buildCombatLossesFromPb(snapshot);
          } catch {
            nextCombatLosses = null;
          }
          setPbArrivals(nextPbArrivals);
          setPbDisposition(nextDisposition);
          setPbSzch(nextSzch);
          setPbCombatLosses(nextCombatLosses);
          setDepartures((current) =>
            current
              ? {
                  ...current,
                  arrivalsAugustPb: nextPbArrivals,
                  dispositionFromArchive: nextDisposition,
                  szchFromRuh: nextSzch,
                  combatLossesFromPb: nextCombatLosses,
                }
              : current,
          );
          const dispPart = nextDisposition
            ? ` Розпорядження (archive): ${nextDisposition.totals.all} (лік. ${nextDisposition.summary[0]?.count ?? 0}, орг. ${nextDisposition.summary[1]?.count ?? 0}, інше ${nextDisposition.summary[2]?.count ?? 0}).`
            : "";
          const szchPart = nextSzch
            ? ` СЗЧ (Рух): ${nextSzch.totals.all} (оф. ${nextSzch.totals.byRank.officer}, серж. ${nextSzch.totals.byRank.sergeant}, солд. ${nextSzch.totals.byRank.soldier}).`
            : "";
          const lossesPart = nextCombatLosses
            ? ` Втрати: ${nextCombatLosses.totals.all} (загиблі ${nextCombatLosses.summary[0]?.count ?? 0}, безвісти ${nextCombatLosses.summary[1]?.count ?? 0}, у бою ${nextCombatLosses.summary[2]?.count ?? 0}, інше ${nextCombatLosses.summary[3]?.count ?? 0}).`
            : "";
          setMessage(
            `1ПБ: ${snapshot.fileName} · sh / Рух / archive. Прибули (${nextPbArrivals.monthLabel}): ${nextPbArrivals.total} (оф. ${nextPbArrivals.byRank.officer}, серж. ${nextPbArrivals.byRank.sergeant}, солд. ${nextPbArrivals.byRank.soldier}). Джерела з «Звідки прибув»: ТЦК ${nextPbArrivals.bySource.tck}, НЦ ${nextPbArrivals.bySource.trainingCenter}, переведені ${nextPbArrivals.bySource.unitTransfer}.${dispPart}${szchPart}${lossesPart}`,
          );
        } catch (pbError) {
          setPbArrivals(null);
          setPbDisposition(null);
          setPbSzch(null);
          setPbCombatLosses(null);
          setMessage(
            `1ПБ: ${snapshot.fileName}. ${
              pbError instanceof Error
                ? pbError.message
                : "Не вдалося порахувати прибули з Рух."
            }`,
          );
        }
      } else if (kind === "morning") {
        setMorning(snapshot);
        let housingPart = "";
        if (housingIdp) {
          try {
            const nextHousing = buildHousingIdpStats(housingIdp, snapshot);
            setHousingIdpStats(nextHousing);
            housingPart = ` ВПО житло: у файлі ${nextHousing.totals.inFile}, лишилось у ранковому ${nextHousing.totals.remaining} (у наявності ${nextHousing.totals.remainingPresent}), вибули/немає ${nextHousing.totals.leftNotInMorning}.`;
          } catch (housingError) {
            setHousingIdpStats(null);
            housingPart = ` ВПО житло: ${
              housingError instanceof Error
                ? housingError.message
                : "не вдалося зіставити."
            }`;
          }
        }
        if (ejoos && departures) {
          const withMorning = withMorningArrivalSources(
            departures,
            ejoos,
            snapshot,
          );
          setDepartures(withMorning);
          const m = withMorning.arrivalsFromMorning;
          setMessage(
            m
              ? `Ранковий (Штатка): ${snapshot.fileName}. Звідки: ${m.total} · БРЕЗ ${m.bySource.brez} / ТЦК ${m.bySource.tck} / перевед. ${m.bySource.unitTransfer} (оф. ${m.byRank.officer}, серж. ${m.byRank.sergeant}, солд. ${m.byRank.soldier}).${housingPart}`
              : `Ранковий звіт: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.${housingPart}`,
          );
        } else {
          setMessage(
            `Ранковий звіт: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.${
              ejoos
                ? ""
                : " Завантажте також ЕЖООС — тоді порахуємо «Звідки» (БРЕЗ / ТЦК / в/ч)."
            }${housingPart}`,
          );
        }
      } else if (kind === "jbd") {
        setJbdExits(snapshot);
        setMessage(
          `ЖБД виходи: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.`,
        );
      } else if (kind === "bpla") {
        setBplaExits(snapshot);
        setMessage(
          `БПЛА виходи: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.`,
        );
      } else if (kind === "housing") {
        setHousingIdp(snapshot);
        if (morning) {
          try {
            const nextHousing = buildHousingIdpStats(snapshot, morning);
            setHousingIdpStats(nextHousing);
            setMessage(
              `ВПО (житло): ${snapshot.fileName} · у списку ${nextHousing.totals.inFile}. Лишилось у ранковому: ${nextHousing.totals.remaining} (оф. ${nextHousing.totals.byRankRemaining.officer}, серж. ${nextHousing.totals.byRankRemaining.sergeant}, солд. ${nextHousing.totals.byRankRemaining.soldier}); у наявності ${nextHousing.totals.remainingPresent}; немає в ранковому ${nextHousing.totals.leftNotInMorning}.`,
            );
            console.group(
              `[ВПО / житло] ${nextHousing.fileName} · лишилось ${nextHousing.totals.remaining} з ${nextHousing.totals.inFile}`,
            );
            console.log("Лишились у ранковому:", nextHousing.remainingNames);
            console.log("Немає в ранковому:", nextHousing.leftNames);
            console.groupEnd();
          } catch (housingError) {
            setHousingIdpStats(null);
            setMessage(
              housingError instanceof Error
                ? housingError.message
                : "Не вдалося зіставити ВПО з ранковим.",
            );
          }
        } else {
          setHousingIdpStats(null);
          setMessage(
            `ВПО (житло): ${snapshot.fileName}. Завантажте ранковий звіт — тоді перевіримо хто лишився.`,
          );
        }
      } else {
        setUbdRoster(snapshot);
        setUbdStats(null);
        setUbdYearFilter(null);
        setBatchListSheet(null);
        setMessage(
          `УБД реєстр: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}. Натисніть «Підрахунок УБД».`,
        );
      }
      setResult(null);
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося прочитати Excel.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const calculate = () => {
    if (!ejoos || !morning) {
      setMessage("Потрібні обидва файли: ЕЖООС (анкети ООС) і ранковий звіт (як у БЧС).");
      return;
    }
    setIsBusy(true);
    try {
      const parsed = parseSocPassportSources({ ejoos, morning, jbdExits, bplaExits, ubdRoster });
      let nextHousingStats: HousingIdpStatsResult | null = null;
      if (housingIdp) {
        try {
          nextHousingStats = buildHousingIdpStats(housingIdp, morning);
          setHousingIdpStats(nextHousingStats);
        } catch (housingError) {
          setHousingIdpStats(null);
          console.warn("[ВПО / житло]", housingError);
        }
      }
      const idpRemaining = housingIdpRemainingNameSet(nextHousingStats);
      const peopleWithHousingIdp = parsed.people.map((person) => {
        const fromHousing =
          idpRemaining.has(normalizePersonName(person.name)) ||
          idpRemaining.has(person.shortName);
        if (!fromHousing) return person;
        return {
          ...person,
          isIdp: true,
          parseNotes: person.parseNotes.includes("ВПО: додаток про житло")
            ? person.parseNotes
            : [...person.parseNotes, "ВПО: додаток про житло"],
        };
      });
      const warnings = [
        morning
          ? ""
          : "Ранковий звіт не завантажено: «В наявності», «В розпорядженні» і виходи на ЛБЗ будуть порожні або нульові.",
        !jbdExits
          ? "ЖБД не завантажено: виходи рахуються лише з «Статус бійців» ранкового звіту."
          : "",
        !bplaExits
          ? "БПЛА не завантажено: екіпажі БПЛА з «Виконував БЗ» не враховуються окремо."
          : "",
        !ubdRoster
          ? "УБД реєстр не завантажено: «Подавалися» / «Не подавалися» не враховуються для бойових виходів."
          : "",
        housingIdp && !nextHousingStats
          ? "Додаток про житло (ВПО) не вдалося зіставити з ранковим."
          : "",
        peopleWithHousingIdp.length === 0
          ? "На ШПО не знайдено осіб на посадах."
          : "",
      ].filter(Boolean);
      const next = buildSocPassportResult({
        people: peopleWithHousingIdp,
        staffSlots: parsed.staffSlots,
        sheets: {
          ...parsed.sheets,
          housingIdp: housingIdp?.fileName ?? "",
        },
        warnings,
      });
      setResult(next);

      if (departures) {
        setDepartures(withMorningArrivalSources(departures, ejoos, morning));
      } else {
        try {
          const nextDepartures = withMorningArrivalSources(
            buildDeparturesFromEjoos(ejoos),
            ejoos,
            morning,
          );
          setDepartures(nextDepartures);
        } catch {
          /* вибули вже могли бути пораховані при імпорті ЕЖООС */
        }
      }

      const brez = parsed.brezDebug;
      console.groupCollapsed(
        `[Соц.паспорт] БРЕЗ: ООС ${brez.counts.oosBrezAll} → ранок ${brez.counts.oosBrezKeptInMorning} (відкинуто ${brez.counts.oosBrezDroppedByMorning}); у підрахунку ${brez.counts.countedBrezAll} (з ООС ${brez.counts.countedBrezFromOos} + лише ранок ${brez.counts.countedBrezFromMorningOnly})`,
      );
      console.log("1) Усі БРЕЗ з ООС (Звідки прибув містить БРЕЗ):", brez.oosBrezAll);
      console.log("2) З них є в ранковому (нова + звання):", brez.oosBrezKeptInMorning);
      console.log("3) Відфільтровано ранковим (немає в списку):", brez.oosBrezDroppedByMorning);
      console.log("4) У підрахунку «З Брез» з ООС:", brez.countedBrezFromOos);
      console.log("5) У підрахунку «З Брез» лише з ранку (БЗВП/БРЕЗ або Відрядження БРЕЗ):", brez.countedBrezFromMorningOnly);
      console.log("6) Увесь підрахунок «З Брез»:", brez.countedBrezAll);
      console.log("counts:", brez.counts);
      console.groupEnd();
      (window as Window & { __SOC_PASSPORT_BREZ__?: unknown }).__SOC_PASSPORT_BREZ__ = brez;

      const ubd = parsed.ubdDebug;
      console.groupCollapsed(
        `[Соц.паспорт] УБД: ООС ${ubd.counts.oosUbdAll} → ранок ${ubd.counts.oosUbdKeptInMorning} (відкинуто ${ubd.counts.oosUbdDroppedByMorning}); у підрахунку ${ubd.counts.countedUbdInMorning}`,
      );
      console.log(
        "1) Усі УБД з ООС:",
        ubd.oosUbdAll.map((row) => row.name),
      );
      console.log(
        "2) Після фільтра ранкового:",
        ubd.oosUbdKeptInMorning.map((row) => row.name),
      );
      console.log(
        "3) Відкинуті ранковим:",
        ubd.oosUbdDroppedByMorning.map((row) => row.name),
      );
      console.log("counts:", ubd.counts);
      console.groupEnd();
      (window as Window & { __SOC_PASSPORT_UBD__?: unknown }).__SOC_PASSPORT_UBD__ =
        ubd;

      const combat = parsed.combatDutyDebug;
      console.groupCollapsed(
        `[Соц.паспорт] Бойові завдання: тимч. прибуття ${combat.counts.tempArrivalCombatZone}; підвищено з 0 виходів ${combat.counts.inferredExitFloor}`,
      );
      console.log(
        "1) Усі «Тимчасово прибулі» у зоні бойових завдань:",
        combat.tempArrivalAll.map((row) => row.name),
      );
      console.log(
        "2) У підрахунку виходів без «Статус бійців» (УБД / дислокація / тимч. прибуття):",
        combat.inferredFromEjoosOnly,
      );
      console.log("counts:", combat.counts);
      console.groupEnd();
      (window as Window & { __SOC_PASSPORT_COMBAT__?: unknown }).__SOC_PASSPORT_COMBAT__ =
        combat;

      const jbd = parsed.jbdExitsDebug;
      console.groupCollapsed(
        `[Соц.паспорт] ЖБД: ${jbd.counts.jbdStampsTotal} виходів у файлі; +${jbd.counts.peopleWithJbdMerged} осіб (нових без дубліката: ${jbd.counts.peopleWithJbdOnly})`,
      );
      console.log("1) Додано до підрахунку (без дублікатів з «Статус бійців»):", jbd.addedToMorning);
      console.log("counts:", jbd.counts);
      console.groupEnd();
      (window as Window & { __SOC_PASSPORT_JBD__?: unknown }).__SOC_PASSPORT_JBD__ = jbd;

      const bpla = parsed.bplaExitsDebug;
      console.groupCollapsed(
        `[Соц.паспорт] БПЛА: ${bpla.counts.performersInFile} «Виконував БЗ» у файлі; зіставлено ${bpla.counts.matchedPerformers} (піднято з 0 виходів: ${bpla.counts.liftedFromNoExits})`,
      );
      console.log("1) Зіставлено з ранковим:", bpla.matchedInMorning);
      console.log("2) У файлі, але немає в ранковому:", bpla.unmatchedInFile);
      console.log("counts:", bpla.counts);
      console.groupEnd();
      (window as Window & { __SOC_PASSPORT_BPLA__?: unknown }).__SOC_PASSPORT_BPLA__ = bpla;

      const ubdReg = parsed.ubdRosterDebug;
      console.groupCollapsed(
        `[Соц.паспорт] УБД реєстр: подали ${ubdReg.counts.submittedInFile}, не подали ${ubdReg.counts.notSubmittedInFile}; у ранковому ${ubdReg.counts.matchedSubmitted + ubdReg.counts.matchedNotSubmitted}`,
      );
      console.log("1) Подавалися у файлі:", ubdReg.counts.submittedInFile);
      console.log("2) Не подавалися у файлі:", ubdReg.counts.notSubmittedInFile);
      console.log("3) Зіставлено з ранковим:", ubdReg.matchedInMorning);
      console.log("counts:", ubdReg.counts);
      console.groupEnd();
      (window as Window & { __SOC_PASSPORT_UBD_ROSTER__?: unknown }).__SOC_PASSPORT_UBD_ROSTER__ =
        ubdReg;

      setMessage(
        `Пораховано: штат ${next.summary.staffSlots}, на посаді ${next.summary.occupied}, ООС ${next.summary.oosMatched}, ранок ${next.summary.morningMatched}. БРЕЗ: ${brez.counts.oosBrezAll}→${brez.counts.countedBrezAll}; УБД: ${ubd.counts.oosUbdAll}→${ubd.counts.countedUbdInMorning}; бойове (ЕЖООС): ${combat.counts.inferredExitFloor}; ЖБД: ${jbd.counts.peopleWithJbdMerged}; БПЛА: ${bpla.counts.matchedPerformers}; УБД реєстр: ${ubdReg.counts.matchedSubmitted + ubdReg.counts.matchedNotSubmitted}${
          nextHousingStats
            ? `; ВПО житло: лишилось ${nextHousingStats.totals.remaining} з ${nextHousingStats.totals.inFile}`
            : ""
        } (див. console).`,
      );
      if (nextHousingStats) {
        console.group(
          `[ВПО / житло] лишилось ${nextHousingStats.totals.remaining} з ${nextHousingStats.totals.inFile}`,
        );
        console.log("Лишились:", nextHousingStats.remainingNames);
        console.log("Немає в ранковому:", nextHousingStats.leftNames);
        console.groupEnd();
        (
          window as Window & { __SOC_PASSPORT_HOUSING_IDP__?: unknown }
        ).__SOC_PASSPORT_HOUSING_IDP__ = nextHousingStats;
      }
    } catch (error) {
      setResult(null);
      setMessage(
        error instanceof Error ? error.message : "Не вдалося порахувати соц.паспорт.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const calculateUbdStats = async () => {
    if (!ubdRoster) {
      setMessage("Спочатку завантажте файл «УБД реєстр».");
      return;
    }
    setIsBusy(true);
    try {
      const next = await parseUbdRegistryStats(ubdRoster);
      setUbdStats(next);
      setUbdYearFilter(null);
      setBatchListSheet(null);
      const yearParts = next.byYear
        .filter((row) => row.year !== "unknown")
        .map((row) => `${row.year}: ${row.submitted}`)
        .join(", ");
      const batchPart = next.colorBatches.length
        ? ` Пакети (отримали 2026): ${next.totals.batchReceived2026} (${next.colorBatches
            .map((batch) => `${batch.sheetName}: ${batch.received2026}`)
            .join("; ")}).`
        : "";
      setMessage(
        `УБД реєстр «${next.fileName}»: подавалися ${next.totals.submitted}, не подавалися ${next.totals.notSubmitted}. По роках (Подавалися): ${yearParts || "немає даних"}${next.totals.yearUnknown ? `; без року: ${next.totals.yearUnknown}` : ""}.${batchPart}`,
      );
      (
        window as Window & { __SOC_PASSPORT_UBD_STATS__?: unknown }
      ).__SOC_PASSPORT_UBD_STATS__ = next;
    } catch (error) {
      setUbdStats(null);
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося порахувати статистику УБД.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const ubdYearPeople = useMemo(() => {
    if (!ubdStats || ubdYearFilter == null) return [];
    return ubdStats.people.filter((person) => person.year === ubdYearFilter);
  }, [ubdStats, ubdYearFilter]);

  const exportWorkbook = async () => {
    if (!result) return;
    setIsBusy(true);
    try {
      const stats = await exportSocPassportWorkbook(
        result,
        departures,
        housingIdpStats,
      );
      const ejoosPart = stats.hasEjoosTables
        ? ` + аркуші ЕЖООС: «Вибули» (${stats.departuresTotal}), «Прибули ООС» (${stats.arrivalsTotal})${
            stats.arrivalsMorningTotal
              ? `, «Звідки Штатка» (${stats.arrivalsMorningTotal})`
              : ""
          }${
            stats.arrivalsPbTotal
              ? `, «Прибули 1ПБ» (${stats.arrivalsPbTotal})`
              : ""
          }`
        : " (немає даних ЕЖООС Вибули/Прибули — завантажте ЕЖООС)";
      const housingPart = stats.housingIdpRemaining
        ? ` + ВПО житло (лишилось ${stats.housingIdpRemaining} з ${stats.housingIdpInFile})`
        : "";
      setMessage(
        `Експортовано Excel: «Соц.портрет», «Розбір», «Офіцери» (${stats.officersCount}), «Рф» (${stats.russiaCount}), «Виходи ЛБЗ» (${stats.exitsCount}), «Не виконували» (${stats.noExitsCount})${ejoosPart}${housingPart}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося експортувати Excel.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const exportExitsOnly = async () => {
    if (!result) return;
    setIsBusy(true);
    try {
      const stats = await exportSocPassportExitsWorkbook(result);
      setMessage(
        `Експортовано «Виходи ЛБЗ»: ${stats.exitsCount} + «Не виконували» (${stats.noExitsCount}) + «Офіцери» (${stats.officersCount}) + «Рф» (${stats.russiaCount}).`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося експортувати Excel.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const exportRussiaOnly = async () => {
    if (!result) return;
    setIsBusy(true);
    try {
      const count = await exportSocPassportRussiaWorkbook(result);
      setMessage(`Експортовано «Рф»: ${count} осіб (ПІБ + підстава класифікації).`);
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося експортувати Excel.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const exportNoExitsOnly = async () => {
    if (!result) return;
    setIsBusy(true);
    try {
      const count = await exportSocPassportNoExitsWorkbook(result);
      setMessage(`Експортовано «Не виконували»: ${count} осіб (ПІБ + примітки).`);
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося експортувати Excel.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const exportDeparturesOnly = async () => {
    if (!departures) return;
    setIsBusy(true);
    try {
      const totals = await exportSocPassportDeparturesWorkbook(departures);
      const arrivalsPart =
        totals.arrivalsTotal > 0
          ? ` + прибули ООС ${totals.arrivalsTotal}`
          : "";
      const arrivalsMorningPart =
        totals.arrivalsMorningTotal > 0
          ? ` + звідки Штатка ${totals.arrivalsMorningTotal}`
          : "";
      const arrivalsPbPart =
        totals.arrivalsPbTotal > 0
          ? ` + прибули 1ПБ ${totals.arrivalsPbTotal}`
          : "";
      setMessage(
        `Експортовано «Вибули»: ${totals.all} записів (звільнення ${totals.discharges}, переведені ${totals.transfers})${arrivalsPart}${arrivalsMorningPart}${arrivalsPbPart}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося експортувати «Вибули».",
      );
    } finally {
      setIsBusy(false);
    }
  };

  return (
    <main className="main-panel bchs-page soc-passport-page">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Соціальний паспорт
          </Typography>
          <Typography variant="body2" color="text.secondary">
            ШПО + ООС + ранковий + додаток про житло (ВПО) + за бажанням ЖБД / БПЛА / УБД
          </Typography>
        </Box>
        <Stack direction="row" spacing={1} sx={{ flexWrap: "wrap", gap: 1 }}>
          <FileButton
            label="ЕЖООС"
            loadedName={ejoos?.fileName}
            disabled={isBusy}
            onFile={(file) => void loadWorkbook(file, "ejoos")}
          />
          <FileButton
            label="1ПБ"
            loadedName={pb?.fileName}
            disabled={isBusy}
            onFile={(file) => void loadWorkbook(file, "pb")}
          />
          <FileButton
            label="Ранковий звіт"
            loadedName={morning?.fileName}
            disabled={isBusy}
            onFile={(file) => void loadWorkbook(file, "morning")}
          />
          <FileButton
            label="Додаток житло"
            loadedName={housingIdp?.fileName}
            disabled={isBusy}
            onFile={(file) => void loadWorkbook(file, "housing")}
          />
          <FileButton
            label="ЖБД виходи"
            loadedName={jbdExits?.fileName}
            disabled={isBusy}
            onFile={(file) => void loadWorkbook(file, "jbd")}
          />
          <FileButton
            label="БПЛА виходи"
            loadedName={bplaExits?.fileName}
            disabled={isBusy}
            onFile={(file) => void loadWorkbook(file, "bpla")}
          />
          <FileButton
            label="УБД реєстр"
            loadedName={ubdRoster?.fileName}
            disabled={isBusy}
            onFile={(file) => void loadWorkbook(file, "ubd")}
          />
          <SciButton variant="EXEC" disabled={isBusy || !ejoos || !morning} onClick={calculate}>
            <TableChartOutlinedIcon fontSize="small" />
            {isBusy ? "Рахую…" : "Порахувати"}
          </SciButton>
          <SciButton
            variant="OUTLINE"
            disabled={isBusy || !ubdRoster}
            onClick={() => void calculateUbdStats()}
            title="Окремий підрахунок по файлу УБД реєстру: скільки подано / зроблено по роках"
          >
            <TableChartOutlinedIcon fontSize="small" />
            Підрахунок УБД
          </SciButton>
          <SciButton
            variant="OUTLINE"
            disabled={isBusy || !result}
            onClick={() => void exportWorkbook()}
            title={
              departures
                ? "Соц.портрет + розбір + Вибули/Прибули з ЕЖООС"
                : "Соц.портрет і розбір (для Вибули/Прибули спочатку завантажте ЕЖООС)"
            }
          >
            <FileDownloadOutlinedIcon fontSize="small" />
            Excel
          </SciButton>
          <SciButton
            variant="OUTLINE"
            disabled={isBusy || !result}
            onClick={() => void exportExitsOnly()}
            title="Виходи на ЛБЗ + аркуш «Не виконували»"
          >
            <FileDownloadOutlinedIcon fontSize="small" />
            Виходи ЛБЗ
          </SciButton>
          <SciButton
            variant="OUTLINE"
            disabled={isBusy || !result}
            onClick={() => void exportNoExitsOnly()}
            title="Окремий файл: ПІБ з рядка «Не виконували»"
          >
            <FileDownloadOutlinedIcon fontSize="small" />
            Не виконували
          </SciButton>
          <SciButton
            variant="OUTLINE"
            disabled={isBusy || !result}
            onClick={() => void exportRussiaOnly()}
            title="Окремий файл: хто потрапив у рядок «Рф»"
          >
            <FileDownloadOutlinedIcon fontSize="small" />
            Рф (ПІБ)
          </SciButton>
          <SciButton
            variant="OUTLINE"
            disabled={isBusy || !departures}
            onClick={() => void exportDeparturesOnly()}
            title="Окремий файл: Вибули з «Виключені» + Прибули за серпень (ООС і/або 1ПБ)"
          >
            <FileDownloadOutlinedIcon fontSize="small" />
            Вибули
          </SciButton>
        </Stack>
      </header>

      <Alert severity="info" sx={{ mb: 2 }}>
        {message}
      </Alert>

      {departures ? (
        <section className="analytics-panel">
          <div className="panel-heading">
            Вибули · {departures.sourceSheet}
            {departures.periodFrom ? ` · з ${departures.periodFrom}` : ""}
          </div>
          <Typography variant="body2" sx={{ mb: 1.5, lineHeight: 1.7 }}>
            Джерело: аркуш ЕЖООС «Виключені». Рахуємо рядки з{" "}
            <strong>датою виключення / наказу</strong>
            {departures.periodFromLabel
              ? ` ${departures.periodFromLabel}`
              : ""}
            . У файлі всього {departures.totalUnfiltered}
            {departures.skippedBeforePeriod || departures.skippedNoDate
              ? ` (до періоду ${departures.skippedBeforePeriod}, без дати ${departures.skippedNoDate})`
              : ""}
            . Звільнення — 5 категорій; переведені / розпорядження — окремо.
          </Typography>
          <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
            <table className="bchs-analytics-table">
              <thead>
                <tr>
                  <th>№</th>
                  <th>Категорія</th>
                  <th>Офіцери</th>
                  <th>Сержанти</th>
                  <th>Солдати</th>
                  <th>Разом</th>
                </tr>
              </thead>
              <tbody>
                {departures.summary.map((row, index) => (
                  <tr key={row.category}>
                    <td>
                      {row.category === "transfer" ? "—" : String(index + 1)}
                    </td>
                    <td>{row.label}</td>
                    <td>{row.byRank.officer || ""}</td>
                    <td>{row.byRank.sergeant || ""}</td>
                    <td>{row.byRank.soldier || ""}</td>
                    <td>
                      <strong>{row.count}</strong>
                    </td>
                  </tr>
                ))}
                <tr>
                  <td />
                  <td>
                    <strong>Усього у періоді</strong>
                  </td>
                  <td>
                    <strong>{departures.totals.byRank.officer}</strong>
                  </td>
                  <td>
                    <strong>{departures.totals.byRank.sergeant}</strong>
                  </td>
                  <td>
                    <strong>{departures.totals.byRank.soldier}</strong>
                  </td>
                  <td>
                    <strong>{departures.totals.all}</strong>
                  </td>
                </tr>
                <tr>
                  <td />
                  <td>Звільнення / інше (без переведень)</td>
                  <td>{departures.totals.dischargesByRank.officer || ""}</td>
                  <td>{departures.totals.dischargesByRank.sergeant || ""}</td>
                  <td>{departures.totals.dischargesByRank.soldier || ""}</td>
                  <td>{departures.totals.discharges}</td>
                </tr>
                <tr>
                  <td />
                  <td>Переведені / розпорядження</td>
                  <td>{departures.totals.transfersByRank.officer || ""}</td>
                  <td>{departures.totals.transfersByRank.sergeant || ""}</td>
                  <td>{departures.totals.transfersByRank.soldier || ""}</td>
                  <td>{departures.totals.transfers}</td>
                </tr>
              </tbody>
            </table>
          </div>

          {departures.arrivalsAugust ? (
            <ArrivalsPreview
              arrivals={departures.arrivalsAugust}
              titleHint="Прибули з ЕЖООС · 2. ООС за «Дата зарахування». Джерела — колонка «Звідки прибув» (порожнє поле = ТЦК)."
            />
          ) : null}
          {departures.arrivalsFromMorning ? (
            <ArrivalsPreview
              arrivals={departures.arrivalsFromMorning}
              headingPrefix="Звідки прибув"
              titleHint="База — Штатка (ранковий): спочатку БРЕЗ з колонок БЗВП/БРЕЗ; решта — «Звідки прибув» з ООС ЕЖООС (порожнє = ТЦК). Розбивка за званням нижче."
            />
          ) : null}
          {departures.arrivalsAugustPb ? (
            <ArrivalsPreview
              arrivals={departures.arrivalsAugustPb}
              titleHint="Прибули з 1ПБ · Рух (ПРИБУВ). Джерела — з «Звідки прибув» (ООС / sh / Рух)."
            />
          ) : null}
          {(departures.dispositionFromArchive || pbDisposition) && (
            <DispositionPreview
              disposition={
                departures.dispositionFromArchive || pbDisposition!
              }
            />
          )}
          {(departures.szchFromRuh || pbSzch) && (
            <SzchPreview szch={departures.szchFromRuh || pbSzch!} />
          )}
          {(departures.combatLossesFromPb || pbCombatLosses) && (
            <CombatLossesPreview
              losses={departures.combatLossesFromPb || pbCombatLosses!}
            />
          )}
        </section>
      ) : null}

      {!departures && pbArrivals ? (
        <section className="analytics-panel">
          <ArrivalsPreview
            arrivals={pbArrivals}
            titleHint="Прибули з 1ПБ · Рух (ПРИБУВ). Для точніших ТЦК/НЦ/переведених завантажте також ЕЖООС — підтягнемо «Звідки прибув» з ООС."
          />
          {pbDisposition ? (
            <DispositionPreview disposition={pbDisposition} />
          ) : null}
          {pbSzch ? <SzchPreview szch={pbSzch} /> : null}
          {pbCombatLosses ? (
            <CombatLossesPreview losses={pbCombatLosses} />
          ) : null}
        </section>
      ) : null}

      {!departures &&
      !pbArrivals &&
      (pbDisposition || pbSzch || pbCombatLosses) ? (
        <section className="analytics-panel">
          {pbDisposition ? (
            <DispositionPreview disposition={pbDisposition} />
          ) : null}
          {pbSzch ? <SzchPreview szch={pbSzch} /> : null}
          {pbCombatLosses ? (
            <CombatLossesPreview losses={pbCombatLosses} />
          ) : null}
        </section>
      ) : null}

      <section className="analytics-panel">
        <div className="panel-heading">Джерела</div>
        <Typography variant="body2" component="div" sx={{ lineHeight: 1.7 }}>
          <p>
            <strong>Ранковий звіт</strong> — основа підрахунку, як у БЧС: лише
            батальйон «нова». За штатом — усі рядки посад; за списком — є звання
            (колонка M); в наявності — «В строю» + «Відком. за межі ПБ»; в
            розпорядженні — решта списку. «Статус бійців» дає виходи на ЛБЗ;
            <strong> ЖБД</strong> (напр. «ВИХІД ПОКРОВСЬК») — додаткові виходи; якщо
            дата і ФІО збігаються з «Статус бійців», рахується як один вихід.
            Якщо виходів немає — підставляємо з ЕЖООС: УБД в анкеті ООС,
            «Місце дислокації» та аркуш «4. Тимчасово прибулі» (зона бойових
            завдань). <strong>УБД реєстр</strong> («Подавалися» / «Не
            подавалися») — хто є в цьому файлі і також у штатці (ранковий
            список), потрапляє в рядок статистики «УБД»; обидві групи також
            рахуються як «виконували» (мін. 1 вихід, якщо інших джерел немає).
            Додатково в «УБД» лишається позначка з анкети ООС.{" "}
            <strong>Додаток житло</strong> (кнопка поруч із «Ранковий звіт») —
            список ВПО з «Додаток про житло.xlsx»; зіставляємо з ранковим і
            рахуємо хто лишився — вони входять у рядок «ВПО» соцпакета.{" "}
            <strong>БПЛА виходи</strong> (таблиця з колонкою «Виконував БЗ» /
            «Не виконував БЗ») — хто позначений як виконував, теж рахується як
            «виконували» (мін. 1 вихід), якщо немає інших джерел. Статус ранкового
            звіту <strong>«Зниклі безвісти»</strong> також зараховується як
            «виконували» (мін. 1 вихід). Колонка <strong>«В якому
            підрозділі» = РРЕБ</strong> (РЕБ / рота РРЕБ) або <strong>ППО</strong> —
            теж «виконували» (мін. 1 вихід).{" "}
            <strong>Статичний список</strong> (14 осіб, прибрані вручну з Excel
            «Не виконували» 22.08 vs повна 23.08) — теж «виконували» (мін. 1
            вихід); список нижче на сторінці.
          </p>
          <p>
            <strong>Колонка «Є в «Статус бійців»»</strong> у Excel — чи є для
            цієї людини <em>хоча б одна дата виходу</em> на аркуші ранкового
            звіту «Статус бійців» (до об’єднання з ЖБД). У таблиці «Виходи ЛБЗ»
            майже завжди «так»; у «Не виконували» — «ні». З «Не виконували»
            виключені транзитери, «Відком. за межі ПБ», СЗЧ, «Лікування» і
            «по пораненню».
          </p>
          <p>
            <strong>ЕЖООС:</strong> {formatFileLabel(ejoos?.fileName)}. ООС —
            анкети (регіон, родичі, діти, УБД, дислокація). ШПО лише підклеює
            індекс посади, штат із нього більше не береться. «Тимчасово
            прибулі» — хто був у зоні виконання бойових завдань.
          </p>
          <p>
            Діти, сімейний стан, національність і регіон — regex по анкетах.
            Якщо поле порожнє: стать = чоловік, національність = Україна,
            сімейний стан = неодружений, дітей немає. Вік без дати лишається
            порожнім. Перевірка — аркуш «Розбір» у Excel.
          </p>
        </Typography>
      </section>

      {housingIdpStats ? (
        <section className="analytics-panel">
          <div className="panel-heading">
            ВПО · додаток про житло · {housingIdpStats.fileName}
          </div>
          <Typography variant="body2" sx={{ mb: 1.5, lineHeight: 1.7 }}>
            ПІБ з додатку про житло зіставлено з ранковим звітом. У соцпакет
            (рядок «ВПО») входять ті, хто{" "}
            <strong>лишився в ранковому</strong>.
          </Typography>
          <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
            <table className="bchs-analytics-table">
              <thead>
                <tr>
                  <th>Показник</th>
                  <th>Офіцери</th>
                  <th>Сержанти</th>
                  <th>Солдати</th>
                  <th>Разом</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>У файлі (додаток)</td>
                  <td colSpan={3} />
                  <td>
                    <strong>{housingIdpStats.totals.inFile}</strong>
                  </td>
                </tr>
                <tr>
                  <td>Лишилось у ранковому</td>
                  <td>
                    {housingIdpStats.totals.byRankRemaining.officer || ""}
                  </td>
                  <td>
                    {housingIdpStats.totals.byRankRemaining.sergeant || ""}
                  </td>
                  <td>
                    {housingIdpStats.totals.byRankRemaining.soldier || ""}
                  </td>
                  <td>
                    <strong>{housingIdpStats.totals.remaining}</strong>
                  </td>
                </tr>
                <tr>
                  <td>З них у наявності (в строю / відком.)</td>
                  <td colSpan={3} />
                  <td>
                    <strong>{housingIdpStats.totals.remainingPresent}</strong>
                  </td>
                </tr>
                <tr>
                  <td>Немає в ранковому (вибули / не знайдені)</td>
                  <td colSpan={3} />
                  <td>
                    <strong>{housingIdpStats.totals.leftNotInMorning}</strong>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
          <div
            className="bchs-analytics-table-wrap soc-passport-table-wrap"
            style={{ marginTop: 12 }}
          >
            <table className="bchs-analytics-table">
              <thead>
                <tr>
                  <th>№</th>
                  <th>ПІБ</th>
                  <th>Звання</th>
                  <th>У ранковому</th>
                  <th>Статус</th>
                  <th>Примітка</th>
                </tr>
              </thead>
              <tbody>
                {housingIdpStats.people.map((person, index) => (
                  <tr key={`${person.fullName}-${person.excelRow}`}>
                    <td>{index + 1}</td>
                    <td>{person.fullName}</td>
                    <td>{person.rank}</td>
                    <td>{person.remaining ? "так" : "ні"}</td>
                    <td>{person.morningStatus || "—"}</td>
                    <td>{person.note || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      ) : null}

      {ubdStats ? (
        <section className="analytics-panel">
          <div className="panel-heading">
            Підрахунок УБД реєстру · {ubdStats.fileName}
          </div>
          <Typography variant="body2" sx={{ mb: 1.5, lineHeight: 1.7 }}>
            Рік для «Подавалися»: спочатку дата видачі / отримання / відправки
            посвідчення; якщо їх немає — рік початку періоду участі в БД.
            Пакети документів: <strong>Июль-август</strong> — усе крім зелених =
            отримали в 2026; <strong>Июнь / май-июнь / апрель-май</strong> —
            зелені = отримали в 2026.
          </Typography>
          {ubdStats.warnings.length ? (
            <Alert severity="warning" sx={{ mb: 1.5 }}>
              {ubdStats.warnings.join(" ")}
            </Alert>
          ) : null}
          <Stack spacing={0.5} sx={{ mb: 2 }}>
            <Typography variant="body2">
              Подавалися: <strong>{ubdStats.totals.submitted}</strong>
              {" · "}
              статус УБД «так»: <strong>{ubdStats.totals.statusTak}</strong>
              {" · "}
              є отримання/видача:{" "}
              <strong>{ubdStats.totals.issuedOrReceived}</strong>
              {" · "}
              з періодом БД: <strong>{ubdStats.totals.withCombatPeriod}</strong>
            </Typography>
            <Typography variant="body2">
              Не подавалися: <strong>{ubdStats.totals.notSubmitted}</strong>
              {" · "}
              пакети документів: <strong>{ubdStats.totals.batchRows}</strong>
              {" · "}
              проблемні: <strong>{ubdStats.totals.problems}</strong>
              {" · "}
              «УБД у мене»: <strong>{ubdStats.totals.onHand}</strong>
            </Typography>
            {ubdStats.colorBatches.length ? (
              <Typography variant="body2">
                Отримали в 2026 (пакети за кольором):{" "}
                <strong>{ubdStats.totals.batchReceived2026}</strong>
              </Typography>
            ) : null}
          </Stack>

          {ubdStats.colorBatches.length ? (
            <>
              <div className="panel-heading" style={{ marginTop: 8 }}>
                Пакети документів · отримали в 2026
              </div>
              <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
                <table className="bchs-analytics-table">
                  <thead>
                    <tr>
                      <th>Аркуш</th>
                      <th>Правило</th>
                      <th>Отримали 2026</th>
                      <th>Зелені</th>
                      <th>У списку</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {ubdStats.colorBatches.map((batch) => (
                      <tr key={batch.sheetName}>
                        <td>{batch.sheetName}</td>
                        <td>{batch.ruleLabel}</td>
                        <td>
                          <strong>{batch.received2026}</strong>
                        </td>
                        <td>{batch.green}</td>
                        <td>{batch.total}</td>
                        <td>
                          <SciButton
                            variant="OUTLINE"
                            size="SM"
                            onClick={() =>
                              setBatchListSheet((current) =>
                                current === batch.sheetName
                                  ? null
                                  : batch.sheetName,
                              )
                            }
                          >
                            {batchListSheet === batch.sheetName
                              ? "Сховати"
                              : "Список"}
                          </SciButton>
                        </td>
                      </tr>
                    ))}
                    <tr>
                      <td>
                        <strong>Разом</strong>
                      </td>
                      <td />
                      <td>
                        <strong>{ubdStats.totals.batchReceived2026}</strong>
                      </td>
                      <td />
                      <td />
                      <td />
                    </tr>
                  </tbody>
                </table>
              </div>
            </>
          ) : null}

          {batchListSheet
            ? (() => {
                const batch = ubdStats.colorBatches.find(
                  (item) => item.sheetName === batchListSheet,
                );
                if (!batch) return null;
                const rows = batch.rows.filter((row) => row.received2026);
                return (
                  <>
                    <div className="panel-heading" style={{ marginTop: 8 }}>
                      {batch.sheetName} · отримали в 2026 ({rows.length})
                    </div>
                    <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
                      <table className="bchs-analytics-table">
                        <thead>
                          <tr>
                            <th>#</th>
                            <th>Позивний</th>
                            <th>ПІБ</th>
                            <th>Мітка</th>
                            <th>Примітка</th>
                          </tr>
                        </thead>
                        <tbody>
                          {rows.map((row) => (
                            <tr key={`${row.excelRow}-${row.name}`}>
                              <td>{row.excelRow}</td>
                              <td>{row.callsign || "—"}</td>
                              <td>{row.name}</td>
                              <td>{row.mark || "—"}</td>
                              <td>{row.note || "—"}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </>
                );
              })()
            : null}

          <div className="panel-heading" style={{ marginTop: 8 }}>
            УБД по роках (аркуш «Подавалися»; видані 2026 + пакети за кольором)
          </div>
          <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
            <table className="bchs-analytics-table">
              <thead>
                <tr>
                  <th>Рік</th>
                  <th>Усього подано</th>
                  <th>З датою процесу</th>
                  <th>З періоду БД</th>
                  <th>Отримано / видано</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {ubdStats.byYear.map((row) => (
                  <tr key={String(row.year)}>
                    <td>{row.year === "unknown" ? "без року" : row.year}</td>
                    <td>
                      <strong>{row.submitted}</strong>
                    </td>
                    <td>{row.fromProcessDate || ""}</td>
                    <td>{row.fromCombatPeriod || ""}</td>
                    <td>
                      <strong>{row.issuedOrReceived || ""}</strong>
                      {row.year === 2026 && row.issuedFromBatches > 0 ? (
                        <span style={{ opacity: 0.75 }}>
                          {" "}
                          (+{row.issuedFromBatches} з пакетів)
                        </span>
                      ) : null}
                    </td>
                    <td>
                        <SciButton
                          variant="OUTLINE"
                          size="SM"
                          onClick={() =>
                            setUbdYearFilter(
                              ubdYearFilter === row.year ? null : row.year,
                            )
                          }
                        >
                        {ubdYearFilter === row.year ? "Сховати" : "Список"}
                      </SciButton>
                    </td>
                  </tr>
                ))}
                <tr>
                  <td>
                    <strong>Разом</strong>
                  </td>
                  <td>
                    <strong>{ubdStats.totals.submitted}</strong>
                  </td>
                  <td />
                  <td />
                  <td>
                    <strong>{ubdStats.totals.issuedOrReceived}</strong>
                  </td>
                  <td />
                </tr>
              </tbody>
            </table>
          </div>

          {ubdYearFilter != null ? (
            <>
              <div className="panel-heading" style={{ marginTop: 16 }}>
                Список ·{" "}
                {ubdYearFilter === "unknown" ? "без року" : ubdYearFilter} (
                {ubdYearPeople.length})
              </div>
              <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
                <table className="bchs-analytics-table">
                  <thead>
                    <tr>
                      <th>ПІБ</th>
                      <th>Звання</th>
                      <th>Період БД</th>
                      <th>Джерело року</th>
                      <th>Статус УБД</th>
                      <th>Отримано / видано</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ubdYearPeople.map((person) => (
                      <tr key={`${person.name}-${person.rnokpp}`}>
                        <td>{person.name}</td>
                        <td>{person.rank}</td>
                        <td>{person.combatPeriod || "—"}</td>
                        <td>
                          {person.yearSource === "issue"
                            ? "дата видачі"
                            : person.yearSource === "received"
                              ? "отримано"
                              : person.yearSource === "sent"
                                ? "відправлено"
                                : person.yearSource === "combat"
                                  ? "період БД"
                                  : "—"}
                        </td>
                        <td>{person.ubdStatus || "—"}</td>
                        <td>
                          {[person.received, person.issued]
                            .filter(Boolean)
                            .join(" · ") || "—"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          ) : null}

          <div className="panel-heading" style={{ marginTop: 16 }}>
            Усі аркуші файлу
          </div>
          <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
            <table className="bchs-analytics-table">
              <thead>
                <tr>
                  <th>Аркуш</th>
                  <th>Тип</th>
                  <th>Людей</th>
                  <th>З періодом БД</th>
                  <th>Статус «так»</th>
                  <th>Отримано</th>
                  <th>Видано</th>
                </tr>
              </thead>
              <tbody>
                {ubdStats.bySheet.map((sheet) => (
                  <tr key={sheet.sheetName}>
                    <td>{sheet.sheetName}</td>
                    <td>
                      {sheet.bucket === "submitted"
                        ? "Подавалися"
                        : sheet.bucket === "notSubmitted"
                          ? "Не подавалися"
                          : sheet.bucket === "batch"
                            ? "Пакет документів"
                            : sheet.bucket === "problems"
                              ? "Проблемні"
                              : sheet.bucket === "onHand"
                                ? "УБД у мене"
                                : "Інше"}
                    </td>
                    <td>{sheet.peopleCount}</td>
                    <td>{sheet.withCombatPeriod || ""}</td>
                    <td>{sheet.statusTak || ""}</td>
                    <td>{sheet.withReceived || ""}</td>
                    <td>{sheet.withIssued || ""}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      ) : null}

      <section className="analytics-panel">
        <div className="panel-heading">
          Статичні бойові виходи ({STATIC_COMBAT_EXIT_OVERRIDES.length}
          {result ? ` · у ранковому: ${staticCombatMatchedCount}` : ""})
        </div>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
          {STATIC_COMBAT_EXIT_OVERRIDE_SOURCE} У статистиці «По виконанню бойових
          завдань» рахуються як мін. 1 вихід (навіть якщо в «Статус бійців» /
          ЖБД порожньо).
        </Typography>
        <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
          <table className="bchs-analytics-table">
            <thead>
              <tr>
                <th>ПІБ</th>
                <th>Позивний</th>
                <th>Звання</th>
                <th>У ранковому</th>
                <th>Виходів (після override)</th>
                <th>Примітка</th>
              </tr>
            </thead>
            <tbody>
              {staticCombatExitRows.map((row) => (
                <tr key={row.name}>
                  <td>{row.name}</td>
                  <td>{row.callsign || "—"}</td>
                  <td>{row.rank}</td>
                  <td>
                    {result ? (row.inMorning ? "так" : "ні") : "—"}
                  </td>
                  <td>{row.exitCount == null ? "—" : row.exitCount}</td>
                  <td>{row.note}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {result ? (
        <>
          <section className="analytics-panel">
            <div className="panel-heading">Короткий зріз</div>
            <Stack spacing={0.75}>
              <Typography variant="body2">
                Штат / на посаді / вакант:{" "}
                <strong>
                  {result.summary.staffSlots} / {result.summary.occupied} /{" "}
                  {result.summary.vacant}
                </strong>
              </Typography>
              <Typography variant="body2">
                Збіги ООС / ранок / виходи (ранок) / бойове (ЕЖООС):{" "}
                <strong>
                  {result.summary.oosMatched} / {result.summary.morningMatched} /{" "}
                  {result.summary.exitsMatched} / {result.summary.combatDutyMatched}
                </strong>
              </Typography>
              <Typography variant="body2">
                Родичі заповнені: {result.summary.relativesParsed} · регіон не
                визначено: {result.summary.unknownRegion} · вік не визначено:{" "}
                {result.summary.unknownAge}
              </Typography>
              <Typography variant="body2" color="text.secondary">
                Аркуші: {result.sheets.shpo} · {result.sheets.oos}
                {result.sheets.morning ? ` · ${result.sheets.morning}` : ""}
                {result.sheets.fighterStatus
                  ? ` · ${result.sheets.fighterStatus}`
                  : ""}
                {result.sheets.tempArrived
                  ? ` · ${result.sheets.tempArrived}`
                  : ""}
                {result.sheets.jbdExits ? ` · ЖБД: ${result.sheets.jbdExits}` : ""}
                {result.sheets.bplaExits ? ` · БПЛА: ${result.sheets.bplaExits}` : ""}
                {result.sheets.ubdRoster
                  ? ` · УБД: ${result.sheets.ubdRoster}`
                  : ""}
              </Typography>
              {result.warnings.map((warning) => (
                <Typography key={warning} variant="body2" color="text.secondary">
                  {warning}
                </Typography>
              ))}
            </Stack>
          </section>

          <section className="analytics-panel">
            <div className="panel-heading">Соціальний портрет</div>
            <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
              <table className="bchs-analytics-table soc-passport-table">
                <thead>
                  <tr>
                    <th rowSpan={2}>№</th>
                    <th rowSpan={2}>Вид обліку</th>
                    <th colSpan={2}>Офіцери</th>
                    <th colSpan={2}>Сержанти</th>
                    <th colSpan={2}>Солдати</th>
                    <th rowSpan={2}>Всього</th>
                  </tr>
                  <tr>
                    <th>Моб.</th>
                    <th>Контр.</th>
                    <th>Моб.</th>
                    <th>Контр.</th>
                    <th>Моб.</th>
                    <th>Контр.</th>
                  </tr>
                </thead>
                <tbody>
                  {result.rows.map((row) => (
                    <tr
                      key={row.id}
                      className={row.kind === "section" ? "soc-passport-section" : undefined}
                    >
                      <td>{row.number}</td>
                      <td>{row.label}</td>
                      <CountsCells row={row} />
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <section className="analytics-panel">
            <div className="panel-heading">
              Не виконували ({noExitsPeople.length})
            </div>
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              Особи з 0 виходів (без транзитерів). Повний список — кнопка «Не
              виконували» або аркуш у загальному Excel.
            </Typography>
            <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
              <table className="bchs-analytics-table">
                <thead>
                  <tr>
                    <th>ПІБ</th>
                    <th>Позивний</th>
                    <th>Посада</th>
                    <th>Звання</th>
                    <th>Кат.</th>
                    <th>Статус (ранок)</th>
                    <th>Є в «Статус бійців»</th>
                  </tr>
                </thead>
                <tbody>
                  {noExitsPeople.slice(0, 200).map((person) => (
                    <tr key={person.id}>
                      <td>{person.name}</td>
                      <td>{person.callsign || "—"}</td>
                      <td>{person.position}</td>
                      <td>{person.rank}</td>
                      <td>
                        {person.rankGroup === "officer"
                          ? "оф."
                          : person.rankGroup === "sergeant"
                            ? "серж."
                            : "солд."}
                      </td>
                      <td>{person.morningStatus}</td>
                      <td>{person.morningExitCount > 0 ? "так" : "ні"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {noExitsPeople.length > 200 ? (
              <Typography variant="caption" color="text.secondary">
                Показано 200 з {noExitsPeople.length}. Решта — у Excel.
              </Typography>
            ) : null}
          </section>

          <section className="analytics-panel">
            <div className="panel-heading">Розбір анкет</div>
            <input
              className="soc-passport-filter"
              value={query}
              placeholder="Пошук за ПІБ, посадою, регіоном, родичами…"
              onChange={(event) => setQuery(event.target.value)}
            />
            <div className="bchs-analytics-table-wrap soc-passport-table-wrap">
              <table className="bchs-analytics-table">
                <thead>
                  <tr>
                    <th>ПІБ</th>
                    <th>Кат.</th>
                    <th>Вік</th>
                    <th>Регіон</th>
                    <th>Сім’я</th>
                    <th>Діти&lt;18</th>
                    <th>УБД</th>
                    <th>Виходи</th>
                    <th>Примітки</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredPeople.slice(0, 250).map((person) => (
                    <tr key={person.id}>
                      <td>{person.name}</td>
                      <td>
                        {person.rankGroup === "officer"
                          ? "оф."
                          : person.rankGroup === "sergeant"
                            ? "серж."
                            : "солд."}
                      </td>
                      <td>{person.age ?? ""}</td>
                      <td>{person.regionLabel}</td>
                      <td>
                        {person.marital === "married"
                          ? "одруж."
                          : person.marital === "civil"
                            ? "цивільн."
                            : person.marital === "unmarried"
                              ? "неодруж."
                              : ""}
                      </td>
                      <td>{person.childrenUnder18 || ""}</td>
                      <td>{person.hasUbd ? "так" : ""}</td>
                      <td>{person.exitCount || ""}</td>
                      <td>{person.parseNotes.slice(0, 2).join("; ")}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {filteredPeople.length > 250 ? (
              <Typography variant="caption" color="text.secondary">
                Показано 250 з {filteredPeople.length}. Повний розбір — у Excel.
              </Typography>
            ) : (
              <Typography variant="caption" color="text.secondary">
                Рядків: {filteredPeople.length}
              </Typography>
            )}
          </section>
        </>
      ) : null}
    </main>
  );
}
