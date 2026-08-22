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
import { countsInNoExitsList } from "./socPassportFields";
import { parseSocPassportSources } from "./socPassportParse";
import type { PassportTableRow, SocPassportResult } from "./socPassportTypes";

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

export function SocPassportPage() {
  const [ejoos, setEjoos] = useState<ExcelWorkbookSnapshot | null>(null);
  const [morning, setMorning] = useState<ExcelWorkbookSnapshot | null>(null);
  const [jbdExits, setJbdExits] = useState<ExcelWorkbookSnapshot | null>(null);
  const [bplaExits, setBplaExits] = useState<ExcelWorkbookSnapshot | null>(null);
  const [ubdRoster, setUbdRoster] = useState<ExcelWorkbookSnapshot | null>(null);
  const [result, setResult] = useState<SocPassportResult | null>(null);
  const [message, setMessage] = useState(
    "Завантажте ЕЖООС (ШПО + ООС) і за бажанням ранковий звіт — потім натисніть «Порахувати».",
  );
  const [isBusy, setIsBusy] = useState(false);
  const [query, setQuery] = useState("");

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
          !person.ubdRosterStatus,
      )
      .sort((left, right) => left.name.localeCompare(right.name, "uk"));
  }, [result]);

  const loadWorkbook = async (
    file: File | undefined,
    kind: "ejoos" | "morning" | "jbd" | "bpla" | "ubd",
  ) => {
    if (!file) return;
    setIsBusy(true);
    try {
      const snapshot = await readWorkbookSnapshot(file);
      if (kind === "ejoos") setEjoos(snapshot);
      else if (kind === "morning") setMorning(snapshot);
      else if (kind === "jbd") setJbdExits(snapshot);
      else if (kind === "bpla") setBplaExits(snapshot);
      else setUbdRoster(snapshot);
      setResult(null);
      setMessage(
        kind === "ejoos"
          ? `ЕЖООС: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.`
          : kind === "morning"
            ? `Ранковий звіт: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.`
            : kind === "jbd"
              ? `ЖБД виходи: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.`
              : kind === "bpla"
                ? `БПЛА виходи: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.`
                : `УБД реєстр: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.`,
      );
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
        parsed.people.length === 0 ? "На ШПО не знайдено осіб на посадах." : "",
      ].filter(Boolean);
      const next = buildSocPassportResult({
        people: parsed.people,
        staffSlots: parsed.staffSlots,
        sheets: parsed.sheets,
        warnings,
      });
      setResult(next);

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
        `Пораховано: штат ${next.summary.staffSlots}, на посаді ${next.summary.occupied}, ООС ${next.summary.oosMatched}, ранок ${next.summary.morningMatched}. БРЕЗ: ${brez.counts.oosBrezAll}→${brez.counts.countedBrezAll}; УБД: ${ubd.counts.oosUbdAll}→${ubd.counts.countedUbdInMorning}; бойове (ЕЖООС): ${combat.counts.inferredExitFloor}; ЖБД: ${jbd.counts.peopleWithJbdMerged}; БПЛА: ${bpla.counts.matchedPerformers}; УБД реєстр: ${ubdReg.counts.matchedSubmitted + ubdReg.counts.matchedNotSubmitted} (див. console).`,
      );
    } catch (error) {
      setResult(null);
      setMessage(
        error instanceof Error ? error.message : "Не вдалося порахувати соц.паспорт.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const exportWorkbook = async () => {
    if (!result) return;
    setIsBusy(true);
    try {
      const stats = await exportSocPassportWorkbook(result);
      setMessage(
        `Експортовано Excel: «Соц.портрет», «Розбір», «Офіцери» (${stats.officersCount}), «Рф» (${stats.russiaCount}), «Виходи ЛБЗ» (${stats.exitsCount}), «Не виконували» (${stats.noExitsCount}).`,
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

  return (
    <main className="main-panel bchs-page soc-passport-page">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Соціальний паспорт
          </Typography>
          <Typography variant="body2" color="text.secondary">
            ШПО (хто на посаді) + анкети ООС + ранковий звіт (наявність і виходи)
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
            label="Ранковий звіт"
            loadedName={morning?.fileName}
            disabled={isBusy}
            onFile={(file) => void loadWorkbook(file, "morning")}
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
            disabled={isBusy || !result}
            onClick={() => void exportWorkbook()}
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
        </Stack>
      </header>

      <Alert severity="info" sx={{ mb: 2 }}>
        {message}
      </Alert>

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
            подавалися») — означає наявність бойових виходів; обидві групи
            рахуються як «виконували» (мін. 1 вихід, якщо інших джерел немає).
            <strong> БПЛА виходи</strong> (таблиця з колонкою «Виконував БЗ» /
            «Не виконував БЗ») — хто позначений як виконував, теж рахується як
            «виконували» (мін. 1 вихід), якщо немає інших джерел. Статус ранкового
            звіту <strong>«Зниклі безвісти»</strong> також зараховується як
            «виконували» (мін. 1 вихід). Колонка <strong>«В якому
            підрозділі» = РРЕБ</strong> (РЕБ / рота РРЕБ) або <strong>ППО</strong> —
            теж «виконували» (мін. 1 вихід).
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
