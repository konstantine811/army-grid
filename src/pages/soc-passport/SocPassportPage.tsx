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
import { exportSocPassportWorkbook } from "./socPassportExport";
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

  const loadWorkbook = async (
    file: File | undefined,
    kind: "ejoos" | "morning",
  ) => {
    if (!file) return;
    setIsBusy(true);
    try {
      const snapshot = await readWorkbookSnapshot(file);
      if (kind === "ejoos") setEjoos(snapshot);
      else setMorning(snapshot);
      setResult(null);
      setMessage(
        kind === "ejoos"
          ? `ЕЖООС: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.`
          : `Ранковий звіт: ${snapshot.fileName} · аркушів ${snapshot.sheets.length}.`,
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
      const parsed = parseSocPassportSources({ ejoos, morning });
      const warnings = [
        morning
          ? ""
          : "Ранковий звіт не завантажено: «В наявності», «В розпорядженні» і виходи на ЛБЗ будуть порожні або нульові.",
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
      console.log("5) У підрахунку «З Брез» лише з колонки БЗВП ранку:", brez.countedBrezFromMorningOnly);
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

      setMessage(
        `Пораховано: штат ${next.summary.staffSlots}, на посаді ${next.summary.occupied}, ООС ${next.summary.oosMatched}, ранок ${next.summary.morningMatched}. БРЕЗ: ${brez.counts.oosBrezAll}→${brez.counts.countedBrezAll}; УБД: ${ubd.counts.oosUbdAll}→${ubd.counts.countedUbdInMorning} (див. console).`,
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
      await exportSocPassportWorkbook(result);
      setMessage("Експортовано Excel: аркуш «Соц.портрет» і розбір по людях.");
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
            розпорядженні — решта списку. «Статус бійців» дає виходи на ЛБЗ.
          </p>
          <p>
            <strong>ЕЖООС:</strong> {formatFileLabel(ejoos?.fileName)}. ООС —
            анкети (регіон, родичі, діти, УБД). ШПО лише підклеює індекс посади,
            штат із нього більше не береться.
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
                Збіги ООС / ранок / виходи:{" "}
                <strong>
                  {result.summary.oosMatched} / {result.summary.morningMatched} /{" "}
                  {result.summary.exitsMatched}
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
