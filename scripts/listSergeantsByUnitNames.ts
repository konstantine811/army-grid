/**
 * Список ПІБ для перевірки звіту «2 а по підрозділам».
 * npx jiti scripts/listSergeantsByUnitNames.ts [month] [year]
 */
import { readFileSync } from "node:fs";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
  type ReadWorkbookOptions,
} from "../src/excelRoundTrip";
import { parseExcelLabState } from "../src/pages/excel-lab/ExcelLabStateParser";
import {
  buildCombatLossesFromPb,
  buildDeparturesFromEjoos,
  buildDispositionFromArchive,
  buildSzchFromRuh,
} from "../src/pages/soc-passport/socPassportDepartures";
import {
  buildRozporadzhennyaPersonKeys,
  classifySergeantsCombatLoss,
  isDateInMonth,
  monthSinceDate,
  monthUntilDate,
} from "../src/pages/sergeants/sergeantsByUnitReport";
import { normalizeLooseText } from "../src/pages/soc-passport/socPassportFields";

const asFile = (path: string) => {
  const buffer = readFileSync(path) as unknown as File & { name: string };
  Object.defineProperty(buffer, "name", {
    value: path.split("/").pop() ?? path,
  });
  return buffer;
};

const STAFF_READ_OPTIONS: ReadWorkbookOptions = {
  sheetFilter: (name) => /загальний\s+список/i.test(name.trim()),
  maxColumns: 45,
  skipStyleFills: true,
  preserveLeadingColumns: true,
};

const month = Number(process.argv[2] || 9);
const year = Number(process.argv[3] || 2026);
const since = monthSinceDate(month, year);
const until = monthUntilDate(month, year);

const pb = await readWorkbookSnapshot(
  asFile("/Volumes/KINGSTON/army_work/ЕЖООС/08.09.2026/1ПБ_07092026.xlsx"),
  EJOOS_SYNC_READ_OPTIONS,
);
const ejoos = await readWorkbookSnapshot(
  asFile(
    "/Volumes/KINGSTON/army_work/ЕЖООС/15.09.2026/ЄЖООС_станом_на_07-09-2026.xlsx",
  ),
  EJOOS_SYNC_READ_OPTIONS,
);
const staff = await readWorkbookSnapshot(
  asFile("/Volumes/KINGSTON/army_work/Ранковий звіт/10.09.2026/Штатка.xlsx"),
  STAFF_READ_OPTIONS,
);

const unitByName = new Map<string, string>();
for (const [fullName, entry] of Object.entries(parseExcelLabState(staff.sheets))) {
  unitByName.set(normalizeLooseText(fullName), entry.unit || "Без підрозділу");
}
const unitOf = (fullName: string) =>
  unitByName.get(normalizeLooseText(fullName)) ?? "Без підрозділу";

type ListedPerson = {
  fullName: string;
  rank: string;
  unit: string;
  date: string;
  note: string;
};

const sergeantsOnly = <T extends { rankGroup: string }>(people: T[]) =>
  people.filter((p) => p.rankGroup === "sergeant");

const formatList = (title: string, people: ListedPerson[]) => {
  console.log(`\n=== ${title} (${people.length}) ===`);
  if (!people.length) {
    console.log("  —");
    return;
  }
  people
    .sort((a, b) => a.fullName.localeCompare(b.fullName, "uk"))
    .forEach((p, index) => {
      console.log(
        `${index + 1}. ${p.fullName} · ${p.rank} · ${p.unit} · ${p.date}${p.note ? ` · ${p.note}` : ""}`,
      );
    });
};

const combat = buildCombatLossesFromPb(pb, { sinceDate: since });
const combatMonth = sergeantsOnly(combat.people).filter((p) =>
  isDateInMonth(p.orderDate, month, year),
);

const combatGroups = {
  "Бойові · безповоротні": [] as ListedPerson[],
  "Бойові · санітарні": [] as ListedPerson[],
  "Бойові · безвісті": [] as ListedPerson[],
  "Бойові · полон": [] as ListedPerson[],
};

for (const person of combatMonth) {
  const bucket = classifySergeantsCombatLoss(
    person.typeOrAbsence,
    person.status,
    person.placeOrNote,
    person.matchNote,
  );
  const label =
    bucket === "combatKilled"
      ? "Бойові · безповоротні"
      : bucket === "combatSanitary"
        ? "Бойові · санітарні"
        : bucket === "combatMissing"
          ? "Бойові · безвісті"
          : "Бойові · полон";
  combatGroups[label].push({
    fullName: person.fullName,
    rank: person.rank,
    unit: unitOf(person.fullName),
    date: person.orderDate,
    note: `${person.reasonLabel} · ${person.placeOrNote || person.typeOrAbsence}`,
  });
}

const departures = buildDeparturesFromEjoos(ejoos, { sinceDate: since });
const departuresMonth = sergeantsOnly(departures.people).filter((p) => {
  const date = p.excludeDate || p.orderDate;
  return isDateInMonth(date, month, year) && p.category !== "transfer";
});

const nonCombatGroups = {
  "Небойові · безповоротні": [] as ListedPerson[],
  "Небойові · санітарні / здоровʼя": [] as ListedPerson[],
};

for (const person of departuresMonth) {
  const text = normalizeLooseText(`${person.ground} ${person.type}`);
  const label = /(поранен|\b300\b|санітар|травм|контуз|здоров)/.test(text)
    ? "Небойові · санітарні / здоровʼя"
    : "Небойові · безповоротні";
  nonCombatGroups[label].push({
    fullName: person.fullName,
    rank: person.rank,
    unit: unitOf(person.fullName),
    date: person.excludeDate || person.orderDate,
    note: `${person.categoryLabel} · ${person.ground}`,
  });
}

const disposition = buildDispositionFromArchive(pb, {
  sinceDate: "01.01.2026",
  openOnly: true,
});
const rozporKeys = buildRozporadzhennyaPersonKeys(pb);
const onTreatment = sergeantsOnly(disposition.people)
  .filter((p) => p.reason === "treatment")
  .filter((p) => !rozporKeys.has(p.personId || normalizeLooseText(p.fullName)))
  .map((p) => ({
    fullName: p.fullName,
    rank: p.rank,
    unit: unitOf(p.fullName),
    date: p.departDate || p.orderDate,
    note: `${p.absenceType} · ${p.place}`,
  }));

const szch = buildSzchFromRuh(pb, { sinceDate: since, untilDate: until });
const szchMonth = sergeantsOnly(szch.people).map((p) => ({
  fullName: p.fullName,
  rank: p.rank,
  unit: unitOf(p.fullName),
  date: p.orderDate,
  note: `${p.type} · ${p.status} · ${p.note || p.destination}`,
}));

console.log(`Перевірка звіту · ${month}.${year} · сержанти/старшини`);

for (const [title, people] of Object.entries(combatGroups)) {
  formatList(title, people);
}
for (const [title, people] of Object.entries(nonCombatGroups)) {
  formatList(title, people);
}
formatList("На лікуванні (відкриті періоди archive)", onTreatment);
formatList("СЗЧ за місяць", szchMonth);
