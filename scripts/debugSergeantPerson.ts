/**
 * npx jiti scripts/debugSergeantPerson.ts "ПАРУБЕЦЬ"
 */
import { readFileSync } from "node:fs";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../src/excelRoundTrip";
import {
  buildCombatLossesFromPb,
  buildDispositionFromArchive,
  buildSzchFromRuh,
} from "../src/pages/soc-passport/socPassportDepartures";
import { normalizeLooseText } from "../src/pages/soc-passport/socPassportFields";
import {
  parsePbArchive,
  parsePbMovements,
} from "../src/pages/ejournal/ejoosSyncPlan";

const query = normalizeLooseText(process.argv[2] || "ПАРУБЕЦЬ");
const pbPath =
  "/Volumes/KINGSTON/army_work/ЕЖООС/08.09.2026/1ПБ_07092026.xlsx";

const asFile = (path: string) => {
  const buffer = readFileSync(path) as unknown as File & { name: string };
  Object.defineProperty(buffer, "name", {
    value: path.split("/").pop() ?? path,
  });
  return buffer;
};

const pb = await readWorkbookSnapshot(asFile(pbPath), EJOOS_SYNC_READ_OPTIONS);
const matches = (name: string) => normalizeLooseText(name).includes(query);

console.log("Query:", query);
console.log("\n--- all archive rows ---");
for (const row of parsePbArchive(pb)) {
  if (!matches(row.fullName)) continue;
  console.log(row);
}

console.log("\n--- рух ---");
for (const row of parsePbMovements(pb)) {
  if (!matches(row.fullName)) continue;
  console.log(row);
}

console.log("\n--- buildDispositionFromArchive ---");
const disposition = buildDispositionFromArchive(pb, {
  sinceDate: "01.01.2026",
  openOnly: true,
});
for (const p of disposition.people.filter((x) => matches(x.fullName))) {
  console.log({
    fullName: p.fullName,
    rank: p.rank,
    rankGroup: p.rankGroup,
    reason: p.reason,
    reasonLabel: p.reasonLabel,
    absenceType: p.absenceType,
    place: p.place,
    departDate: p.departDate,
    returnDate: p.returnDate,
    matchNote: p.matchNote,
  });
}

console.log("\n--- buildCombatLossesFromPb ---");
const combat = buildCombatLossesFromPb(pb, { sinceDate: "01.01.2026" });
for (const p of combat.people.filter((x) => matches(x.fullName))) {
  console.log({
    fullName: p.fullName,
    rank: p.rank,
    rankGroup: p.rankGroup,
    reason: p.reason,
    reasonLabel: p.reasonLabel,
    orderDate: p.orderDate,
    typeOrAbsence: p.typeOrAbsence,
    placeOrNote: p.placeOrNote,
    matchNote: p.matchNote,
    sourceSheet: p.sourceSheet,
  });
}

console.log("\n--- buildSzchFromRuh ---");
const szch = buildSzchFromRuh(pb, { sinceDate: "01.01.2026" });
for (const p of szch.people.filter((x) => matches(x.fullName))) {
  console.log(p);
}
