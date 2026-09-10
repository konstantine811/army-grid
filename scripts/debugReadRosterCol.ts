import { readRosterColumnValue } from "../src/pages/excel-fill/rosterSourceSnapshot";
import { buildStaffSheetColumnsRecord } from "../src/pages/overview/overviewStaffSheetColumns";

const row = {
  column_31: "ППД Вишневе",
  column_14: "ГУК",
  місце_перебування: "ППД Вишневе",
};

console.log("31", readRosterColumnValue(row, 31));
console.log("35", readRosterColumnValue(row, 35));
console.log("record", buildStaffSheetColumnsRecord(row));
