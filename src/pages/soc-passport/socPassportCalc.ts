import { bucketOf, countsInExitMetrics, countsInNoExitsList, emptyBucketCounts } from "./socPassportFields";
import type {
  BucketCounts,
  PassportMetricId,
  PassportTableRow,
  RankGroup,
  ServiceType,
  SocPassportResult,
  SocPerson,
  SocStaffSlot,
} from "./socPassportTypes";

const add = (counts: BucketCounts, rankGroup: RankGroup, serviceType: ServiceType) => {
  counts[bucketOf(rankGroup, serviceType)] += 1;
};

const totalOf = (counts: BucketCounts) =>
  counts.officerMobilized +
  counts.officerContract +
  counts.sergeantMobilized +
  counts.sergeantContract +
  counts.soldierMobilized +
  counts.soldierContract;

const layout: Array<{
  id: string;
  kind: "section" | "metric";
  number: string;
  label: string;
  metricId?: PassportMetricId;
}> = [
  { id: "staff", kind: "metric", number: "1.", label: "За штатом", metricId: "staff" },
  { id: "listed", kind: "metric", number: "2.", label: "За списком", metricId: "listed" },
  { id: "present", kind: "metric", number: "3.", label: "В наявності", metricId: "present" },
  { id: "disposition", kind: "metric", number: "4.", label: "В розпорядженні", metricId: "disposition" },
  { id: "arrived", kind: "metric", number: "", label: "Прибули до в/ч.", metricId: "arrived" },
  { id: "fromTck", kind: "metric", number: "5.", label: "З ТЦК та СП", metricId: "fromTck" },
  { id: "fromNc", kind: "metric", number: "6.", label: "З НЦ", metricId: "fromTrainingCenter" },
  { id: "fromRecruit", kind: "metric", number: "7.", label: "Прямий рекрутинг", metricId: "fromRecruiting" },
  { id: "fromBrez", kind: "metric", number: "8.", label: "З Брез", metricId: "fromBrez" },
  { id: "sex", kind: "section", number: "", label: "Стать" },
  { id: "male", kind: "metric", number: "9.", label: "Чоловіки", metricId: "male" },
  { id: "female", kind: "metric", number: "10.", label: "Жінки", metricId: "female" },
  { id: "age", kind: "section", number: "", label: "За віком" },
  { id: "age18", kind: "metric", number: "11.", label: "18-25", metricId: "age18_25" },
  { id: "age26", kind: "metric", number: "12", label: "26-30", metricId: "age26_30" },
  { id: "age31", kind: "metric", number: "13", label: "31-40", metricId: "age31_40" },
  { id: "age41", kind: "metric", number: "14", label: "41-50", metricId: "age41_50" },
  { id: "age50", kind: "metric", number: "15", label: "Старші 50", metricId: "age50plus" },
  { id: "nat", kind: "section", number: "", label: "Національність" },
  { id: "natUa", kind: "metric", number: "16", label: "Україна", metricId: "natUkraine" },
  { id: "natPl", kind: "metric", number: "17", label: "Польща", metricId: "natPoland" },
  { id: "natUs", kind: "metric", number: "18", label: "США", metricId: "natUsa" },
  { id: "natGb", kind: "metric", number: "19", label: "Британія", metricId: "natBritain" },
  { id: "natPt", kind: "metric", number: "20", label: "Португалія", metricId: "natPortugal" },
  { id: "natAr", kind: "metric", number: "21", label: "Аргентина", metricId: "natArgentina" },
  { id: "natOther", kind: "metric", number: "22", label: "Інші", metricId: "natOther" },
  { id: "natRu", kind: "metric", number: "23", label: "Рф", metricId: "natRussia" },
  { id: "exits", kind: "section", number: "", label: "По виконанню бойових завдань" },
  { id: "exitsNone", kind: "metric", number: "", label: "Не виконували", metricId: "exitsNone" },
  { id: "exits14", kind: "metric", number: "", label: "До 5ти виходів (завдань на ЛБЗ)", metricId: "exits1_4" },
  { id: "exits510", kind: "metric", number: "", label: "5-10 виходів", metricId: "exits5_10" },
  { id: "exits1115", kind: "metric", number: "", label: "11-15 виходів", metricId: "exits11_15" },
  { id: "exits1620", kind: "metric", number: "", label: "16-20 виходів", metricId: "exits16_20" },
  { id: "exits2125", kind: "metric", number: "", label: "21-25 виходів", metricId: "exits21_25" },
  { id: "exits2630", kind: "metric", number: "", label: "26-30 виходів", metricId: "exits26_30" },
  { id: "exits30", kind: "metric", number: "", label: "30 і більше", metricId: "exits30plus" },
  { id: "regions", kind: "section", number: "", label: "По регіонам" },
  { id: "kyivCity", kind: "metric", number: "", label: "м. Київ", metricId: "kyivCity" },
  { id: "kyiv", kind: "metric", number: "", label: "Київська обл.", metricId: "kyiv" },
  { id: "vinnytsia", kind: "metric", number: "", label: "Вінницька обл.", metricId: "vinnytsia" },
  { id: "volyn", kind: "metric", number: "", label: "Волинська обл.", metricId: "volyn" },
  { id: "dnipro", kind: "metric", number: "", label: "Дніпропетровська обл.", metricId: "dnipro" },
  {
    id: "donetskControlled",
    kind: "metric",
    number: "",
    label: "Донецька обл. (підконтрольна частина)",
    metricId: "donetskControlled",
  },
  {
    id: "donetskOccupied",
    kind: "metric",
    number: "",
    label: "Донецька обл. (тимчасово окупована частина)",
    metricId: "donetskOccupied",
  },
  { id: "zhytomyr", kind: "metric", number: "", label: "Житомирська обл.", metricId: "zhytomyr" },
  { id: "zakarpattia", kind: "metric", number: "", label: "Закарпатська обл.", metricId: "zakarpattia" },
  {
    id: "ivanoFrankivsk",
    kind: "metric",
    number: "",
    label: "Івано-Франківська обл.",
    metricId: "ivanoFrankivsk",
  },
  { id: "kirovohrad", kind: "metric", number: "", label: "Кіровоградська обл.", metricId: "kirovohrad" },
  { id: "luhansk", kind: "metric", number: "", label: "Луганська обл.", metricId: "luhansk" },
  { id: "lviv", kind: "metric", number: "", label: "Львівська обл.", metricId: "lviv" },
  { id: "mykolaiv", kind: "metric", number: "", label: "Миколаївська обл.", metricId: "mykolaiv" },
  { id: "odesa", kind: "metric", number: "", label: "Одеська обл.", metricId: "odesa" },
  { id: "poltava", kind: "metric", number: "", label: "Полтавська обл.", metricId: "poltava" },
  { id: "rivne", kind: "metric", number: "", label: "Рівненська обл.", metricId: "rivne" },
  { id: "sumy", kind: "metric", number: "", label: "Сумська обл.", metricId: "sumy" },
  { id: "ternopil", kind: "metric", number: "", label: "Тернопільська обл.", metricId: "ternopil" },
  {
    id: "zaporizhzhiaControlled",
    kind: "metric",
    number: "",
    label: "Запорізька обл. (підконтрольна частина)",
    metricId: "zaporizhzhiaControlled",
  },
  {
    id: "zaporizhzhiaOccupied",
    kind: "metric",
    number: "",
    label: "Запорізька обл. (не підконтрольна частина)",
    metricId: "zaporizhzhiaOccupied",
  },
  {
    id: "kharkivControlled",
    kind: "metric",
    number: "",
    label: "Харківська обл. (підконтрольна частина)",
    metricId: "kharkivControlled",
  },
  {
    id: "kharkivOccupied",
    kind: "metric",
    number: "",
    label: "Харківська обл. (непідконтрольна частина)",
    metricId: "kharkivOccupied",
  },
  {
    id: "khersonControlled",
    kind: "metric",
    number: "",
    label: "Херсонська обл. (підконтрольна частина)",
    metricId: "khersonControlled",
  },
  {
    id: "khersonOccupied",
    kind: "metric",
    number: "",
    label: "Херсонська обл. (непідконтрольна частина)",
    metricId: "khersonOccupied",
  },
  { id: "khmelnytskyi", kind: "metric", number: "", label: "Хмельницька обл.", metricId: "khmelnytskyi" },
  { id: "cherkasy", kind: "metric", number: "", label: "Черкаська обл.", metricId: "cherkasy" },
  { id: "chernivtsi", kind: "metric", number: "", label: "Чернівецька обл.", metricId: "chernivtsi" },
  { id: "chernihiv", kind: "metric", number: "", label: "Чернігівська обл.", metricId: "chernihiv" },
  { id: "crimea", kind: "metric", number: "", label: "АР Крим", metricId: "crimea" },
  { id: "family", kind: "section", number: "", label: "Сімейний стан" },
  { id: "married", kind: "metric", number: "", label: "Одружені", metricId: "married" },
  { id: "unmarried", kind: "metric", number: "", label: "Неодружені", metricId: "unmarried" },
  { id: "civil", kind: "metric", number: "", label: "Цивільний шлюб", metricId: "civil" },
  { id: "kin", kind: "section", number: "", label: "Родинні зв’язки" },
  { id: "kids", kind: "metric", number: "", label: "Мають дітей до 18 років", metricId: "childrenUnder18" },
  { id: "kids3", kind: "metric", number: "", label: "З них 3 і більше до 18 років", metricId: "children3plus" },
  { id: "relServe", kind: "metric", number: "", label: "Близькі родичі, що служать", metricId: "relativesServing" },
  { id: "relAbroad", kind: "metric", number: "", label: "Близькі родичі за кордоном", metricId: "relativesAbroad" },
  {
    id: "relHostile",
    kind: "metric",
    number: "",
    label: "Близькі родичі в недружніх країнах (рф, Білорусь)",
    metricId: "relativesHostile",
  },
  { id: "status", kind: "section", number: "", label: "Статус" },
  { id: "ubd", kind: "metric", number: "", label: "УБД", metricId: "ubd" },
  { id: "idp", kind: "metric", number: "", label: "ВПО", metricId: "idp" },
];

const bump = (
  metrics: Record<string, BucketCounts>,
  metricId: PassportMetricId,
  person: Pick<SocPerson, "rankGroup" | "serviceType">,
) => {
  const counts = metrics[metricId] ?? emptyBucketCounts();
  add(counts, person.rankGroup, person.serviceType);
  metrics[metricId] = counts;
};

const countStaff = (slots: SocStaffSlot[]) => {
  const counts = emptyBucketCounts();
  for (const slot of slots) {
    add(counts, slot.rankGroup, "mobilized");
  }
  return counts;
};

export const buildSocPassportResult = ({
  people,
  staffSlots,
  sheets,
  warnings = [],
}: {
  people: SocPerson[];
  staffSlots: SocStaffSlot[];
  sheets: SocPassportResult["sheets"];
  warnings?: string[];
}): SocPassportResult => {
  const metrics: Record<string, BucketCounts> = {
    staff: countStaff(staffSlots),
  };

  for (const person of people) {
    if (person.onList) bump(metrics, "listed", person);
    if (person.present) bump(metrics, "present", person);
    if (person.inDisposition) bump(metrics, "disposition", person);
    if (person.arrivalSource !== "unknown") bump(metrics, "arrived", person);
    if (person.arrivalSource === "tck") bump(metrics, "fromTck", person);
    if (person.arrivalSource === "trainingCenter") {
      bump(metrics, "fromTrainingCenter", person);
    }
    if (person.arrivalSource === "recruiting") bump(metrics, "fromRecruiting", person);
    if (person.arrivalSource === "brez") bump(metrics, "fromBrez", person);
    if (person.sex === "female") bump(metrics, "female", person);
    else bump(metrics, "male", person);
    if (person.ageBand === "18-25") bump(metrics, "age18_25", person);
    if (person.ageBand === "26-30") bump(metrics, "age26_30", person);
    if (person.ageBand === "31-40") bump(metrics, "age31_40", person);
    if (person.ageBand === "41-50") bump(metrics, "age41_50", person);
    if (person.ageBand === "50+") bump(metrics, "age50plus", person);
    if (person.nationality === "poland") bump(metrics, "natPoland", person);
    else if (person.nationality === "usa") bump(metrics, "natUsa", person);
    else if (person.nationality === "britain") bump(metrics, "natBritain", person);
    else if (person.nationality === "portugal") bump(metrics, "natPortugal", person);
    else if (person.nationality === "argentina") bump(metrics, "natArgentina", person);
    else if (person.nationality === "other") bump(metrics, "natOther", person);
    else if (person.nationality === "russia") bump(metrics, "natRussia", person);
    else bump(metrics, "natUkraine", person);
    if (
      person.exitBand === "none" &&
      countsInNoExitsList(person) &&
      !person.ubdRosterStatus &&
      !person.staticCombatExitOverride
    ) {
      bump(metrics, "exitsNone", person);
    }
    if (person.exitBand === "1-4" && countsInExitMetrics(person)) {
      bump(metrics, "exits1_4", person);
    }
    if (person.exitBand === "5-10" && countsInExitMetrics(person)) {
      bump(metrics, "exits5_10", person);
    }
    if (person.exitBand === "11-15" && countsInExitMetrics(person)) {
      bump(metrics, "exits11_15", person);
    }
    if (person.exitBand === "16-20" && countsInExitMetrics(person)) {
      bump(metrics, "exits16_20", person);
    }
    if (person.exitBand === "21-25" && countsInExitMetrics(person)) {
      bump(metrics, "exits21_25", person);
    }
    if (person.exitBand === "26-30" && countsInExitMetrics(person)) {
      bump(metrics, "exits26_30", person);
    }
    if (person.exitBand === "30+" && countsInExitMetrics(person)) {
      bump(metrics, "exits30plus", person);
    }
    if (person.region !== "unknown") bump(metrics, person.region, person);
    if (person.marital === "married") bump(metrics, "married", person);
    else if (person.marital === "civil") bump(metrics, "civil", person);
    else bump(metrics, "unmarried", person);
    if (person.childrenUnder18 > 0) bump(metrics, "childrenUnder18", person);
    if (person.children3plus) bump(metrics, "children3plus", person);
    if (person.relativesServing) bump(metrics, "relativesServing", person);
    if (person.relativesAbroad) bump(metrics, "relativesAbroad", person);
    if (person.relativesHostile) bump(metrics, "relativesHostile", person);
    if (person.hasUbd || person.ubdRosterStatus) bump(metrics, "ubd", person);
    if (person.isIdp) bump(metrics, "idp", person);
  }

  const rows: PassportTableRow[] = layout.map((row) => {
    if (row.kind === "section" || !row.metricId) {
      return {
        id: row.id,
        kind: "section",
        number: row.number,
        label: row.label,
      };
    }
    const counts = metrics[row.metricId] ?? emptyBucketCounts();
    return {
      id: row.id,
      kind: "metric",
      number: row.number,
      label: row.label,
      metricId: row.metricId,
      counts,
      total: totalOf(counts),
    };
  });

  return {
    people,
    staffSlots,
    rows,
    summary: {
      staffSlots: staffSlots.length,
      occupied: people.length,
      vacant: staffSlots.filter((slot) => !slot.occupied).length,
      oosMatched: people.filter((person) => person.match.oos).length,
      morningMatched: people.filter((person) => person.match.morning).length,
      exitsMatched: people.filter(
        (person) =>
          person.morningExitCount > 0 ||
          person.match.jbdExits ||
          person.match.bplaExits,
      ).length,
      combatDutyMatched: people.filter((person) => person.combatDutyEvidence.length > 0)
        .length,
      relativesParsed: people.filter((person) => person.relativesRaw).length,
      unknownRegion: people.filter((person) => person.region === "unknown").length,
      unknownAge: people.filter((person) => person.age == null).length,
    },
    warnings,
    sheets,
  };
};
