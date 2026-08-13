import {
  type ExcelRow,
  type ExcelWorkbookSnapshot,
  getColumnHeader,
  hasRowData,
  valueToDisplay,
} from "../../excelRoundTrip";

export type AnalyticsMetric = {
  label: string;
  value: number;
  tone?: "good" | "warn" | "bad";
  detail?: string;
};

export type AnalyticsRecord = {
  excelRowNumber: number;
  roster: string;
  rank: string;
  name: string;
  callSign: string;
  status: string;
  type: string;
  bg: string;
  bzvp: string;
  location: string;
  position: string;
  reason: string;
};

export type AnalyticsData = {
  sourceFileName: string;
  reportDate: string;
  sheet: ExcelWorkbookSnapshot;
  rows: ExcelRow[];
  metrics: {
    staff: number;
    listed: number;
    absent: number;
    inRanks: number;
    understaff: number;
    understaffPercent: number;
    staffedPercent: number;
    balanceDiff: number;
    newcomersNoRole: number;
  };
  absenceReasons: AnalyticsMetric[];
  detachedDetails: AnalyticsMetric[];
  attachedDetails: AnalyticsMetric[];
  structure: AnalyticsMetric[];
  managementDetails: AnalyticsMetric[];
  supportDetails: AnalyticsMetric[];
  execution: AnalyticsMetric[];
  executionDetails: AnalyticsMetric[];
  newcomerDetails: AnalyticsMetric[];
  bzvp: AnalyticsMetric[];
  ageGroups: AnalyticsMetric[];
  hintManagementDetails: AnalyticsMetric[];
  hintSupportDetails: AnalyticsMetric[];
  hintHealingDetails: AnalyticsMetric[];
  hintInfantryDetails: AnalyticsMetric[];
  hintPpdDetails: AnalyticsMetric[];
  hintRpakDetails: AnalyticsMetric[];
  unclassifiedInRanks: AnalyticsRecord[];
  quality: AnalyticsMetric[];
};

export const columnValue = (row: ExcelRow, index: number) =>
  valueToDisplay(row.values[index]).trim();
export const normalized = (value: string) =>
  value.replace(/\s+/g, " ").trim().toLowerCase();
export const isBlank = (value: string) => normalized(value) === "";
export const equalsText = (value: string, expected: string) =>
  normalized(value) === normalized(expected);
export const includesText = (value: string, expected: string) =>
  normalized(value).includes(normalized(expected));
export const toNumber = (value: string) => {
  const cleanedValue = value.replace(",", ".").replace(/[^\d.-]/g, "");
  if (!cleanedValue) return null;

  const number = Number(cleanedValue);

  return Number.isFinite(number) ? number : null;
};
export const toAge = (value: string) => {
  const number = toNumber(value);

  return number !== null && number >= 16 && number <= 80 ? number : null;
};

export const countRows = (rows: ExcelRow[], predicate: (row: ExcelRow) => boolean) =>
  rows.filter(predicate).length;
export const extractCourseDate = (value: string) => {
  const match = value.match(/\b(\d{1,2})[./](\d{1,2})(?:[./](\d{2,4}))?\b/);
  if (!match) return null;

  const day = match[1].padStart(2, "0");
  const month = match[2].padStart(2, "0");
  const year = match[3] ? match[3].slice(-2) : "";

  return {
    label: `${day}.${month}`,
    sortKey: `${year || "99"}-${month}-${day}`,
  };
};

export const makeAnalyticsData = (
  snapshot: ExcelWorkbookSnapshot,
): AnalyticsData | null => {
  const sourceSheet = snapshot.sheets.find(
    (sheet) => sheet.sheetName === "Аркуш1" || sheet.sheetName === "Аркуш",
  );
  if (!sourceSheet) return null;

  const sheet = {
    ...snapshot,
    sheetName: sourceSheet.sheetName,
    headerRows: sourceSheet.headerRows,
    rows: sourceSheet.rows,
    columnCount: sourceSheet.columnCount,
    columnIndexes: sourceSheet.columnIndexes,
    dataStartRow: sourceSheet.dataStartRow,
  };
  const rows = sourceSheet.rows.filter((row) => hasRowData(row.values));
  const excelColumn = (originalIndex: number) => {
    const compactIndex = sourceSheet.columnIndexes.indexOf(originalIndex);

    return compactIndex >= 0 ? compactIndex : originalIndex;
  };
  const findColumnByHeader = (header: string) => {
    const normalizedHeader = normalized(header);
    const index = Array.from(
      { length: sourceSheet.columnCount },
      (_, columnIndex) => columnIndex,
    ).find((columnIndex) =>
      [
        ...sourceSheet.headerRows,
        ...sourceSheet.rawRows.slice(0, sourceSheet.dataStartRow),
      ]
        .map((row) => normalized(valueToDisplay(row[columnIndex])))
        .some((cellHeader) => cellHeader.includes(normalizedHeader)),
    );

    return index ?? -1;
  };
  const getColumnAgeValues = (columnIndex: number) =>
    columnIndex >= 0
      ? rows
          .map((row) => toAge(columnValue(row, columnIndex)))
          .filter((value): value is number => value !== null)
      : [];
  const getColumnAgeDiagnostics = (columnIndex: number) =>
    columnIndex >= 0
      ? rows.map((row) => {
          const rawValue = columnValue(row, columnIndex);
          const parsedValue = toAge(rawValue);

          return {
            excelRowNumber: row.excelRowNumber,
            rawValue,
            parsedValue,
            accepted: parsedValue !== null,
          };
        })
      : [];
  const findAgeColumn = () => {
    const headerColumn = findColumnByHeader("Повних років");
    const fallbackColumn = excelColumn(17);
    const candidates = Array.from(
      new Set(
        [headerColumn, fallbackColumn].filter(
          (columnIndex) => columnIndex >= 0,
        ),
      ),
    )
      .map((columnIndex) => ({
        columnIndex,
        ages: getColumnAgeValues(columnIndex),
      }))
      .sort((first, second) => second.ages.length - first.ages.length);

    return candidates[0]?.ages.length ? candidates[0].columnIndex : -1;
  };
  const c = {
    unit: excelColumn(0),
    roster: excelColumn(1),
    position: excelColumn(4),
    rank: excelColumn(12),
    person: excelColumn(13),
    callSign: excelColumn(14),
    age: findAgeColumn(),
    staffMarker: excelColumn(8),
    status: excelColumn(20),
    type: excelColumn(21),
    bg: excelColumn(22),
    bzvp: excelColumn(23),
    course: excelColumn(25),
    restriction: excelColumn(27),
    externalUnit: excelColumn(28),
    location: excelColumn(30),
  };
  const nova = (row: ExcelRow) => equalsText(columnValue(row, c.unit), "нова");
  const statusIs = (row: ExcelRow, text: string) =>
    equalsText(columnValue(row, c.status), text);
  const statusHas = (row: ExcelRow, text: string) =>
    includesText(columnValue(row, c.status), text);
  const typeHas = (row: ExcelRow, text: string) =>
    includesText(columnValue(row, c.type), text);
  const bgIs = (row: ExcelRow, text: string) =>
    equalsText(columnValue(row, c.bg), text);
  const bgHas = (row: ExcelRow, text: string) =>
    includesText(columnValue(row, c.bg), text);
  const locationHas = (row: ExcelRow, text: string) =>
    includesText(columnValue(row, c.location), text);
  const externalIs = (row: ExcelRow, text: string) =>
    equalsText(columnValue(row, c.externalUnit), text);
  const externalHas = (row: ExcelRow, text: string) =>
    includesText(columnValue(row, c.externalUnit), text);
  const inRanksStatus = (row: ExcelRow) => statusHas(row, "в строю");
  const newcomer = (row: ExcelRow) => statusHas(row, "Новоприбулий");
  const activeOrNew = (row: ExcelRow) => inRanksStatus(row) || newcomer(row);

  const staff = countRows(
    rows,
    (row) => nova(row) && !isBlank(columnValue(row, c.staffMarker)),
  );
  const listed = countRows(
    rows,
    (row) => nova(row) && !isBlank(columnValue(row, c.status)),
  );
  const healing = countRows(
    rows,
    (row) =>
      nova(row) &&
      statusIs(row, "Лікування") &&
      !isBlank(columnValue(row, c.location)),
  );
  const missing = countRows(
    rows,
    (row) => nova(row) && statusIs(row, "Зниклі безвісти"),
  );
  const deceased = countRows(
    rows,
    (row) => nova(row) && statusIs(row, "Загиблі"),
  );
  const awol = countRows(rows, (row) => nova(row) && statusIs(row, "СЗЧ"));
  const notInRanks = countRows(
    rows,
    (row) => nova(row) && statusIs(row, "Не в сторою"),
  );
  const detached = countRows(
    rows,
    (row) => nova(row) && statusHas(row, "Відком. за межі ПБ"),
  );
  const vacation = countRows(
    rows,
    (row) => nova(row) && statusIs(row, "Відпустка"),
  );
  const training = countRows(
    rows,
    (row) => nova(row) && statusIs(row, "Від-ні навчання"),
  );
  const businessTrip = countRows(
    rows,
    (row) => nova(row) && statusIs(row, "Відрядження"),
  );
  const absent =
    healing +
    missing +
    deceased +
    awol +
    notInRanks +
    detached +
    vacation +
    training +
    businessTrip;
  const inRanks = countRows(rows, (row) => nova(row) && inRanksStatus(row));
  const understaff = staff - listed;
  const staffedPercent = staff ? (inRanks / staff) * 100 : 0;
  const understaffPercent = staff ? (understaff / staff) * 100 : 0;
  const balanceDiff = listed - absent - inRanks;
  const detachedDetails = [
    {
      label: "2ПБ",
      value: countRows(
        rows,
        (row) =>
          nova(row) &&
          statusHas(row, "Відком. за межі ПБ") &&
          externalIs(row, "2ПБ"),
      ),
    },
    {
      label: "2ББС",
      value: countRows(
        rows,
        (row) =>
          nova(row) &&
          statusHas(row, "Відком. за межі ПБ") &&
          externalIs(row, "2ББС"),
      ),
    },
    {
      label: "Знахарь",
      value: countRows(
        rows,
        (row) =>
          nova(row) &&
          statusHas(row, "Відком. за межі ПБ") &&
          externalIs(row, "Знахарь"),
      ),
    },
    {
      label: "ППО",
      value: countRows(
        rows,
        (row) =>
          nova(row) &&
          statusHas(row, "Відком. за межі ПБ") &&
          externalIs(row, "ППО"),
      ),
    },
    {
      label: "Рота РЕБ",
      value: countRows(
        rows,
        (row) =>
          nova(row) &&
          statusHas(row, "Відком. за межі ПБ") &&
          externalIs(row, "РРЕБ"),
      ),
    },
    {
      label: "Полігон Б",
      value: countRows(
        rows,
        (row) =>
          nova(row) &&
          statusHas(row, "Відком. за межі ПБ") &&
          externalHas(row, "Полігон Б"),
      ),
    },
    {
      label: "БМЗ",
      value: countRows(
        rows,
        (row) =>
          nova(row) &&
          statusHas(row, "Відком. за межі ПБ") &&
          externalHas(row, "БМЗ"),
      ),
    },
    {
      label: "Шквал",
      value: countRows(
        rows,
        (row) =>
          nova(row) &&
          statusHas(row, "Відком. за межі ПБ") &&
          externalIs(row, "Шквал"),
      ),
    },
    {
      label: "Транзит/Полк",
      value: countRows(
        rows,
        (row) =>
          nova(row) &&
          statusHas(row, "Відком. за межі ПБ") &&
          (externalIs(row, "ПОЛК") || externalIs(row, "ТРАНЗИТЕР")),
      ),
      tone: "warn" as const,
    },
  ];
  const attached210 = countRows(
    rows,
    (row) => equalsText(columnValue(row, c.unit), "210") && inRanksStatus(row),
  );
  const attachedBrez = countRows(
    rows,
    (row) =>
      equalsText(columnValue(row, c.roster), "БРЕЗ") && inRanksStatus(row),
  );
  const attached41 = countRows(
    rows,
    (row) =>
      includesText(columnValue(row, c.unit), "41ОМБр") && inRanksStatus(row),
  );
  const attachedShkval = countRows(
    rows,
    (row) =>
      includesText(columnValue(row, c.unit), "ШКВАЛ") && inRanksStatus(row),
  );
  const attachedOther =
    countRows(
      rows,
      (row) =>
        equalsText(columnValue(row, c.roster), "БРЕЗ") && inRanksStatus(row),
    ) +
    countRows(
      rows,
      (row) =>
        equalsText(columnValue(row, c.roster), "ПРИКОМАНДИРОВАНІ") &&
        inRanksStatus(row),
    ) -
    attached210 -
    attachedBrez -
    attached41 -
    attachedShkval;
  const attachedDetails = [
    { label: "210", value: attached210 },
    { label: "БРЕЗ", value: attachedBrez },
    { label: "41 ОМБр", value: attached41 },
    { label: "ШКВАЛ", value: attachedShkval },
    { label: "інші", value: attachedOther },
  ];

  const managementDetails = [
    {
      label: "Управління Бату",
      value: countRows(
        rows,
        (row) => nova(row) && inRanksStatus(row) && typeHas(row, "Упр. бату"),
      ),
    },
    {
      label: "Штаб Бату",
      value: countRows(
        rows,
        (row) => nova(row) && inRanksStatus(row) && typeHas(row, "Штаб. бату"),
      ),
    },
    {
      label: "Управління підрозділу",
      value: countRows(
        rows,
        (row) =>
          nova(row) && inRanksStatus(row) && typeHas(row, "Упр. підрозділ."),
      ),
    },
    {
      label: "Управління взводу",
      value: countRows(
        rows,
        (row) => nova(row) && inRanksStatus(row) && typeHas(row, "Упр. взводу"),
      ),
    },
  ];
  const supportDetails = [
    {
      label: "Водії/навідники",
      value: countRows(
        rows,
        (row) => nova(row) && inRanksStatus(row) && typeHas(row, "Водій"),
      ),
    },
    {
      label: "Забезпечення підрозділів",
      value: countRows(
        rows,
        (row) =>
          nova(row) &&
          inRanksStatus(row) &&
          ["Забезпечення", "Медики", "Інструктор", "Кухар", "Майстер"].some(
            (item) => typeHas(row, item),
          ),
      ),
    },
    {
      label: "Охорона",
      value: countRows(
        rows,
        (row) => nova(row) && inRanksStatus(row) && typeHas(row, "Охорона"),
      ),
    },
  ];
  const management = managementDetails.reduce(
    (sum, item) => sum + item.value,
    0,
  );
  const support = supportDetails.reduce((sum, item) => sum + item.value, 0);
  const rpakBase = countRows(
    rows,
    (row) => nova(row) && inRanksStatus(row) && typeHas(row, "Пілот"),
  );
  const rpakPpdBusy = countRows(
    rows,
    (row) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Пілот") &&
      locationHas(row, "ППД") &&
      includesText(columnValue(row, c.bzvp), "БЗВП"),
  );
  const rpakExecution = countRows(
    rows,
    (row) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Пілот") &&
      locationHas(row, "на виконанні"),
  );
  const rpak = Math.max(0, rpakBase - rpakPpdBusy - rpakExecution);
  const bgInfantry = countRows(
    rows,
    (row) => nova(row) && inRanksStatus(row) && bgIs(row, "БГ"),
  );
  const notBg =
    countRows(
      rows,
      (row) => nova(row) && inRanksStatus(row) && bgHas(row, "БГ"),
    ) - bgInfantry;

  const infantryExecutionTotal = countRows(
    rows,
    (row) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Піхота") &&
      locationHas(row, "на виконанні"),
  );
  const communication = countRows(
    rows,
    (row) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Піхота") &&
      locationHas(row, "на виконанні") &&
      externalHas(row, "Зв'язок"),
  );
  const infantryExecution = infantryExecutionTotal - communication;
  const grenadeExecution = countRows(
    rows,
    (row) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Гранатометники") &&
      locationHas(row, "на виконанні"),
  );
  const medicsExecution = countRows(
    rows,
    (row) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Медики") &&
      locationHas(row, "на виконанні"),
  );
  const attachedRpakExecution = countRows(
    rows,
    (row) =>
      (equalsText(columnValue(row, c.roster), "ПРИКОМАНДИРОВАНІ") ||
        equalsText(columnValue(row, c.roster), "БРЕЗ")) &&
      inRanksStatus(row) &&
      typeHas(row, "Пілот") &&
      locationHas(row, "на виконанні"),
  );
  const infantryExtraExecution = countRows(
    rows,
    (row) =>
      (equalsText(columnValue(row, c.roster), "ПРИКОМАНДИРОВАНІ") ||
        equalsText(columnValue(row, c.roster), "БРЕЗ")) &&
      inRanksStatus(row) &&
      typeHas(row, "Піхота") &&
      locationHas(row, "на виконанні"),
  );
  const executionTotal =
    infantryExecution +
    rpakExecution +
    attachedRpakExecution +
    communication +
    grenadeExecution +
    medicsExecution;

  const hasNoBzvp = (row: ExcelRow) =>
    includesText(columnValue(row, c.bzvp), "Без БЗВП");
  const bzvpRows = rows.filter((row) => activeOrNew(row) && hasNoBzvp(row));
  const noCourse = countRows(bzvpRows, (row) =>
    includesText(columnValue(row, c.course), "не поданий"),
  );
  const formulaNoBzvpInRanks = (row: ExcelRow) =>
    nova(row) &&
    inRanksStatus(row) &&
    hasNoBzvp(row) &&
    (equalsText(columnValue(row, c.type), "Піхота") ||
      typeHas(row, "Пілот") ||
      equalsText(columnValue(row, c.type), "Гранатометники")) &&
    !(!isBlank(columnValue(row, c.type)) && locationHas(row, "на виконанні"));
  const formulaInRanksCategories = [
    (row: ExcelRow) =>
      nova(row) &&
      inRanksStatus(row) &&
      ["Упр. бату", "Штаб. бату", "Упр. підрозділ.", "Упр. взводу"].some(
        (item) => typeHas(row, item),
      ),
    (row: ExcelRow) =>
      nova(row) &&
      inRanksStatus(row) &&
      (typeHas(row, "Водій") ||
        ["Забезпечення", "Медики", "Інструктор", "Кухар", "Майстер"].some(
          (item) => typeHas(row, item),
        ) ||
        typeHas(row, "Охорона")),
    (row: ExcelRow) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Пілот") &&
      !(locationHas(row, "ППД") && includesText(columnValue(row, c.bzvp), "БЗВП")) &&
      !locationHas(row, "на виконанні"),
    (row: ExcelRow) => nova(row) && inRanksStatus(row) && bgIs(row, "БГ"),
    (row: ExcelRow) =>
      nova(row) && inRanksStatus(row) && bgHas(row, "БГ") && !bgIs(row, "БГ"),
    formulaNoBzvpInRanks,
    (row: ExcelRow) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Піхота") &&
      locationHas(row, "на виконанні") &&
      !externalHas(row, "Зв'язок"),
    (row: ExcelRow) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Пілот") &&
      locationHas(row, "на виконанні"),
    (row: ExcelRow) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Гранатометники") &&
      locationHas(row, "на виконанні"),
    (row: ExcelRow) =>
      nova(row) &&
      inRanksStatus(row) &&
      typeHas(row, "Піхота") &&
      locationHas(row, "на виконанні") &&
      externalHas(row, "Зв'язок"),
  ];
  const unclassifiedInRanks = rows
    .filter(
      (row) =>
        nova(row) &&
        inRanksStatus(row) &&
        !formulaInRanksCategories.some((predicate) => predicate(row)),
    )
    .map((row) => ({
      excelRowNumber: row.excelRowNumber,
      roster: columnValue(row, c.roster),
      rank: columnValue(row, c.rank),
      name: columnValue(row, c.person),
      callSign: columnValue(row, c.callSign),
      status: columnValue(row, c.status),
      type: columnValue(row, c.type),
      bg: columnValue(row, c.bg),
      bzvp: columnValue(row, c.bzvp),
      location: columnValue(row, c.location),
      position: columnValue(row, c.position),
      reason:
        "В строю, але немає БГ/тимчасово не БГ, БЗВП/БРЕЗ, виконання, управління, забезпечення або РБпАК.",
    }));
  const bzvpByDate = Array.from(
    bzvpRows.reduce((dateMap, row) => {
      const courseDate = extractCourseDate(columnValue(row, c.course));
      if (!courseDate) return dateMap;

      const current = dateMap.get(courseDate.label) ?? {
        label: courseDate.label,
        sortKey: courseDate.sortKey,
        value: 0,
      };
      dateMap.set(courseDate.label, { ...current, value: current.value + 1 });

      return dateMap;
    }, new Map<string, { label: string; sortKey: string; value: number }>()),
  )
    .map(([, item]) => item)
    .sort((first, second) => first.sortKey.localeCompare(second.sortKey));
  const newcomersNoRole = countRows(
    rows,
    (row) =>
      newcomer(row) &&
      (equalsText(columnValue(row, c.bzvp), "") ||
        includesText(columnValue(row, c.bzvp), "БРЕЗ")),
  );
  const newcomerDetails = [
    {
      label: "Без БЗВП",
      value: countRows(
        rows,
        (row) =>
          newcomer(row) && equalsText(columnValue(row, c.bzvp), "Без БЗВП"),
      ),
    },
    {
      label: "БРЕЗ/Без БЗВП",
      value: countRows(
        rows,
        (row) =>
          newcomer(row) &&
          equalsText(columnValue(row, c.bzvp), "Без БЗВП/БРЕЗ"),
      ),
    },
    {
      label: "БРЕЗ",
      value: countRows(
        rows,
        (row) => newcomer(row) && equalsText(columnValue(row, c.bzvp), "БРЕЗ"),
      ),
    },
    {
      label: "не БРЕЗ",
      value: countRows(
        rows,
        (row) => newcomer(row) && isBlank(columnValue(row, c.bzvp)),
      ),
    },
  ];
  const hintManagementDetails = [
    {
      label: "Управління батальйону",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Упр. бату"),
      ),
    },
    {
      label: "Штаб батальйону",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Штаб. бату"),
      ),
    },
    {
      label: "Управління підрозділу",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Упр. підрозділ."),
      ),
    },
    {
      label: "Управління взводу",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Упр. взводу"),
      ),
    },
  ];
  const hintSupportDetails = [
    {
      label: "Водії (тил)",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Водій (тил)"),
      ),
    },
    {
      label: "Водії (виконання)",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Водій (викон"),
      ),
    },
    {
      label: "Охорона",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Охорона"),
      ),
    },
    {
      label: "Медики",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Медики"),
      ),
    },
    {
      label: "Кухарі",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Кухар"),
      ),
    },
    {
      label: "Майстри",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Майстер"),
      ),
    },
    {
      label: "Інструктори",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Інструктор"),
      ),
    },
    {
      label: "Інші",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Забезпечення"),
      ),
    },
  ];
  const hintHealingDetails = [
    {
      label: "Тимчасово не БГ",
      value:
        countRows(rows, (row) => activeOrNew(row) && bgHas(row, "БГ")) -
        countRows(rows, (row) => activeOrNew(row) && bgIs(row, "БГ")),
    },
    {
      label: "Шпиталь",
      value: countRows(
        rows,
        (row) =>
          statusIs(row, "Лікування") &&
          equalsText(columnValue(row, c.location), "Шпиталь"),
      ),
    },
    {
      label: "Мед.рота",
      value: countRows(
        rows,
        (row) =>
          statusIs(row, "Лікування") &&
          equalsText(columnValue(row, c.location), "Мед. рота"),
      ),
    },
    {
      label: "Лікувальна відпустка",
      value: countRows(
        rows,
        (row) =>
          statusIs(row, "Лікування") &&
          equalsText(columnValue(row, c.location), "Відпустка лікув."),
      ),
    },
  ];
  const hintInfantryNoBzvpLocations = [
    {
      label: "ППД Вишневе",
      value: countRows(
        rows,
        (row) =>
          activeOrNew(row) &&
          typeHas(row, "Піхота") &&
          !isBlank(columnValue(row, c.bzvp)) &&
          isBlank(columnValue(row, c.bg)) &&
          locationHas(row, "ППД Вишневе"),
      ),
    },
    {
      label: "ППД Слов'янка 13",
      value: countRows(
        rows,
        (row) =>
          activeOrNew(row) &&
          typeHas(row, "Піхота") &&
          !isBlank(columnValue(row, c.bzvp)) &&
          isBlank(columnValue(row, c.bg)) &&
          locationHas(row, "ППД Слов'янка"),
      ),
    },
    {
      label: "КСП Павлоград",
      value: countRows(
        rows,
        (row) =>
          activeOrNew(row) &&
          typeHas(row, "Піхота") &&
          !isBlank(columnValue(row, c.bzvp)) &&
          isBlank(columnValue(row, c.bg)) &&
          locationHas(row, "КСП Павлоград"),
      ),
    },
  ];
  const hintInfantryDetails = [
    {
      label: "На виконанні",
      value: countRows(
        rows,
        (row) =>
          activeOrNew(row) &&
          typeHas(row, "Піхота") &&
          equalsText(columnValue(row, c.location), "На виконанні"),
      ),
    },
    {
      label: "БГ",
      value: countRows(
        rows,
        (row) => activeOrNew(row) && typeHas(row, "Піхота") && bgIs(row, "БГ"),
      ),
    },
    {
      label: "Без БЗВП/БРЕЗ",
      value: hintInfantryNoBzvpLocations.reduce(
        (sum, item) => sum + item.value,
        0,
      ),
      detail: hintInfantryNoBzvpLocations
        .map((item) => `${item.label}: ${item.value}`)
        .join(" · "),
    },
  ];
  const ppdManagement = countRows(
    rows,
    (row) =>
      activeOrNew(row) &&
      locationHas(row, "ППД Вишневе") &&
      ["Упр. бату", "Штаб. бату", "Упр. підрозділ.", "Упр. взводу"].some(
        (item) => typeHas(row, item),
      ),
  );
  const ppdInfantryRestricted = countRows(
    rows,
    (row) =>
      activeOrNew(row) &&
      typeHas(row, "Піхота") &&
      !isBlank(columnValue(row, c.restriction)) &&
      locationHas(row, "ППД Вишневе"),
  );
  const ppdInfantry =
    countRows(
      rows,
      (row) =>
        activeOrNew(row) &&
        typeHas(row, "Піхота") &&
        locationHas(row, "ППД Вишневе"),
    ) - ppdInfantryRestricted;
  const ppdPilots = countRows(
    rows,
    (row) =>
      activeOrNew(row) &&
      typeHas(row, "Пілот") &&
      locationHas(row, "ППД Вишневе") &&
      isBlank(columnValue(row, c.restriction)),
  );
  const ppdGrenadeRestricted = countRows(
    rows,
    (row) =>
      activeOrNew(row) &&
      typeHas(row, "Гранатометники") &&
      !isBlank(columnValue(row, c.restriction)) &&
      locationHas(row, "ППД Вишневе"),
  );
  const ppdGrenade =
    countRows(
      rows,
      (row) =>
        activeOrNew(row) &&
        typeHas(row, "Гранатометники") &&
        locationHas(row, "ППД Вишневе"),
    ) - ppdGrenadeRestricted;
  const ppdRestricted =
    ppdInfantryRestricted +
    countRows(
      rows,
      (row) =>
        activeOrNew(row) &&
        typeHas(row, "Пілот") &&
        !isBlank(columnValue(row, c.restriction)) &&
        locationHas(row, "ППД Вишневе"),
    ) +
    ppdGrenadeRestricted;
  const ppdTotal = countRows(
    rows,
    (row) => activeOrNew(row) && locationHas(row, "ППД Вишневе"),
  );
  const hintPpdDetails = [
    { label: "Управління", value: ppdManagement },
    {
      label: "Забезпечення",
      value:
        ppdTotal -
        ppdManagement -
        ppdInfantry -
        ppdPilots -
        ppdGrenade -
        ppdRestricted,
    },
    { label: "Піхота", value: ppdInfantry },
    { label: "Пілоти", value: ppdPilots },
    { label: "Гранатометники", value: ppdGrenade },
    {
      label: "Призупинка/списання/обмеження",
      value: ppdRestricted,
      tone: "warn" as const,
    },
  ];
  const hintSheet = snapshot.sheets.find(
    (sheet) => sheet.sheetName === "Підказка",
  );
  const hintValueByLabel = (label: string) => {
    const row = hintSheet?.rawRows.find((rawRow) =>
      rawRow.some((cell) => includesText(valueToDisplay(cell), label)),
    );
    const numericValue = row?.find((cell) => typeof cell === "number");

    return typeof numericValue === "number" ? numericValue : 0;
  };
  const hintRpakDetails = [
    { label: "FPV (позиції)", value: hintValueByLabel("FPV") },
    { label: "Бомбери (позиції)", value: hintValueByLabel("Бомбери") },
    { label: "Розвідка (позиції)", value: hintValueByLabel("Розвідка") },
  ];

  const absenceReasons = [
    {
      label: "Лікування",
      value: healing,
      detail: `Шпиталь ${countRows(rows, (row) => statusIs(row, "Лікування") && equalsText(columnValue(row, c.location), "Шпиталь"))} · Медрота ${countRows(rows, (row) => statusIs(row, "Лікування") && equalsText(columnValue(row, c.location), "Мед. рота"))} · Лік. відп. ${countRows(rows, (row) => statusIs(row, "Лікування") && equalsText(columnValue(row, c.location), "Відпустка лікув."))}`,
    },
    { label: "Відкомандировані", value: detached },
    { label: "Зниклі безвісти", value: missing },
    { label: "СЗЧ", value: awol },
    { label: "Загиблі", value: deceased, tone: "bad" as const },
    { label: "Відпустка", value: vacation },
    { label: "Не в строю", value: notInRanks },
    { label: "Навчання", value: training },
    { label: "Відрядження", value: businessTrip },
  ];
  const structure = [
    { label: "Управління", value: management },
    { label: "Забезпечення", value: support },
    { label: "РБпАК", value: rpak },
    { label: "БГ піхота", value: bgInfantry },
    { label: "Тимчасово не БГ", value: notBg },
    {
      label: "Не класифіковано",
      value: unclassifiedInRanks.length,
      tone: unclassifiedInRanks.length ? ("warn" as const) : undefined,
      detail: unclassifiedInRanks.length
        ? unclassifiedInRanks
            .map((record) => `${record.excelRowNumber}: ${record.name}`)
            .join(" · ")
        : undefined,
    },
  ];
  const execution = [
    { label: "Піхота", value: infantryExecution },
    { label: "РБпАК", value: rpakExecution + attachedRpakExecution },
    { label: "Зв'язок", value: communication },
    { label: "Розрахунок МК", value: grenadeExecution },
    { label: "Медики", value: medicsExecution },
  ];
  const executionDetails = [
    { label: "Піхота штат", value: infantryExecution },
    { label: "Піхота позаштат/БРЕЗ", value: infantryExtraExecution },
    { label: "РБпАК штат", value: rpakExecution },
    { label: "РБпАК позаштат", value: attachedRpakExecution },
    {
      label: "Разом контроль",
      value:
        infantryExecution +
        infantryExtraExecution +
        rpakExecution +
        attachedRpakExecution +
        communication +
        grenadeExecution +
        medicsExecution,
    },
  ];
  const bzvp = [
    ...bzvpByDate.map((item) => ({
      label: `курс до ${item.label}`,
      value: item.value,
    })),
    { label: "не подані", value: noCourse, tone: "bad" as const },
  ];
  const ages = getColumnAgeValues(c.age);
  const ageGroups = [
    { label: "до 25", value: ages.filter((age) => age < 25).length },
    {
      label: "25-35",
      value: ages.filter((age) => age >= 25 && age <= 35).length,
    },
    {
      label: "36-45",
      value: ages.filter((age) => age >= 36 && age <= 45).length,
    },
    {
      label: "46-55",
      value: ages.filter((age) => age >= 46 && age <= 55).length,
    },
    { label: "56+", value: ages.filter((age) => age >= 56).length },
  ].filter((item) => item.value > 0);
  const averageAge = ages.length
    ? ages.reduce((sum, age) => sum + age, 0) / ages.length
    : 0;
  const ageRange = ages.length
    ? `${Math.min(...ages)}-${Math.max(...ages)}`
    : "немає даних";
  const ageDiagnostics = getColumnAgeDiagnostics(c.age);
  console.groupCollapsed("[army-grid] Analytics age diagnostics");
  console.log({
    selectedColumnIndex: c.age,
    selectedOriginalExcelColumnIndex:
      c.age >= 0 ? sourceSheet.columnIndexes[c.age] : null,
    selectedExcelColumnNumber:
      c.age >= 0 ? (sourceSheet.columnIndexes[c.age] ?? c.age) + 1 : null,
    selectedHeader: c.age >= 0 ? getColumnHeader(sourceSheet, c.age) : null,
    totalRows: rows.length,
    acceptedAgeRows: ages.length,
    rejectedRows: ageDiagnostics.filter(
      (item) => !item.accepted && item.rawValue,
    ).length,
    averageAge,
    ageRange,
    ageGroups,
  });
  console.table(ageGroups);
  console.table(ageDiagnostics.filter((item) => item.accepted));
  console.table(
    ageDiagnostics.filter((item) => !item.accepted && item.rawValue),
  );
  console.groupEnd();
  const quality = [
    {
      label: "Контроль балансу",
      value: balanceDiff,
      tone: balanceDiff === 0 ? ("good" as const) : ("warn" as const),
      detail: `${listed} - ${absent} - ${inRanks}`,
    },
    {
      label: "Виконання завдань",
      value:
        executionTotal - execution.reduce((sum, item) => sum + item.value, 0),
      tone: "good" as const,
    },
    {
      label: "Не класифіковано у строю",
      value: unclassifiedInRanks.length,
      tone: unclassifiedInRanks.length ? ("warn" as const) : ("good" as const),
      detail: unclassifiedInRanks.length
        ? unclassifiedInRanks
            .map(
              (record) =>
                `${record.excelRowNumber}: ${record.name || "без ПІБ"} · ${record.type || "без типу"} · ${record.location || "без місця"}`,
            )
            .join(" · ")
        : "усі записи в строю входять у формулу",
    },
    {
      label: "Вікові групи",
      value: ages.length,
      tone: ages.length ? ("good" as const) : ("warn" as const),
      detail: ages.length
        ? `Повних років · середній ${averageAge.toFixed(1)} · діапазон ${ageRange} · ${ageGroups.map((item) => `${item.label}: ${item.value}`).join(" · ")}`
        : 'не знайдено колонку "Повних років"',
    },
  ];

  return {
    sourceFileName: snapshot.fileName,
    reportDate: new Intl.DateTimeFormat("uk-UA").format(new Date()),
    sheet,
    rows,
    metrics: {
      staff,
      listed,
      absent,
      inRanks,
      understaff,
      understaffPercent,
      staffedPercent,
      balanceDiff,
      newcomersNoRole,
    },
    absenceReasons,
    detachedDetails,
    attachedDetails,
    structure,
    managementDetails,
    supportDetails,
    execution,
    executionDetails,
    newcomerDetails,
    bzvp,
    ageGroups,
    hintManagementDetails,
    hintSupportDetails,
    hintHealingDetails,
    hintInfantryDetails,
    hintPpdDetails,
    hintRpakDetails,
    unclassifiedInRanks,
    quality,
  };
};
