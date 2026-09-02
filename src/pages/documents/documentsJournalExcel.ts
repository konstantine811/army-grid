import writeXlsxFile, {
  type Cell,
  type CellObject,
  type SheetData,
} from "write-excel-file/browser";
import dayjs from "dayjs";

export type DocumentsJournalExcelRow = {
  personName: string;
  personId: string;
  personStatus: string;
  documentType: string;
  progressPercent: number;
  status: string;
  note: string;
  files: string;
  formPurpose?: string;
  taskPeriodEnd?: string;
  createdAt: string;
  updatedAt: string;
};

const FULL_BASE_COLUMNS = 11;
const COMPACT_FORM_COLUMNS = 6;

const journalColumnCount = ({
  includeUbdExitDate,
  includeFormPurpose,
  compactFormExport,
}: {
  includeUbdExitDate: boolean;
  includeFormPurpose: boolean;
  compactFormExport: boolean;
}) =>
  compactFormExport
    ? COMPACT_FORM_COLUMNS
    : FULL_BASE_COLUMNS +
      (includeUbdExitDate ? 1 : 0) +
      (includeFormPurpose ? 1 : 0);
const HEADER_GREEN = "#1F4E3D";
const TITLE_GREEN = "#14382C";
const META_BG = "#E8F0EC";
const BORDER = "#A9BDB4";
const ZEBRA = "#F4F8F6";
const WHITE = "#FFFFFF";
const TEXT = "#1A1A14";
const MUTED = "#5C6B64";

const border = {
  borderColor: BORDER,
  borderStyle: "thin" as const,
};

const cell = (
  value: string | number | Date | null,
  extra: Omit<CellObject, "value"> = {},
): CellObject => ({
  value: value ?? undefined,
  fontFamily: "Calibri",
  fontSize: 11,
  alignVertical: "center",
  ...border,
  ...extra,
});

const emptyCell = (extra: Omit<CellObject, "value"> = {}): CellObject =>
  cell(null, extra);

const progressFill = (percent: number) => {
  if (percent >= 75) return { backgroundColor: "#CEEAD6", textColor: "#0D652D" };
  if (percent >= 50) return { backgroundColor: "#E6F4EA", textColor: "#137333" };
  if (percent >= 25) return { backgroundColor: "#FEF7E0", textColor: "#8A5A00" };
  return { backgroundColor: "#FCE8E6", textColor: "#A50E0E" };
};

const countBy = (rows: DocumentsJournalExcelRow[], key: keyof DocumentsJournalExcelRow) => {
  const map = new Map<string, { count: number; progressSum: number }>();
  for (const row of rows) {
    const label = String(row[key] || "—");
    const current = map.get(label) ?? { count: 0, progressSum: 0 };
    current.count += 1;
    current.progressSum += row.progressPercent;
    map.set(label, current);
  }
  return [...map.entries()].sort((left, right) => right[1].count - left[1].count);
};

const titleRow = (title: string, columnCount: number): Cell[] => [
  cell(title, {
    fontSize: 16,
    fontWeight: "bold",
    textColor: WHITE,
    backgroundColor: TITLE_GREEN,
    align: "left",
    height: 28,
    columnSpan: columnCount,
  }),
  ...Array.from({ length: columnCount - 1 }, () => null),
];

const metaRow = (text: string, columnCount: number): Cell[] => [
  cell(text, {
    fontSize: 10,
    textColor: MUTED,
    backgroundColor: META_BG,
    align: "left",
    height: 20,
    columnSpan: columnCount,
  }),
  ...Array.from({ length: columnCount - 1 }, () => null),
];

const headerRow = (labels: string[]): CellObject[] =>
  labels.map((label) =>
    cell(label, {
      fontWeight: "bold",
      fontSize: 10,
      textColor: WHITE,
      backgroundColor: HEADER_GREEN,
      align: "center",
      wrap: true,
      height: 24,
    }),
  );

const buildJournalSheet = (
  rows: DocumentsJournalExcelRow[],
  meta: string,
  includeUbdExitDate: boolean,
  includeFormPurpose: boolean,
  compactFormExport: boolean,
): SheetData => {
  const columns = journalColumnCount({
    includeUbdExitDate,
    includeFormPurpose,
    compactFormExport,
  });
  const headers = compactFormExport
    ? [
        "№",
        "Службовець",
        "ID",
        "Статус службовця",
        "Документ",
        "Для чого форма",
      ]
    : [
        "№",
        "Службовець",
        "ID",
        "Статус службовця",
        "Документ",
        "Прогрес",
        "Статус документа",
        "Коментар",
        "Файли",
        ...(includeFormPurpose ? ["Для чого форма"] : []),
        ...(includeUbdExitDate ? ["Вихід від"] : []),
        "Створено",
        "Оновлено",
      ];

  return [
    titleRow("Журнал документів", columns),
    metaRow(meta, columns),
    Array.from({ length: columns }, () => emptyCell({ height: 8 })),
    headerRow(headers),
    ...rows.map((row, index) => {
      const zebra = index % 2 === 1 ? { backgroundColor: ZEBRA } : { backgroundColor: WHITE };
      const created = dayjs(row.createdAt);
      const updated = dayjs(row.updatedAt);
      const lead = [
        cell(index + 1, { type: Number, align: "center", textColor: MUTED, ...zebra }),
        cell(row.personName, { fontWeight: "bold", textColor: TEXT, wrap: true, ...zebra }),
        cell(row.personId || "—", { align: "left", textColor: MUTED, wrap: true, ...zebra }),
        cell(row.personStatus || "—", { wrap: true, ...zebra }),
        cell(row.documentType, { wrap: true, ...zebra }),
      ];
      if (compactFormExport) {
        return [
          ...lead,
          cell(row.formPurpose || "—", { wrap: true, ...zebra }),
        ];
      }
      return [
        ...lead,
        cell(row.progressPercent / 100, {
          type: Number,
          format: "0%",
          align: "center",
          fontWeight: "bold",
          ...progressFill(row.progressPercent),
          ...border,
        }),
        cell(row.status, { wrap: true, ...zebra }),
        cell(row.note || "—", { wrap: true, ...zebra }),
        cell(row.files, { align: "center", ...zebra }),
        ...(includeFormPurpose
          ? [cell(row.formPurpose || "—", { wrap: true, ...zebra })]
          : []),
        ...(includeUbdExitDate
          ? [cell(row.taskPeriodEnd || "—", { align: "center", ...zebra })]
          : []),
        cell(created.isValid() ? created.toDate() : row.createdAt, {
          type: created.isValid() ? Date : String,
          format: "DD.MM.YYYY HH:mm",
          align: "center",
          ...zebra,
        }),
        cell(updated.isValid() ? updated.toDate() : row.updatedAt, {
          type: updated.isValid() ? Date : String,
          format: "DD.MM.YYYY HH:mm",
          align: "center",
          ...zebra,
        }),
      ];
    }),
  ];
};

const buildSummaryBlock = (
  title: string,
  items: Array<[string, { count: number; progressSum: number }]>,
  total: number,
): SheetData => [
  headerRow(["Показник", "Кількість", "Частка", "Середній прогрес"]).map((item, index) =>
    index === 0 ? { ...item, value: title } : item,
  ),
  ...items.map(([label, stats], index) => {
    const zebra = index % 2 === 1 ? { backgroundColor: ZEBRA } : { backgroundColor: WHITE };
    const average = stats.count ? Math.round(stats.progressSum / stats.count) : 0;
    return [
      cell(label, { wrap: true, ...zebra }),
      cell(stats.count, { type: Number, align: "center", ...zebra }),
      cell(total ? stats.count / total : 0, {
        type: Number,
        format: "0%",
        align: "center",
        ...zebra,
      }),
      cell(average / 100, {
        type: Number,
        format: "0%",
        align: "center",
        fontWeight: "bold",
        ...progressFill(average),
        ...border,
      }),
    ];
  }),
  [],
];

const buildSummarySheet = (
  rows: DocumentsJournalExcelRow[],
  meta: string,
): SheetData => [
  titleRow("Зведення журналу документів", 4),
  metaRow(meta, 4),
  Array.from({ length: 4 }, () => emptyCell({ height: 8 })),
  ...buildSummaryBlock("Тип документа", countBy(rows, "documentType"), rows.length),
  ...buildSummaryBlock("Статус документа", countBy(rows, "status"), rows.length),
  ...buildSummaryBlock("Статус службовця", countBy(rows, "personStatus"), rows.length),
];

export const exportDocumentsJournalExcel = async ({
  rows,
  typeFilterLabel,
  statusFilterLabel,
  periodFilterLabel,
  totalCount,
  includeUbdExitDate = false,
  includeFormPurpose = false,
  compactFormExport = false,
}: {
  rows: DocumentsJournalExcelRow[];
  typeFilterLabel: string;
  statusFilterLabel: string;
  periodFilterLabel: string;
  totalCount: number;
  includeUbdExitDate?: boolean;
  includeFormPurpose?: boolean;
  compactFormExport?: boolean;
}) => {
  const exportedAt = dayjs();
  const meta = [
    `Станом на ${exportedAt.format("DD.MM.YYYY HH:mm")}`,
    `Фільтр: ${typeFilterLabel} · ${statusFilterLabel} · ${periodFilterLabel}`,
    `Рядків: ${rows.length}${totalCount !== rows.length ? ` з ${totalCount}` : ""}`,
  ].join("  ·  ");
  const periodSuffix =
    periodFilterLabel === "Усі місяці" ? "" : ` ${periodFilterLabel}`;
  const fileName = `Журнал документів${periodSuffix} ${exportedAt.format("DD.MM.YYYY")}.xlsx`;
  const journalColumns = compactFormExport
    ? [
        { width: 6 },
        { width: 34 },
        { width: 22 },
        { width: 22 },
        { width: 32 },
        { width: 36 },
      ]
    : [
        { width: 6 },
        { width: 34 },
        { width: 22 },
        { width: 22 },
        { width: 32 },
        { width: 12 },
        { width: 24 },
        { width: 36 },
        { width: 14 },
        ...(includeFormPurpose ? [{ width: 28 }] : []),
        ...(includeUbdExitDate ? [{ width: 14 }] : []),
        { width: 20 },
        { width: 20 },
      ];

  await writeXlsxFile(
    [
      {
        sheet: "Журнал",
        data: buildJournalSheet(
          rows,
          meta,
          includeUbdExitDate,
          includeFormPurpose,
          compactFormExport,
        ),
        orientation: "landscape",
        stickyRowsCount: 4,
        stickyColumnsCount: 2,
        dateFormat: "DD.MM.YYYY HH:mm",
        showGridLines: false,
        columns: journalColumns,
      },
      {
        sheet: "Зведення",
        data: buildSummarySheet(rows, meta),
        stickyRowsCount: 3,
        showGridLines: false,
        columns: [{ width: 36 }, { width: 14 }, { width: 12 }, { width: 18 }],
      },
    ],
    { fontFamily: "Calibri", fontSize: 11 },
  ).toFile(fileName);

  return { fileName, rowCount: rows.length };
};
