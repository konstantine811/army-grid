import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type PointerEvent as ReactPointerEvent,
} from "react";
import {
  Box,
  Button,
  Chip,
  IconButton,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  PushPinOutlinedIcon,
  WrapTextOutlinedIcon,
} from "@/components/sci/icons";
import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { useEjoosWorkspace } from "./ejoosWorkspaceState";

export type EjoosSheetKind =
  | "shpo"
  | "oos"
  | "excluded"
  | "tempArrivals"
  | "tempAbsents"
  | "timesheet"
  | "irrevocableLosses";

const SHEET_META: Record<
  EjoosSheetKind,
  { title: string; patterns: RegExp[]; maxRows: number }
> = {
  shpo: {
    title: "1. ШПО",
    patterns: [/1\.?\s*шпо/i, /штатно.?посад/i, /\bшпо\b/i],
    maxRows: 800,
  },
  oos: {
    title: "2. ООС",
    patterns: [/2\.?\s*оос/i, /\bоос\b/i, /обл[іi]к\s*особов/i],
    maxRows: 800,
  },
  excluded: {
    title: "3. Виключені",
    patterns: [/3\.?\s*виключ/i, /виключен/i],
    maxRows: 800,
  },
  tempArrivals: {
    title: "4. Тимчасово прибулі",
    patterns: [/4\.?\s*тимчасово\s*прибу/i, /тимчасово\s*прибул/i],
    maxRows: 500,
  },
  tempAbsents: {
    title: "5. Тимчасово відсутні",
    patterns: [/5\.?\s*тимчасово\s*відсут/i, /тимчасово\s*відсут/i],
    maxRows: 500,
  },
  timesheet: {
    title: "6. Табель",
    patterns: [/6\.?\s*табель/i, /табель/i],
    maxRows: 800,
  },
  irrevocableLosses: {
    title: "7. Безповоротні втрати",
    patterns: [/7\.?\s*безповорот/i, /безповоротн/i],
    maxRows: 500,
  },
};

const ROW_NUMBER_COLUMN_WIDTH = 48;
const COMPACT_COLUMN_WIDTH = 40;
const DEFAULT_COLUMN_WIDTH = 148;
const MIN_COLUMN_WIDTH = 32;
const MAX_COLUMN_WIDTH = 560;

const splitPackedStaffIndexes = (text: string) => {
  const packed = text.replace(/\s+/g, "");
  if (/^(?:\d{7}){2,}$/.test(packed)) {
    return packed.match(/\d{7}/g)?.join("\n") ?? text;
  }
  return text;
};

const cellDisplay = (value: CellValue | undefined): string => {
  if (value == null) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleDateString("uk-UA");
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    // Excel serial dates often appear as numbers in rawRows
    if (value > 20_000 && value < 60_000) {
      const epoch = Date.UTC(1899, 11, 30);
      const date = new Date(epoch + value * 86_400_000);
      if (!Number.isNaN(date.getTime())) {
        return date.toLocaleDateString("uk-UA");
      }
    }
    return splitPackedStaffIndexes(String(value));
  }
  const text = String(value)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");
  if (text.trim() === "[object Object]") return "#N/A";
  return splitPackedStaffIndexes(text)
    .split("\n")
    .map((line) => line.replace(/[^\S\n]+/g, " ").trim())
    .join("\n")
    .trim();
};

const columnWidthsKey = (kind: EjoosSheetKind, sheetName: string) =>
  `ejoos-sheet-column-widths:${kind}:${sheetName}`;

const wrapTextKey = (kind: EjoosSheetKind, sheetName: string) =>
  `ejoos-sheet-wrap-text:${kind}:${sheetName}`;

const pinSecondRowKey = (kind: EjoosSheetKind, sheetName: string) =>
  `ejoos-sheet-pin-second-row:${kind}:${sheetName}`;

const getDefaultColumnWidths = (compactColumns: boolean[]) => [
  ROW_NUMBER_COLUMN_WIDTH,
  ...compactColumns.map((isCompact) =>
    isCompact ? COMPACT_COLUMN_WIDTH : DEFAULT_COLUMN_WIDTH,
  ),
];

const readColumnWidths = (
  kind: EjoosSheetKind,
  sheetName: string,
  compactColumns: boolean[],
) => {
  const defaults = getDefaultColumnWidths(compactColumns);
  if (typeof window === "undefined") return defaults;
  try {
    const parsed = JSON.parse(
      window.localStorage.getItem(columnWidthsKey(kind, sheetName)) || "[]",
    );
    if (!Array.isArray(parsed)) return defaults;
    return defaults.map((defaultWidth, index) => {
      const saved = Number(parsed[index]);
      return Number.isFinite(saved)
        ? Math.min(MAX_COLUMN_WIDTH, Math.max(MIN_COLUMN_WIDTH, saved))
        : defaultWidth;
    });
  } catch {
    return defaults;
  }
};

const writeColumnWidths = (
  kind: EjoosSheetKind,
  sheetName: string,
  widths: number[],
) => {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(
    columnWidthsKey(kind, sheetName),
    JSON.stringify(widths.map((width) => Math.round(width))),
  );
};

const readWrapText = (kind: EjoosSheetKind, sheetName: string) => {
  if (typeof window === "undefined") return true;
  const saved = window.localStorage.getItem(wrapTextKey(kind, sheetName));
  return saved === null ? true : saved === "true";
};

const writeWrapText = (
  kind: EjoosSheetKind,
  sheetName: string,
  wrapped: boolean,
) => {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(wrapTextKey(kind, sheetName), String(wrapped));
};

const readPinSecondRow = (kind: EjoosSheetKind, sheetName: string) => {
  if (typeof window === "undefined") return false;
  return window.localStorage.getItem(pinSecondRowKey(kind, sheetName)) === "true";
};

const writePinSecondRow = (
  kind: EjoosSheetKind,
  sheetName: string,
  pinned: boolean,
) => {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(pinSecondRowKey(kind, sheetName), String(pinned));
};

const findSheet = (
  workbook: ExcelWorkbookSnapshot | null,
  patterns: RegExp[],
): ExcelSheetSnapshot | undefined => {
  if (!workbook) return undefined;
  for (const pattern of patterns) {
    const match = workbook.sheets.find((sheet) =>
      pattern.test(sheet.sheetName),
    );
    if (match) return match;
  }
  return undefined;
};

const findHeaderRowIndex = (sheet: ExcelSheetSnapshot): number => {
  const limit = Math.min(12, sheet.rawRows.length);
  for (let i = 0; i < limit; i += 1) {
    const texts = (sheet.rawRows[i] ?? [])
      .map((cell) => cellDisplay(cell).toLowerCase())
      .filter(Boolean);
    if (texts.length < 2) continue;
    const joined = texts.join(" ");
    if (
      /прізвище|піб|звання|посада|індекс|статус|дата|звідки|табель/.test(
        joined,
      )
    ) {
      return i;
    }
  }
  return 0;
};

const buildTable = (sheet: ExcelSheetSnapshot, maxRows: number) => {
  const headerIndex = findHeaderRowIndex(sheet);
  const headerRow = sheet.rawRows[headerIndex] ?? [];
  let lastCol = 0;
  headerRow.forEach((cell, index) => {
    if (cellDisplay(cell)) lastCol = index + 1;
  });
  // Also expand by first data rows if header is sparse
  for (
    let r = headerIndex + 1;
    r < Math.min(headerIndex + 8, sheet.rawRows.length);
    r += 1
  ) {
    (sheet.rawRows[r] ?? []).forEach((cell, index) => {
      if (cellDisplay(cell)) lastCol = Math.max(lastCol, index + 1);
    });
  }
  lastCol = Math.min(Math.max(lastCol, 4), 40);

  const rawHeaders = Array.from({ length: lastCol }, (_, index) =>
    cellDisplay(headerRow[index]),
  );
  const headers = rawHeaders.map((text, index) => text || `Col ${index + 1}`);

  const dataRows: string[][] = [];
  for (let i = headerIndex + 1; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i] ?? [];
    const values = Array.from({ length: lastCol }, (_, index) =>
      cellDisplay(row[index]),
    );
    if (!values.some((value) => value)) continue;
    dataRows.push(values);
    if (dataRows.length >= maxRows) break;
  }

  return {
    sheetName: sheet.sheetName,
    headerIndex,
    headers,
    compactColumns: rawHeaders.map((header, index) => {
      const headerText = header.trim();
      if (!headerText || /^[-–—]+$/.test(headerText)) return true;
      const filledValues = dataRows
        .map((row) => row[index]?.trim() ?? "")
        .filter(Boolean);
      return (
        filledValues.length > 0 &&
        filledValues.every((value) => /^[-–—]+$/.test(value))
      );
    }),
    rows: dataRows,
    totalRawRows: sheet.rawRows.length,
    truncated: dataRows.length >= maxRows,
  };
};

export function EjoosSheetViewPanel({ kind }: { kind: EjoosSheetKind }) {
  const {
    ejoosSnapshot,
    setTab,
    ensureEjoosLoaded,
    isLoading,
    fillSheetFromAnketa,
  } = useEjoosWorkspace();
  const [query, setQuery] = useState("");
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const [wrapText, setWrapText] = useState(true);
  const [pinSecondRow, setPinSecondRow] = useState(false);
  const [headerHeight, setHeaderHeight] = useState(40);
  const tableHeadRef = useRef<HTMLTableSectionElement | null>(null);
  const meta = SHEET_META[kind];

  const sheet = useMemo(
    () => findSheet(ejoosSnapshot, meta.patterns),
    [ejoosSnapshot, meta.patterns],
  );

  const table = useMemo(
    () => (sheet ? buildTable(sheet, meta.maxRows) : null),
    [sheet, meta.maxRows],
  );

  const filteredRows = useMemo(() => {
    if (!table) return [];
    const needle = query.trim().toLocaleLowerCase("uk-UA");
    if (!needle) return table.rows;
    return table.rows.filter((row) =>
      row.join(" ").toLocaleLowerCase("uk-UA").includes(needle),
    );
  }, [query, table]);

  const columnLayoutKey = table
    ? `${kind}:${table.sheetName}:${table.headers.length}:${table.compactColumns.join(",")}`
    : "";

  useEffect(() => {
    if (!table) {
      setColumnWidths([]);
      return;
    }
    setColumnWidths(
      readColumnWidths(kind, table.sheetName, table.compactColumns),
    );
    setWrapText(readWrapText(kind, table.sheetName));
    setPinSecondRow(readPinSecondRow(kind, table.sheetName));
  }, [columnLayoutKey]);

  useEffect(() => {
    if (!tableHeadRef.current || typeof ResizeObserver === "undefined") return;
    const updateHeight = () => {
      const next = Math.ceil(
        tableHeadRef.current?.getBoundingClientRect().height || 40,
      );
      setHeaderHeight(next);
    };
    updateHeight();
    const observer = new ResizeObserver(updateHeight);
    observer.observe(tableHeadRef.current);
    return () => observer.disconnect();
  }, [columnLayoutKey, columnWidths, wrapText]);

  const resizeColumn = useCallback(
    (columnIndex: number, nextWidth: number) => {
      if (!table) return;
      setColumnWidths((current) => {
        const base =
          current.length === table.headers.length + 1
            ? current
            : getDefaultColumnWidths(table.compactColumns);
        const next = [...base];
        next[columnIndex] = Math.min(
          MAX_COLUMN_WIDTH,
          Math.max(MIN_COLUMN_WIDTH, nextWidth),
        );
        writeColumnWidths(kind, table.sheetName, next);
        return next;
      });
    },
    [kind, table],
  );

  const resetColumnWidths = useCallback(() => {
    if (!table) return;
    const next = getDefaultColumnWidths(table.compactColumns);
    setColumnWidths(next);
    if (typeof window !== "undefined") {
      window.localStorage.removeItem(columnWidthsKey(kind, table.sheetName));
    }
  }, [kind, table]);

  const toggleWrapText = useCallback(() => {
    if (!table) return;
    setWrapText((current) => {
      const next = !current;
      writeWrapText(kind, table.sheetName, next);
      return next;
    });
  }, [kind, table]);

  const togglePinSecondRow = useCallback(() => {
    if (!table) return;
    setPinSecondRow((current) => {
      const next = !current;
      writePinSecondRow(kind, table.sheetName, next);
      return next;
    });
  }, [kind, table]);

  const startResize = useCallback(
    (columnIndex: number, event: ReactPointerEvent<HTMLSpanElement>) => {
      if (!table) return;
      event.preventDefault();
      const startX = event.clientX;
      const startWidth =
        columnWidths[columnIndex] ??
        getDefaultColumnWidths(table.compactColumns)[columnIndex] ??
        DEFAULT_COLUMN_WIDTH;

      const handlePointerMove = (moveEvent: PointerEvent) => {
        resizeColumn(columnIndex, startWidth + moveEvent.clientX - startX);
      };
      const handlePointerUp = () => {
        window.removeEventListener("pointermove", handlePointerMove);
        window.removeEventListener("pointerup", handlePointerUp);
      };

      window.addEventListener("pointermove", handlePointerMove);
      window.addEventListener("pointerup", handlePointerUp, { once: true });
    },
    [columnWidths, resizeColumn, table],
  );

  if (!ejoosSnapshot) {
    return (
      <Box className="ejoos-sheet-view">
        <Typography variant="h6">{meta.title}</Typography>
        <Typography variant="body2" className="ejoos-muted" sx={{ mt: 1 }}>
          Спочатку завантажте файл ЕЖООС на вкладці «Імпорт».
        </Typography>
        <button
          type="button"
          className="ejoos-workspace-nav-btn"
          style={{ marginTop: 12 }}
          onClick={() => {
            void ensureEjoosLoaded();
            setTab("import");
          }}
          disabled={isLoading}
        >
          До імпорту
        </button>
      </Box>
    );
  }

  if (!table) {
    return (
      <Box className="ejoos-sheet-view">
        <Typography variant="h6">{meta.title}</Typography>
        <Typography variant="body2" className="ejoos-muted" sx={{ mt: 1 }}>
          У файлі «{ejoosSnapshot.fileName}» не знайдено аркуш для{" "}
          {meta.title}. Є:{" "}
          {ejoosSnapshot.sheets.map((item) => item.sheetName).join(", ")}.
        </Typography>
      </Box>
    );
  }

  return (
    <Box className="ejoos-sheet-view">
      <Stack
        direction="row"
        spacing={1}
        alignItems="center"
        sx={{ mb: 1.5, flexWrap: "wrap", gap: 1 }}
      >
        <Typography variant="h6" sx={{ mr: 1 }}>
          {meta.title}
        </Typography>
        <Chip size="small" label={table.sheetName} />
        <Chip
          size="small"
          variant="outlined"
          label={`рядків: ${filteredRows.length}${
            table.truncated ? ` (ліміт ${meta.maxRows})` : ""
          }`}
        />
        <Chip
          size="small"
          variant="outlined"
          label={ejoosSnapshot.fileName}
        />
        <Button size="small" variant="outlined" onClick={resetColumnWidths}>
          Скинути ширини
        </Button>
        {kind === "oos" || kind === "excluded" ? (
          <Button
            size="small"
            variant="contained"
            disabled={isLoading}
            onClick={() =>
              void fillSheetFromAnketa(kind === "excluded" ? "excluded" : "oos")
            }
            sx={{ color: "#1a1a14" }}
            title="Знайти людей в анкетних даних і дописати лише порожні поля"
          >
            {isLoading ? "Доповнюю…" : "Доповнити з анкетних даних"}
          </Button>
        ) : null}
        {kind === "oos" || kind === "excluded" ? (
          <Button
            size="small"
            variant="contained"
            disabled={isLoading}
            onClick={() =>
              void fillSheetFromAnketa(
                kind === "excluded" ? "excluded" : "oos",
                "merge",
              )
            }
            sx={{ color: "#1a1a14" }}
            title="Колонки пропусків з анкет → ця таблиця. Збіг по ID, інакше по унікальному ПІБ. Пише в порожнє або якщо в анкеті повніше."
          >
            {isLoading ? "Мердж…" : "Мердж"}
          </Button>
        ) : null}
        <IconButton
          size="small"
          className={wrapText ? "is-active" : undefined}
          title={wrapText ? "Перенос тексту увімкнено" : "Перенос тексту вимкнено"}
          aria-label={
            wrapText ? "Вимкнути перенос тексту" : "Увімкнути перенос тексту"
          }
          aria-pressed={wrapText}
          onClick={toggleWrapText}
        >
          <WrapTextOutlinedIcon fontSize="small" />
        </IconButton>
        <IconButton
          size="small"
          className={pinSecondRow ? "is-active" : undefined}
          title={
            pinSecondRow
              ? "Другий рядок закріплено"
              : "Другий рядок не закріплено"
          }
          aria-label={
            pinSecondRow
              ? "Відкріпити другий рядок"
              : "Закріпити другий рядок"
          }
          aria-pressed={pinSecondRow}
          onClick={togglePinSecondRow}
        >
          <PushPinOutlinedIcon fontSize="small" />
        </IconButton>
      </Stack>

      <input
        className="ejoos-sheet-search"
        type="search"
        placeholder="Пошук по таблиці…"
        value={query}
        onChange={(event) => setQuery(event.target.value)}
        style={{
          width: "100%",
          maxWidth: 420,
          marginBottom: 12,
          padding: "8px 12px",
          borderRadius: 6,
          border: "1px solid rgba(255,255,255,0.2)",
          background: "rgba(0,0,0,0.25)",
          color: "inherit",
        }}
      />

      <div className="bchs-analytics-table-wrap ejoos-sheet-table-wrap">
        <table
          className={
            wrapText
              ? `bchs-analytics-table ejoos-sheet-table is-wrap-text${pinSecondRow ? " is-pin-second-row" : ""}`
              : `bchs-analytics-table ejoos-sheet-table is-nowrap-text${pinSecondRow ? " is-pin-second-row" : ""}`
          }
          style={{
            "--ejoos-header-height": `${headerHeight}px`,
            minWidth: "100%",
            width: `${(columnWidths.length ? columnWidths : getDefaultColumnWidths(table.compactColumns)).reduce(
              (sum, width) => sum + width,
              0,
            )}px`,
          } as CSSProperties & Record<string, string>}
        >
          <colgroup>
            {(columnWidths.length
              ? columnWidths
              : getDefaultColumnWidths(table.compactColumns)
            ).map((width, index) => (
              <col key={`col-${index}`} style={{ width: `${width}px` }} />
            ))}
          </colgroup>
          <thead ref={tableHeadRef}>
            <tr>
              <th>
                <span className="ejoos-sheet-th-content">#</span>
                <span
                  className="ejoos-column-resize-handle"
                  role="separator"
                  aria-orientation="vertical"
                  aria-label="Змінити ширину колонки #"
                  onPointerDown={(event) => startResize(0, event)}
                />
              </th>
              {table.headers.map((header, index) => (
                <th
                  key={`${header}-${index}`}
                  className={table.compactColumns[index] ? "is-compact-col" : undefined}
                >
                  <span className="ejoos-sheet-th-content">{header}</span>
                  <span
                    className="ejoos-column-resize-handle"
                    role="separator"
                    aria-orientation="vertical"
                    aria-label={`Змінити ширину колонки ${header}`}
                    onPointerDown={(event) => startResize(index + 1, event)}
                  />
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredRows.map((row, rowIndex) => (
              <tr
                key={`r-${rowIndex}`}
                className={rowIndex === 0 ? "ejoos-second-row" : undefined}
              >
                <td>{rowIndex + 1}</td>
                {row.map((cell, cellIndex) => (
                  <td
                    key={`c-${rowIndex}-${cellIndex}`}
                    className={
                      table.compactColumns[cellIndex] ? "is-compact-col" : undefined
                    }
                  >
                    {cell}
                  </td>
                ))}
              </tr>
            ))}
            {filteredRows.length === 0 ? (
              <tr>
                <td colSpan={table.headers.length + 1}>Немає рядків</td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>
    </Box>
  );
}
