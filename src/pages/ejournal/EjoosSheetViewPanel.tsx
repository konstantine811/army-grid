import { useMemo, useState } from "react";
import { Box, Chip, Stack, Typography } from "@/components/sci/SciPrimitives";
import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { useEjoosWorkspace } from "./EjoosWorkspaceContext";

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
    return String(value);
  }
  return String(value).replace(/\s+/g, " ").trim();
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

  const headers = Array.from({ length: lastCol }, (_, index) => {
    const text = cellDisplay(headerRow[index]);
    return text || `Col ${index + 1}`;
  });

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
    rows: dataRows,
    totalRawRows: sheet.rawRows.length,
    truncated: dataRows.length >= maxRows,
  };
};

export function EjoosSheetViewPanel({ kind }: { kind: EjoosSheetKind }) {
  const { ejoosSnapshot, setTab, ensureEjoosLoaded, isLoading } =
    useEjoosWorkspace();
  const [query, setQuery] = useState("");
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
        <table className="bchs-analytics-table ejoos-sheet-table">
          <thead>
            <tr>
              <th>#</th>
              {table.headers.map((header, index) => (
                <th key={`${header}-${index}`}>{header}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredRows.map((row, rowIndex) => (
              <tr key={`r-${rowIndex}`}>
                <td>{rowIndex + 1}</td>
                {row.map((cell, cellIndex) => (
                  <td key={`c-${rowIndex}-${cellIndex}`}>{cell}</td>
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
