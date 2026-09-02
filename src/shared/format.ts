export const formatFileSize = (size: number) => `${(size / 1024).toFixed(1)} KB`;

export const KYIV_TIME_ZONE = "Europe/Kyiv";

const parseApiDate = (value: string | Date | null | undefined) => {
  if (value == null || value === "") return null;
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
};

/** API instants (ISO) as Kyiv wall clock, independent of the browser time zone. */
export const formatApiDateTime = (value: string | Date | null | undefined) => {
  const date = parseApiDate(value);
  if (!date) return value == null ? "" : String(value);
  return date.toLocaleString("uk-UA", { timeZone: KYIV_TIME_ZONE });
};

export const formatApiDate = (value: string | Date | null | undefined) => {
  const date = parseApiDate(value);
  if (!date) return "";
  return date.toLocaleDateString("uk-UA", {
    timeZone: KYIV_TIME_ZONE,
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
};

export const formatDateTime = () =>
  new Intl.DateTimeFormat("uk-UA", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    timeZone: KYIV_TIME_ZONE,
  }).format(new Date());

const EXCEL_SERIAL_MIN = 1000;
const EXCEL_SERIAL_MAX = 100000;

/** Excel stores dates as day count from 1899-12-30; 46187 ≈ 14.06.2026. */
export const tryParseExcelSerialDate = (value: unknown): Date | null => {
  const numeric =
    typeof value === "number"
      ? value
      : typeof value === "string" && /^\d{4,5}(\.0+)?$/.test(value.trim())
        ? Number(value.trim())
        : null;
  if (numeric === null || !Number.isFinite(numeric)) return null;

  const serial = Math.floor(numeric);
  if (serial < EXCEL_SERIAL_MIN || serial > EXCEL_SERIAL_MAX) return null;

  const date = new Date((serial - 25569) * 86400000);
  const year = date.getUTCFullYear();
  if (year < 1950 || year > 2100) return null;
  return date;
};

export const formatUkDate = (value: Date) =>
  new Intl.DateTimeFormat("uk-UA", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(value);

export const formatValueForDisplay = (value: unknown): string => {
  if (value === undefined || value === null) return "";
  if (value instanceof Date) return formatUkDate(value);

  const parsedSerial = tryParseExcelSerialDate(value);
  if (parsedSerial) return formatUkDate(parsedSerial);

  if (typeof value === "string") {
    if (/^\d{4}-\d{2}-\d{2}T/.test(value)) {
      const parsedIso = new Date(value);
      if (!Number.isNaN(parsedIso.getTime())) return formatUkDate(parsedIso);
    }
  }

  return String(value);
};

export const normalizeDatasetKey = (value: string) =>
  value
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^\p{L}\p{N}_-]/gu, "")
    .toLowerCase();

export const cellValueToJson = (value: unknown) => {
  if (value === undefined || value === null) return null;
  if (value instanceof Date) return value.toISOString();
  const parsedSerial = tryParseExcelSerialDate(value);
  if (parsedSerial) return parsedSerial.toISOString();
  return value;
};
