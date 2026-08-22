import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "./rosterSourceSnapshot";

/** Google Sheet «Штатка». */
export const STAFF_SHEET_ID = "1dqorfOj7TTZt6sBBV1mc_dupxRkUSyeRjxJjlV07A38";
export const STAFF_SHEET_GID = "0";

export const staffSheetEditUrl = () =>
  `https://docs.google.com/spreadsheets/d/${STAFF_SHEET_ID}/edit?usp=sharing`;

const APPS_SCRIPT_STORAGE_KEY = "army-grid:staff-sheet-apps-script-url";

/** Roster column numbers → sequential columns A… in «Штатка» (row 1 headers). */
export const STAFF_SHEET_ROSTER_COLUMNS = [
  1, 2, 3, 4, 5, 8, 10, 11, 12, 13, 14, 15, 21, 22, 23, 24, 25, 26, 27, 28, 29,
  31, 32, 33, 34,
] as const;

export const STAFF_SHEET_HEADERS = [
  "№",
  "Підрозділ",
  "Взвод",
  "Відділення",
  "Посада",
  "ШПК факт",
  "Анкета",
  "Військовий квиток",
  "Мобілізація/контракт",
  "Звання",
  "ПІБ",
  "Позивний",
  "СТАТУС",
  "Тип В\\С",
  "Статус БГ",
  "БЗВП/БРЕЗ",
  "Наявність БЗВП",
  "Курс БЗВП",
  "Відрядження (БРЕЗ)",
  "Обмеження",
  "В якому підрозділі",
  "Місце перебування",
  "Примітки",
  "Напрямок",
  "Примітка 3",
] as const;

const hasPersonName = (row: EjournalPreviewRow) => {
  const name = readRosterColumnValue(row, 14);
  return name.trim().length >= 3;
};

export const rosterRowsToStaffSheetValues = (
  rows: EjournalPreviewRow[],
): string[][] =>
  rows
    .filter(hasPersonName)
    .map((row, index) =>
      STAFF_SHEET_ROSTER_COLUMNS.map((columnNumber, columnIndex) => {
        if (columnIndex === 0) return String(index + 1);
        return readRosterColumnValue(row, columnNumber);
      }),
    );

export const readStaffAppsScriptUrl = () => {
  try {
    return window.localStorage.getItem(APPS_SCRIPT_STORAGE_KEY)?.trim() || "";
  } catch {
    return "";
  }
};

export const writeStaffAppsScriptUrl = (url: string) => {
  try {
    const next = url.trim();
    if (next) window.localStorage.setItem(APPS_SCRIPT_STORAGE_KEY, next);
    else window.localStorage.removeItem(APPS_SCRIPT_STORAGE_KEY);
  } catch {
    /* ignore */
  }
};

export const pushStaffSheetToGoogle = async (
  values: string[][],
  options?: { sheetId?: string; gid?: string },
) => {
  const endpoint = readStaffAppsScriptUrl();
  if (!endpoint) {
    throw new Error(
      "Немає URL Apps Script для «Штатка». Вставте посилання вебзастосунку нижче.",
    );
  }
  if (!values.length) {
    throw new Error("Немає рядків для оновлення Google Sheet.");
  }

  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({
      sheetId: options?.sheetId ?? STAFF_SHEET_ID,
      gid: options?.gid ?? STAFF_SHEET_GID,
      mode: "bulk",
      startRow: 2,
      startCol: 1,
      values,
      clearAfterRow: true,
    }),
  });

  if (!response.ok) {
    throw new Error(`Google sync HTTP ${response.status}`);
  }

  const text = await response.text();
  try {
    const json = JSON.parse(text) as { ok?: boolean; error?: string; rows?: number };
    if (json.error) throw new Error(json.error);
    if (json.ok === false) throw new Error("Apps Script повернув помилку.");
    return json.rows ?? values.length;
  } catch (error) {
    if (error instanceof SyntaxError) return values.length;
    throw error;
  }
};

export const STAFF_APPS_SCRIPT_TEMPLATE = `/**
 * Google Apps Script — таблиця «Штатка».
 * Deploy → Web app · Execute as: Me · Who has access: Anyone
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var ss = SpreadsheetApp.openById(body.sheetId);
    var sheet = ss.getSheets()[0];
    if (body.gid != null) {
      var byGid = ss.getSheets().filter(function(s) {
        return String(s.getSheetId()) === String(body.gid);
      })[0];
      if (byGid) sheet = byGid;
    }

    if (body.mode === 'bulk' && body.values && body.values.length) {
      var startRow = Number(body.startRow) || 2;
      var startCol = Number(body.startCol) || 1;
      var rows = body.values.length;
      var cols = body.values[0].length;
      sheet.getRange(startRow, startCol, rows, cols).setValues(body.values);

      if (body.clearAfterRow) {
        var last = sheet.getLastRow();
        var clearFrom = startRow + rows;
        if (last >= clearFrom) {
          sheet.getRange(clearFrom, startCol, last - clearFrom + 1, cols).clearContent();
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, rows: rows }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.a1) {
      sheet.getRange(String(body.a1)).setValue(body.value == null ? '' : body.value);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, a1: body.a1 }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    throw new Error('Unknown payload');
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
`;
