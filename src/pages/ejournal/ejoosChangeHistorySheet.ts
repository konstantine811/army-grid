import JSZip from "jszip";
import type { BackendEjournalLiveVersion } from "../../api";
import { formatApiDate } from "../../shared/format";
import { collectProcessedMovementKeys } from "./ejoosSyncPlan";
import {
  applyInlineStringWritesToWorkbook,
  resolveSheetPath,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";

export const EJOOS_CHANGE_HISTORY_SHEET = "10. Історія змін";

const KIND_TYPE_LABEL: Record<string, string> = {
  apply: "Застосування",
  review: "ПОТРЕБУЄ ПЕРЕВІРКИ",
  import: "Імпорт",
  rollback: "Відкат",
};

const OP_KIND_LABEL: Record<string, string> = {
  exclude_transfer: "виключені",
  move_to_disposition: "розпорядження",
  position_change: "посада",
  rank_change: "звання",
  arrival: "прибуття",
  timesheet_day: "табель",
  shpo_occupant: "ШПО",
  absent_close: "закрито відсутність",
  absent_upsert: "тимч. відсутні",
  data_mismatch: "дані",
  other_manual: "інше",
};

/** Що пишемо в історію: фактично застосовані ops, не евристика по файлу. */
const HISTORY_OP_KINDS = new Set([
  "exclude_transfer",
  "move_to_disposition",
  "position_change",
  "rank_change",
  "arrival",
  "shpo_occupant",
  "absent_close",
  "absent_upsert",
  "other_manual",
]);

export type EjoosHistorySheetRow = {
  versionLabel: string;
  date: string;
  type: string;
  description: string;
  comment: string;
  createdAtIso: string;
};

const parseAttr = (attrs: string, name: string) =>
  attrs.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];

const decodeXml = (value: string) =>
  value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&");

const parseSharedStrings = (xml: string) => {
  const values: string[] = [];
  for (const item of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/gi)) {
    const parts = [...item[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)].map(
      (match) => decodeXml(match[1]),
    );
    values.push(parts.join(""));
  }
  return values;
};

const extractCellText = (
  cellXml: string,
  attrs: string,
  sharedStrings: string[],
) => {
  const type = (parseAttr(attrs, "t") || "").toLowerCase();
  if (type === "inlinestr") {
    return [...cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)]
      .map((match) => decodeXml(match[1]))
      .join("");
  }
  const raw = cellXml.match(/<v\b[^>]*>([\s\S]*?)<\/v>/i)?.[1] ?? "";
  if (type === "s") {
    const index = Number(raw);
    return Number.isFinite(index) ? (sharedStrings[index] ?? "") : "";
  }
  return decodeXml(raw);
};

const cellHasValue = (cellXml: string) => /<(?:v|is|f)\b/i.test(cellXml);

const readHistorySheetState = async (file: Blob | File) => {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const sheetPath = await resolveSheetPath(zip, EJOOS_CHANGE_HISTORY_SHEET);
  if (!sheetPath) return null;
  const sheetXml = await zip.file(sheetPath)?.async("string");
  if (!sheetXml) return null;
  const sharedXml = await zip.file("xl/sharedStrings.xml")?.async("string");
  const sharedStrings = sharedXml ? parseSharedStrings(sharedXml) : [];

  const columnA = new Map<number, string>();
  let lastContentRow = 1;
  for (const cell of sheetXml.matchAll(/<c\b([^<>]*?)(\/>|>[\s\S]*?<\/c>)/gi)) {
    const attrs = cell[1];
    const ref = parseAttr(attrs, "r") || "";
    const parts = ref.match(/^([A-E])(\d+)$/i);
    if (!parts || !cellHasValue(cell[0])) continue;
    const row = Number(parts[2]);
    lastContentRow = Math.max(lastContentRow, row);
    if (parts[1].toUpperCase() === "A") {
      columnA.set(row, extractCellText(cell[0], attrs, sharedStrings).trim());
    }
  }

  return { lastContentRow, columnA };
};

const formatUaDate = (iso: string) => formatApiDate(iso);

const uaPlural = (count: number, one: string, few: string, many: string) => {
  const abs = Math.abs(count);
  const mod10 = abs % 10;
  const mod100 = abs % 100;
  if (mod10 === 1 && mod100 !== 11) return `${count} ${one}`;
  if (mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14)) {
    return `${count} ${few}`;
  }
  return `${count} ${many}`;
};

export const formatUaChangeCount = (count: number) =>
  uaPlural(count, "зміна", "зміни", "змін");

const asRecord = (value: unknown): Record<string, unknown> | null =>
  value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;

const asString = (value: unknown) =>
  typeof value === "string" ? value.trim() : "";

const versionLabelOf = (version: number) => `v${version}`;

const isOurVersionLabel = (text: string) => /^v\d+$/i.test(text.trim());

const versionNumberOfLabel = (text: string) => {
  const match = text.trim().match(/^v(\d+)$/i);
  if (!match) return 0;
  const value = Number(match[1]);
  return Number.isFinite(value) ? value : 0;
};

const maxVersionLabelInSheet = (columnA: Map<number, string>) =>
  Math.max(
    0,
    ...[...columnA.values()]
      .map(versionNumberOfLabel)
      .filter((value) => value > 0),
  );

const isOfficialTemplateVersion = (text: string) =>
  /^\d+\.\d+(\.\d+)*$/.test(text.trim());

const isMisplacedIsoTimestamp = (text: string) =>
  /^\d{4}-\d{2}-\d{2}T/.test(text.trim());

const typeLabelForKind = (kind: string) => {
  if (KIND_TYPE_LABEL[kind]) return KIND_TYPE_LABEL[kind];
  if (kind.startsWith("anketa")) return "Анкетні дані";
  return "Зміни";
};

const titleUaWord = (word: string) => {
  if (!word) return word;
  if (/^[A-Za-zА-Яа-яЁёІіЇїЄєҐґ]\.$/u.test(word)) {
    return word.toLocaleUpperCase("uk-UA");
  }
  return (
    word.charAt(0).toLocaleUpperCase("uk-UA") +
    word.slice(1).toLocaleLowerCase("uk-UA")
  );
};

/** Як у ШПО: ПРІЗВИЩЕ Ім’я По батькові, без ініціалів. */
const formatPersonFull = (fullName: string, personId: string) => {
  const parts = fullName.replace(/\s+/g, " ").trim().split(" ").filter(Boolean);
  if (!parts.length) return personId ? `ID ${personId}` : "—";
  return [parts[0].toLocaleUpperCase("uk-UA"), ...parts.slice(1).map(titleUaWord)].join(
    " ",
  );
};

const shortCell = (value: string, max = 48) => {
  const text = value.replace(/\s+/g, " ").trim();
  if (!text || text === "—") return "";
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
};

const isHistoryWorthyOp = (op: Record<string, unknown>) =>
  HISTORY_OP_KINDS.has(asString(op.kind));

const unitsOf = (text: string) =>
  [...text.matchAll(/\bA?\d{4,}\b/gi)]
    .map((match) => match[0].toUpperCase())
    .filter((value, index, all) => all.indexOf(value) === index)
    .join(" ");

const absenceTypeOf = (text: string) => {
  if (/поранен/i.test(text)) return "поранення";
  if (/сзч/i.test(text)) return "СЗЧ";
  if (/лікуван/i.test(text)) return "лікування";
  if (/відрядж/i.test(text)) return "відрядження";
  if (/відпуст/i.test(text)) return "відпустка";
  return "";
};

const firstDateOf = (text: string) =>
  text.match(/\d{2}\.\d{2}\.\d{4}/)?.[0] ?? "";

const indexOfOp = (op: Record<string, unknown>, payload: Record<string, unknown> | null) =>
  asString(payload?.nextIndex) ||
  asString(op.positionIndex) ||
  asString(op.after).match(/\d{6,7}/)?.[0] ||
  "";

/** Коротко: одна фраза на операцію, без before→after простирадла. */
const formatOpBrief = (op: Record<string, unknown>) => {
  const kind = asString(op.kind);
  const payload = asRecord(op.payload);
  const after = asString(op.after);
  const blob = `${asString(op.before)} ${after}`;
  const dest =
    asString(payload?.destination) ||
    asString(payload?.timesheetDestination) ||
    asString(payload?.documentsDest) ||
    unitsOf(blob);
  const index = indexOfOp(op, payload);
  const absence = absenceTypeOf(blob);
  const date = firstDateOf(blob);

  if (kind === "exclude_transfer") {
    return dest ? `ПЕРЕВ ${dest}` : "ПЕРЕВ";
  }
  if (kind === "move_to_disposition") {
    return dest ? `розпорядження ${dest}` : "розпорядження";
  }
  if (kind === "position_change") {
    return index ? `посада ${index}` : "зміна посади";
  }
  if (kind === "rank_change") {
    const rank = after.replace(/^.*→\s*/u, "").replace(/\s+/g, " ").trim();
    return rank ? `звання ${shortCell(rank, 28)}` : "звання";
  }
  if (kind === "arrival") {
    return index ? `прибуття · посада ${index}` : "прибуття";
  }
  if (kind === "shpo_occupant") {
    return index ? `ШПО ${index}` : "ШПО";
  }
  if (kind === "absent_close") {
    return absence ? `закрито ${absence}` : "закрито відсутність";
  }
  if (kind === "absent_upsert") {
    return [absence || "відсутність", date ? `з ${date}` : ""]
      .filter(Boolean)
      .join(" ");
  }
  if (kind === "other_manual") {
    const review = asString(payload?.reviewReason);
    if (
      review === "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH" ||
      /NEEDS_REVIEW/i.test(after)
    ) {
      return "скасування ПЕРЕВ (немає в sh)";
    }
    if (/CANCELLED|скасуван/i.test(`${after} ${asString(payload?.type)}`)) {
      return "скасовано ПЕРЕВ";
    }
    return shortCell(after, 40) || "інше";
  }
  return OP_KIND_LABEL[kind] || "зміна";
};

const asStringArray = (value: unknown) =>
  Array.isArray(value)
    ? value.map((item) => asString(item)).filter(Boolean)
    : [];

const padDay = (value: number) => String(value).padStart(2, "0");

const timesheetSummaryOf = (value: unknown) => {
  const rec = asRecord(value);
  if (!rec) return "";
  const runs = Array.isArray(rec.runs) ? rec.runs : [];
  return runs
    .map((item) => {
      const run = asRecord(item);
      const from = Number(run?.from);
      const to = Number(run?.to);
      const mark = asString(run?.mark);
      if (!from || !mark) return "";
      return from === to
        ? `${padDay(from)}: ${mark}`
        : `${padDay(from)}–${padDay(to)}: ${mark}`;
    })
    .filter(Boolean)
    .join(" · ");
};

/** Повний коментар для аркуша «10. Історія змін»: було / стало / чому / як. */
export const formatPersonApplyComment = (person: Record<string, unknown>) => {
  const will = asStringArray(person.ejoosWillDo);
  const impacts = Array.isArray(person.sheetImpacts) ? person.sheetImpacts : [];
  const sources = Array.isArray(person.sourceInfluences)
    ? person.sourceInfluences
    : [];
  const sheets = impacts
    .map((item) => {
      const rec = asRecord(item);
      if (!rec) return "";
      const effect = asString(rec.effect);
      if (effect === "untouched" || effect === "skip") return "";
      const row = asString(rec.rowHint);
      return `${asString(rec.sheetLabel)}: ${asString(rec.effectLabel)} — ${asString(rec.detail)}${row ? ` (${row})` : ""}`;
    })
    .filter(Boolean);
  const sourceLines = sources
    .map((item) => {
      const rec = asRecord(item);
      if (!rec) return "";
      return [
        asString(rec.sourceLabel),
        asString(rec.ref),
        asString(rec.event),
        asString(rec.effect),
      ]
        .filter(Boolean)
        .join(" · ");
    })
    .filter(Boolean);
  const timesheet = timesheetSummaryOf(person.timesheetPreview);
  return [
    `Було: ${asString(person.summaryBefore) || "—"}`,
    `Стало: ${asString(person.summaryAfter) || "—"}`,
    will.length ? `Що зроблено:\n${will.map((step) => `• ${step}`).join("\n")}` : "",
    timesheet ? `Табель: ${timesheet}` : "",
    sourceLines.length
      ? `З 1ПБ:\n${sourceLines.map((line) => `• ${line}`).join("\n")}`
      : "",
    sheets.length
      ? `Аркуші:\n${sheets.map((line) => `• ${line}`).join("\n")}`
      : "",
  ]
    .filter(Boolean)
    .join("\n");
};

const formatOpsPersonComment = (ops: Record<string, unknown>[]) => {
  const first = ops[0];
  if (!first) return "";
  const steps: string[] = [];
  const seen = new Set<string>();
  for (const op of ops) {
    const step = formatOpBrief(op);
    const why = asString(op.why);
    const line = why && why !== step ? `${step} (${why})` : step;
    if (!line || seen.has(line)) continue;
    seen.add(line);
    steps.push(line);
  }
  return [
    `Було: ${asString(first.before) || "—"}`,
    `Стало: ${asString(first.after) || "—"}`,
    steps.length ? `Що зроблено:\n${steps.map((step) => `• ${step}`).join("\n")}` : "",
  ]
    .filter(Boolean)
    .join("\n");
};

const isReviewOnlyOp = (op: Record<string, unknown>) => {
  const payload = asRecord(op.payload);
  const review = asString(payload?.reviewReason);
  const after = asString(op.after);
  return (
    asString(op.kind) === "other_manual" &&
    (review === "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH" ||
      /NEEDS_REVIEW/i.test(after))
  );
};

const personHistoryKey = (op: Record<string, unknown>) => {
  const id = asString(op.personId).replace(/\s+/g, "").toLowerCase();
  if (id) return `id:${id}`;
  const name = asString(op.fullName).replace(/\s+/g, " ").trim().toLocaleLowerCase("uk-UA");
  return name ? `name:${name}` : "";
};

const opDedupeKey = (op: Record<string, unknown>) => {
  const movement = asString(op.movementKey);
  if (movement) return `m:${movement}`;
  const id = asString(op.personId);
  const kind = asString(op.kind);
  const after = asString(op.after);
  if (id) return `p:${id}|${kind}|${after}`;
  return `n:${asString(op.fullName)}|${kind}|${after}`;
};

const makeHistoryRow = (
  version: BackendEjournalLiveVersion,
  type: string,
  description: string,
  comment = "",
): EjoosHistorySheetRow => ({
  versionLabel: versionLabelOf(version.version),
  date: formatUaDate(version.createdAt) || asString(version.asOfDate),
  type,
  description,
  comment,
  createdAtIso: version.createdAt || "",
});

const historyOpsOf = (
  version: BackendEjournalLiveVersion,
  skipMovementKeys?: Set<string>,
) => {
  const protocol = asRecord(version.changeProtocol);
  const kind = asString(protocol?.kind);
  if (kind === "rollback" || kind === "import" || kind.startsWith("anketa")) {
    return [];
  }
  const ops = Array.isArray(protocol?.ops) ? protocol.ops : [];
  if (kind !== "apply" && !ops.length) return [];
  const unique: unknown[] = [];
  const seen = new Set<string>();
  for (const item of ops) {
    const op = asRecord(item);
    if (!op || !isHistoryWorthyOp(op)) continue;
    const movementKey = asString(op.movementKey);
    if (movementKey && skipMovementKeys?.has(movementKey)) continue;
    const key = opDedupeKey(op);
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(op);
  }
  return unique;
};

/**
 * Один рядок на людину в межах застосування цієї версії.
 * Повний запис: було / стало / що зроблено / табель / джерела / аркуші.
 */
export const buildMergedHistorySheetRows = (
  versions: BackendEjournalLiveVersion[],
  options?: { skipMovementKeys?: Set<string> },
): EjoosHistorySheetRow[] => {
  const priorMovementKeys = new Set(options?.skipMovementKeys ?? []);
  const rows: EjoosHistorySheetRow[] = [];
  for (const version of versions) {
    const protocol = asRecord(version.changeProtocol);
    const storedPeople = Array.isArray(protocol?.people) ? protocol.people : [];
    if (storedPeople.length) {
      for (const item of storedPeople) {
        const person = asRecord(item);
        if (!person) continue;
        const name = formatPersonFull(
          asString(person.fullName),
          asString(person.personId),
        );
        if (!name) continue;
        rows.push(
          makeHistoryRow(
            version,
            typeLabelForKind("apply"),
            name,
            formatPersonApplyComment(person),
          ),
        );
      }
      continue;
    }
    const ops = historyOpsOf(version, priorMovementKeys).flatMap((item) => {
      const op = asRecord(item);
      return op ? [op] : [];
    });
    for (const op of ops) {
      const movementKey = asString(op.movementKey);
      if (movementKey) priorMovementKeys.add(movementKey);
    }
    const byPerson = new Map<string, Record<string, unknown>[]>();
    for (const op of ops) {
      const key = personHistoryKey(op);
      if (!key) continue;
      const list = byPerson.get(key) ?? [];
      list.push(op);
      byPerson.set(key, list);
    }
    for (const personOps of byPerson.values()) {
      const first = personOps[0];
      const name = formatPersonFull(
        asString(first.fullName),
        asString(first.personId),
      );
      if (!name) continue;
      rows.push(
        makeHistoryRow(
          version,
          typeLabelForKind(personOps.every(isReviewOnlyOp) ? "review" : "apply"),
          name,
          formatOpsPersonComment(personOps),
        ),
      );
    }
  }
  return rows;
};

const isReusableHistoryRow = (text: string) =>
  isOurVersionLabel(text) || isMisplacedIsoTimestamp(text);

const assignHistoryRows = (
  entries: EjoosHistorySheetRow[],
  columnA: Map<number, string>,
  lastContentRow: number,
) => {
  const reusable = [...columnA.entries()]
    .filter(([row, text]) => row > 1 && isReusableHistoryRow(text))
    .sort((a, b) => a[0] - b[0])
    .map(([row]) => row);
  let nextRow = lastContentRow + 1;
  while (isOfficialTemplateVersion(columnA.get(nextRow) || "")) {
    nextRow += 1;
  }
  const taken = new Set<number>();
  const assigned = entries.map((entry) => {
    const reused = reusable.find((row) => !taken.has(row));
    const row = reused ?? nextRow;
    taken.add(row);
    if (!reused) {
      nextRow = row + 1;
      while (
        taken.has(nextRow) ||
        isOfficialTemplateVersion(columnA.get(nextRow) || "")
      ) {
        nextRow += 1;
      }
    }
    return { entry, row };
  });
  const clearRows = reusable.filter((row) => !taken.has(row));
  return { assigned, clearRows };
};

/** Дописує нові рядки в кінець аркуша, не чіпаючи існуючу історію файлу. */
const assignHistoryRowsAppend = (
  entries: EjoosHistorySheetRow[],
  columnA: Map<number, string>,
  lastContentRow: number,
) => {
  let nextRow = lastContentRow + 1;
  while (isOfficialTemplateVersion(columnA.get(nextRow) || "")) {
    nextRow += 1;
  }
  const assigned = entries.map((entry) => {
    const row = nextRow;
    nextRow += 1;
    while (isOfficialTemplateVersion(columnA.get(nextRow) || "")) {
      nextRow += 1;
    }
    return { entry, row };
  });
  return { assigned, clearRows: [] as number[] };
};

const toWrites = (
  assigned: Array<{ entry: EjoosHistorySheetRow; row: number }>,
  clearRows: number[],
  styleSourceRow: number,
): ZipCellWrite[] => {
  const writes: ZipCellWrite[] = [];
  for (const target of assigned) {
    const values = [
      target.entry.versionLabel,
      target.entry.date,
      target.entry.type,
      target.entry.description,
      target.entry.comment,
    ];
    values.forEach((value, index) => {
      const column = index + 1;
      const isComment = column === 5;
      const text = isComment
        ? String(value ?? "")
            .replace(/[^\S\n]+/g, " ")
            .replace(/\n{3,}/g, "\n\n")
            .trim()
        : String(value ?? "").replace(/\s+/g, " ").trim();
      writes.push({
        row: target.row,
        column,
        value: text,
        styleSourceRow,
        styleSourceColumn: 4,
        copyNeighborStyle: true,
        wrapText: isComment && text.includes("\n"),
      });
    });
  }
  for (const row of clearRows) {
    for (let column = 1; column <= 5; column += 1) {
      writes.push({
        row,
        column,
        value: "",
        styleSourceRow,
        styleSourceColumn: 4,
        copyNeighborStyle: true,
      });
    }
  }
  return writes;
};

/** Останній імпорт повного файлу до обраної версії включно. */
export const findLastImportVersion = (
  versions: BackendEjournalLiveVersion[],
  upToVersion: number,
): number => {
  let lastImport = 0;
  for (const version of versions) {
    const versionNumber = version.version || 0;
    if (versionNumber > upToVersion) continue;
    const kind = asString(asRecord(version.changeProtocol)?.kind);
    if (kind === "import") {
      lastImport = Math.max(lastImport, versionNumber);
    }
  }
  return lastImport;
};

/**
 * Після відкату в БД лишаються всі версії, але аркуш «Історія змін» має
 * відображати лише застосування до точки відкату (+ нові після нього).
 */
export const selectVersionsForHistorySheet = (
  versions: BackendEjournalLiveVersion[],
  currentVersionNumber: number,
): BackendEjournalLiveVersion[] => {
  const sorted = [...versions]
    .filter((version) => version.version <= currentVersionNumber)
    .sort((left, right) => left.version - right.version);

  let restoredCutoff = 0;
  let resumeFrom = currentVersionNumber + 1;

  for (const version of sorted) {
    const protocol = asRecord(version.changeProtocol);
    if (asString(protocol?.kind) !== "rollback") continue;
    const restored = Number(protocol?.restoredFromVersion);
    if (Number.isFinite(restored) && restored > 0) {
      restoredCutoff = restored;
      resumeFrom = version.version + 1;
    }
  }

  if (!restoredCutoff) return sorted;

  return sorted.filter(
    (version) =>
      version.version <= restoredCutoff || version.version >= resumeFrom,
  );
};

export type EjoosChangeHistoryExportMode = "append" | "rebuild";

export type EjoosChangeHistoryExportOptions = {
  /**
   * append — лише ops після останнього імпорту файлу; існуючі рядки аркуша
   * не змінюються (експорт / застосування без відкату).
   * rebuild — повна синхронізація з протоколами БД (відкат).
   */
  mode?: EjoosChangeHistoryExportMode;
  /** Метадані експортованої версії (для перевірки «лише що імпортовано»). */
  currentVersion?: BackendEjournalLiveVersion | null;
};

/**
 * Дописує повні картки цього застосування в «10. Історія змін».
 */
export async function appendApplyHistorySheetRows(
  file: Blob | File,
  input: {
    version: number;
    appliedAt: string;
    people: unknown[];
  },
): Promise<Blob> {
  if (!input.people.length) return file;
  const state = await readHistorySheetState(file);
  if (!state) return file;
  const fakeVersion: BackendEjournalLiveVersion = {
    id: `apply-${input.version}`,
    unitLabel: "1ПБ",
    version: input.version,
    sha256: "",
    byteSize: 0,
    createdAt: input.appliedAt,
  };
  const entries = input.people.flatMap((item) => {
    const person = asRecord(item);
    if (!person) return [];
    const name = formatPersonFull(
      asString(person.fullName),
      asString(person.personId),
    );
    if (!name) return [];
    return [
      makeHistoryRow(
        fakeVersion,
        typeLabelForKind("apply"),
        name,
        formatPersonApplyComment(person),
      ),
    ];
  });
  if (!entries.length) return file;
  const { assigned, clearRows } = assignHistoryRowsAppend(
    entries,
    state.columnA,
    state.lastContentRow,
  );
  const writes = toWrites(assigned, clearRows, 2);
  if (!writes.length) return file;
  try {
    return await applyInlineStringWritesToWorkbook(
      file,
      EJOOS_CHANGE_HISTORY_SHEET,
      writes,
    );
  } catch {
    return file;
  }
}

/**
 * Дописує історію в «10. Історія змін» з протоколів застосувань.
 */
export async function appendEjoosChangeHistoryOnExport(
  file: Blob | File,
  versions: BackendEjournalLiveVersion[],
  upToVersion?: number,
  options?: EjoosChangeHistoryExportOptions,
): Promise<Blob> {
  const limit =
    upToVersion ??
    Math.max(0, ...versions.map((version) => version.version || 0));
  const mode = options?.mode ?? "append";
  if (mode === "append") {
    const currentKind = asString(
      asRecord(options?.currentVersion?.changeProtocol)?.kind,
    );
    if (currentKind === "import") return file;
  }
  let selected = selectVersionsForHistorySheet(versions, limit).sort(
    (a, b) => a.version - b.version,
  );
  if (!selected.length) return file;

  if (mode === "append") {
    const lastImport = findLastImportVersion(versions, limit);
    if (lastImport > 0) {
      selected = selected.filter((version) => version.version > lastImport);
    }
    if (!selected.length) return file;
  }

  const state = await readHistorySheetState(file);
  if (!state) return file;

  const maxExistingVersion =
    mode === "append" ? maxVersionLabelInSheet(state.columnA) : 0;
  const priorMovementKeys =
    mode === "append"
      ? collectProcessedMovementKeys(
          versions.filter((version) => version.version <= maxExistingVersion),
        )
      : new Set<string>();

  let entries = buildMergedHistorySheetRows(selected, {
    skipMovementKeys: priorMovementKeys,
  });
  if (!entries.length) return file;

  if (mode === "append") {
    entries = entries.filter(
      (entry) => versionNumberOfLabel(entry.versionLabel) > maxExistingVersion,
    );
    if (!entries.length) return file;
  }

  const { assigned, clearRows } =
    mode === "rebuild"
      ? assignHistoryRows(entries, state.columnA, state.lastContentRow)
      : assignHistoryRowsAppend(entries, state.columnA, state.lastContentRow);
  const writes = toWrites(assigned, clearRows, 2);
  if (!writes.length) return file;

  try {
    return await applyInlineStringWritesToWorkbook(
      file,
      EJOOS_CHANGE_HISTORY_SHEET,
      writes,
    );
  } catch {
    return file;
  }
}
