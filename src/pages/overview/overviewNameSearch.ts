import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import { normalizePersonSearchKeyboard } from "../personnel/personnelSearch";

const PATRONYMIC_SUFFIX =
  /^(ович|овіч|евич|евіч|йович|йовіч|івна|ївна|овна|овна|евна|евна|ич|іч)$/;

/** Звання на початку рядка (копія зі штатки: «старший солдат КОВАЛЬ …»). */
const RANK_NAME_PREFIX =
  /^(?:головний\s+майстер[-\s]?сержант|старший\s+майстер[-\s]?сержант|майстер[-\s]?сержант|штаб[-\s]?сержант|головний\s+сержант|старший\s+сержант|молодший\s+сержант|сержант|старший\s+солдат|солдат|старший\s+матрос|матрос|старший\s+лейтенант|молодший\s+лейтенант|лейтенант|капітан|підполковник|полковник|майор|бригадний\s+генерал|генерал[-\s]?лейтенант|генерал[-\s]?майор|генерал|старшина)\s+/i;

const RANK_ONLY_CELL =
  /^(?:головний\s+майстер[-\s]?сержант|старший\s+майстер[-\s]?сержант|майстер[-\s]?сержант|штаб[-\s]?сержант|головний\s+сержант|старший\s+сержант|молодший\s+сержант|сержант|старший\s+солдат|солдат|старший\s+матрос|матрос|старший\s+лейтенант|молодший\s+лейтенант|лейтенант|капітан|підполковник|полковник|майор|бригадний\s+генерал|генерал[-\s]?лейтенант|генерал[-\s]?майор|генерал|старшина)$/i;

export const stripOverviewRankPrefix = (value: string) =>
  value.replace(RANK_NAME_PREFIX, "").trim();

const isRankOnlyCell = (value: string) =>
  RANK_ONLY_CELL.test(value.replace(/\s+/g, " ").trim());

const countNameWords = (line: string) =>
  line
    .replace(/\([^)]*\)/g, " ")
    .trim()
    .split(/\s+/)
    .filter(Boolean).length;

/** Рядок зі штатки: ПІБ, далі таб і позивний / звання. */
export const pickOverviewPastedPersonName = (line: string) => {
  const cells = line
    .split(/\t+/)
    .map((cell) => cell.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  if (!cells.length) return "";
  if (cells.length === 1) return stripOverviewRankPrefix(cells[0]!);

  let best = "";
  let bestScore = Number.NEGATIVE_INFINITY;
  for (const cell of cells) {
    const stripped = stripOverviewRankPrefix(cell);
    if (!stripped || isRankOnlyCell(cell) || isRankOnlyCell(stripped)) continue;
    const words = countNameWords(stripped);
    let score = words;
    if (words >= 2) score += 10;
    if (words <= 1) score -= 8;
    if (score > bestScore) {
      best = stripped;
      bestScore = score;
    }
  }
  return best || stripOverviewRankPrefix(cells[0]!);
};

export const normalizeOverviewName = (value: unknown) =>
  normalizePersonSearchKeyboard(
    normalizeRosterMatchText(value)
      .replace(/[ьъ]/g, "")
      .replace(/ё/g, "е")
      .replace(/[`´]/g, "")
      .replace(/-/g, " ")
      .replace(/\s+/g, " ")
      .trim(),
  );

/** Two single-word lines from mobile paste → one «Прізвище Ім'я», not OR-list. */
const isMultiNamePaste = (lines: string[]) => {
  if (lines.length <= 1) return false;
  const fullNameLines = lines.filter((line) => countNameWords(line) >= 2);
  if (fullNameLines.length >= 2) return true;
  return lines.length >= 3;
};

const splitOverviewPasteLines = (text: string) => {
  if (/[\n;]/.test(text)) {
    return text
      .split(/[\n;]+/)
      .map((chunk) => chunk.trim())
      .filter(Boolean);
  }
  if (text.includes(",") && /\s/.test(text)) {
    return text.split(/[\n;]+/).flatMap((chunk) =>
      chunk.includes(",") && chunk.trim().split(/\s+/).length >= 2
        ? [chunk]
        : chunk.split(",").map((part) => part.trim()),
    );
  }
  return [text];
};

/** Split pasted list: one name per line, or `;` / `,` separated. */
export const parseOverviewNameQueries = (raw: string) => {
  const text = raw.replace(/\r/g, "").trim();
  if (!text) return [] as string[];

  const candidateLines = splitOverviewPasteLines(text)
    .map((part) =>
      pickOverviewPastedPersonName(part.replace(/^[\s\-•·\d.)]+/u, "").trim()),
    )
    .filter(Boolean);
  const parts = isMultiNamePaste(candidateLines)
    ? candidateLines
    : [candidateLines.join(" ").replace(/\s+/g, " ").trim()].filter(Boolean);

  const seen = new Set<string>();
  const result: string[] = [];
  for (const part of parts) {
    const key = normalizeOverviewName(part);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(part);
  }
  return result;
};

export const overviewNameQueryTokenCount = (queryName: string) =>
  normalizeOverviewName(queryName)
    .split(/\s+/)
    .filter((token) => token.length >= 2).length;

const tokensMatch = (part: string, token: string) => {
  if (part === token) return true;
  const longer = part.length >= token.length ? part : token;
  const shorter = part.length >= token.length ? token : part;
  if (shorter.length < 2 || !longer.startsWith(shorter)) return false;
  const rest = longer.slice(shorter.length);
  if (PATRONYMIC_SUFFIX.test(rest)) return false;
  return rest.length <= 4;
};

export const overviewNameMatchesQuery = (personName: string, queryName: string) => {
  const person = normalizeOverviewName(stripOverviewRankPrefix(personName));
  const query = normalizeOverviewName(stripOverviewRankPrefix(queryName));
  if (!person || !query) return false;
  if (person === query) return true;

  const personTokens = person.split(/\s+/).filter(Boolean);
  const queryTokens = query.split(/\s+/).filter((token) => token.length >= 2);
  if (!queryTokens.length) return person.includes(query);

  if (person.includes(query)) return true;
  if (personTokens.length >= 2 && query.includes(person)) return true;

  const significantPersonTokens = personTokens.filter((token) => token.length >= 2);

  if (queryTokens.length >= 2) {
    if (
      queryTokens.every((token) =>
        personTokens.some((part) => tokensMatch(part, token)),
      )
    ) {
      return true;
    }
    return (
      significantPersonTokens.length >= 2 &&
      significantPersonTokens.every((part) =>
        queryTokens.some((token) => tokensMatch(part, token)),
      )
    );
  }

  return personTokens.some((part) => tokensMatch(part, queryTokens[0]!));
};

/** Pre-normalized blob for global search — built once per row set, not per keystroke. */
export const buildOverviewRowSearchText = (
  row: {
    name?: string;
    externalId?: string;
    rank?: string;
    positionTitle?: string;
    unit?: string;
    statusLabel?: string;
    fighterDirection?: string;
    fighterExitDate?: string;
    fighterReturnDate?: string;
    fighterTotalDays?: string;
    fighterStatus?: string;
    callSign?: string;
  },
  documentLabels = "",
) =>
  [
    row.name,
    row.externalId,
    row.rank,
    row.positionTitle,
    row.unit,
    row.statusLabel,
    row.fighterDirection,
    row.fighterExitDate,
    row.fighterReturnDate,
    row.fighterTotalDays,
    row.fighterStatus,
    row.callSign,
    documentLabels,
  ]
    .join(" ")
    .split(" ")
    .map(normalizeRosterMatchText)
    .join(" ");
