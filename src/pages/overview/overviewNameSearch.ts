import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";

const PATRONYMIC_SUFFIX = /^(ович|евич|йович|івна|ївна|овна|евна|ич)$/;

export const normalizeOverviewName = (value: unknown) =>
  normalizeRosterMatchText(value)
    .replace(/[ьъ]/g, "")
    .replace(/ё/g, "е")
    .replace(/[`´]/g, "")
    .replace(/-/g, " ")
    .replace(/\s+/g, " ")
    .trim();

/** Split pasted list: one name per line, or `;` / `,` separated. */
export const parseOverviewNameQueries = (raw: string) => {
  const text = raw.replace(/\r/g, "").trim();
  if (!text) return [] as string[];

  const parts =
    /[\n;]/.test(text) || (text.includes(",") && /\s/.test(text))
      ? text.split(/[\n;]+/).flatMap((chunk) =>
          chunk.includes(",") && chunk.trim().split(/\s+/).length >= 2
            ? [chunk]
            : chunk.split(","),
        )
      : [text];

  const seen = new Set<string>();
  const result: string[] = [];
  for (const part of parts) {
    const cleaned = part.replace(/^[\s\-•·\d.)]+/u, "").trim();
    if (!cleaned) continue;
    const key = normalizeOverviewName(cleaned);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(cleaned);
  }
  return result;
};

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
  const person = normalizeOverviewName(personName);
  const query = normalizeOverviewName(queryName);
  if (!person || !query) return false;
  if (person === query || person.includes(query) || query.includes(person)) {
    return true;
  }

  const personTokens = person.split(/\s+/).filter(Boolean);
  const queryTokens = query.split(/\s+/).filter((token) => token.length >= 2);
  if (!queryTokens.length) return false;

  return queryTokens.every((token) =>
    personTokens.some((part) => tokensMatch(part, token)),
  );
};
