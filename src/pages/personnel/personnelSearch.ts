/** Latin keys that look like Cyrillic on a Ukrainian keyboard layout. */
const LATIN_TO_CYRILLIC: Record<string, string> = {
  A: "А",
  a: "а",
  B: "В",
  b: "в",
  C: "С",
  c: "с",
  E: "Е",
  e: "е",
  H: "Н",
  h: "н",
  I: "І",
  K: "К",
  k: "к",
  M: "М",
  m: "м",
  O: "О",
  o: "о",
  P: "Р",
  p: "р",
  T: "Т",
  t: "т",
  X: "Х",
  x: "х",
  Y: "У",
  y: "у",
};

const normalizeLatinHomoglyphs = (value: string) =>
  [...value].map((char) => LATIN_TO_CYRILLIC[char] ?? char).join("");

export const normalizePersonSearchKeyboard = (value: string) =>
  normalizeLatinHomoglyphs(value).replace(/[иіi]/giu, "і");

const hasWildcardMarkers = (query: string) => /[*?]/.test(query);

const escapeRegExp = (value: string) =>
  value.replace(/[.+^${}()|[\]\\]/g, "\\$&");

/** Одна літера + * — типовий «усі на букву А» (початок ПІБ/позивного). */
const isSingleLetterPrefixWildcard = (query: string) =>
  /^\p{L}\*$/u.test(normalizePersonSearchKeyboard(query.trim()));

/**
 * Word-style wildcards: ? — один символ (без пробілів), * — будь-яка кількість.
 * Шаблон шукається як підрядок (як «Знайти» у Word з wildcards).
 */
export const personnelWildcardQueryToRegExp = (pattern: string) => {
  const normalized = normalizePersonSearchKeyboard(pattern.trim());
  let body = "";
  for (const char of normalized) {
    if (char === "*") body += ".*";
    else if (char === "?") body += "[^\\s]";
    else body += escapeRegExp(char);
  }
  return new RegExp(body, "iu");
};

export type PersonnelSearchMatchOptions = {
  /** ПІБ — основне поле для масок. */
  primaryText?: string;
  /** Позивний. */
  callSignText?: string;
};

const textMatchesWildcard = (haystack: string, pattern: RegExp) =>
  Boolean(haystack.trim()) && pattern.test(haystack);

const textMatchesLetterPrefix = (haystack: string, query: string) => {
  const normalized = normalizePersonSearchKeyboard(haystack.trim());
  const letter = normalizePersonSearchKeyboard(query.trim()).charAt(0);
  if (!normalized || !letter) return false;
  const first = [...normalized][0] ?? "";
  return (
    first.toLocaleLowerCase("uk-UA") === letter.toLocaleLowerCase("uk-UA")
  );
};

export const personnelSearchMatchesWildcardQuery = (
  searchableText: string,
  query: string,
  options?: PersonnelSearchMatchOptions,
) => {
  const primaryText = normalizePersonSearchKeyboard(
    options?.primaryText?.trim() ?? "",
  );
  const callSignText = normalizePersonSearchKeyboard(
    options?.callSignText?.trim() ?? "",
  );

  if (isSingleLetterPrefixWildcard(query)) {
    return [primaryText, callSignText].some((text) =>
      textMatchesLetterPrefix(text, query),
    );
  }

  const pattern = personnelWildcardQueryToRegExp(query);
  const normalizedText = normalizePersonSearchKeyboard(searchableText);

  if (options?.primaryText !== undefined) {
    return [primaryText, callSignText].some((text) =>
      textMatchesWildcard(text, pattern),
    );
  }

  return textMatchesWildcard(normalizedText, pattern);
};

export const personnelSearchMatchesQuery = (
  searchableText: string,
  query: string,
  options?: PersonnelSearchMatchOptions,
) => {
  if (!query) return true;

  if (hasWildcardMarkers(query)) {
    return personnelSearchMatchesWildcardQuery(
      searchableText,
      query,
      options,
    );
  }

  if (searchableText.includes(query)) return true;

  const normalizedText = normalizePersonSearchKeyboard(searchableText);
  const normalizedQuery = normalizePersonSearchKeyboard(query);
  if (normalizedText.includes(normalizedQuery)) return true;

  const tokens = normalizedQuery
    .split(/\s+/)
    .filter((token) => token.length >= 2);
  if (!tokens.length) return normalizedText.includes(normalizedQuery);

  // A multi-word query identifies one person. Never degrade it to a
  // surname-only match, because that returns unrelated namesakes.
  return tokens.every((token) => normalizedText.includes(token));
};

/** Підказка, якщо маска з «?» не дала результатів — часта плутанина з «*». */
export const personnelSearchEmptyHint = (query: string) => {
  const trimmed = query.trim();
  if (!/[*?]/.test(trimmed)) return "";

  const normalized = normalizePersonSearchKeyboard(trimmed);
  if (/^.\?$/u.test(normalized) && !normalized.includes("*")) {
    const letter = normalized.charAt(0).toLocaleUpperCase("uk-UA");
    return `«${trimmed}» — «${letter}» + один символ (як «${letter}л»). Щоб знайти всіх на «${letter}», введіть ${letter}*`;
  }

  if (normalized.includes("*") || normalized.includes("?")) {
    return "Маски як у Word: ? — один символ (без пробілу), * — будь-яка кількість символів.";
  }

  return "";
};
