const normalizeUkrainianKeyboardVariant = (value: string) =>
  value.replace(/[иіi]/giu, "і");

export const personnelSearchMatchesQuery = (
  searchableText: string,
  query: string,
) => {
  if (!query) return true;
  if (searchableText.includes(query)) return true;

  const normalizedText = normalizeUkrainianKeyboardVariant(searchableText);
  const normalizedQuery = normalizeUkrainianKeyboardVariant(query);
  if (normalizedText.includes(normalizedQuery)) return true;

  const tokens = normalizedQuery
    .split(/\s+/)
    .filter((token) => token.length >= 2);
  if (!tokens.length) return normalizedText.includes(normalizedQuery);

  // A multi-word query identifies one person. Never degrade it to a
  // surname-only match, because that returns unrelated namesakes.
  return tokens.every((token) => normalizedText.includes(token));
};
