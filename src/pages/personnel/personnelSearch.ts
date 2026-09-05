export const personnelSearchMatchesQuery = (
  searchableText: string,
  query: string,
) => {
  if (!query) return true;
  if (searchableText.includes(query)) return true;

  const tokens = query.split(/\s+/).filter((token) => token.length >= 2);
  if (!tokens.length) return searchableText.includes(query);

  // A multi-word query identifies one person. Never degrade it to a
  // surname-only match, because that returns unrelated namesakes.
  return tokens.every((token) => searchableText.includes(token));
};
