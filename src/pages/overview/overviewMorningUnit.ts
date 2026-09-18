import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";

// Do not use prefix matching: neighbouring platoons can share a long prefix.
export const morningUnitKey = (label: string) => {
  const normalized = normalizeRosterMatchText(label).replace(/\s+/g, " ").trim();
  const company = normalized.match(/^(\d+)\s*(?:піхотна\s+)?рота$/u)
    ?? normalized.match(/^(\d+)\s*пр$/u);
  return company ? `company:${company[1]}` : normalized;
};

export const morningUnitMatches = (unit: string, selected: string) =>
  Boolean(morningUnitKey(selected)) && morningUnitKey(unit) === morningUnitKey(selected);
