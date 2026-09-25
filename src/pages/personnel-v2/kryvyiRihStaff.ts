import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { previewValueToDisplay } from "../ejournal/ejournalUtils";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import { getPersonDisplayName } from "../personnel/personnelUtils";
import {
  getRosterPersonName,
  isPersonnelInStaffRoster,
} from "../personnel/personnelRosterMerge";

const foldPlace = (value: string) =>
  value
    .replace(/['’ʼ`]/g, "")
    .replace(/ё/g, "е")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const STREET_KRYVYI_RIH_RE =
  /(?:вул(?:иц[яі])?|пров(?:улок)?|просп(?:ект)?|пр-т|бул(?:ьвар)?|шосе)\.?\s+кривор[іi]зьк\w*/g;

const DISTINCTIVE_SETTLEMENT_RE =
  /(?:^|[^а-яіїєґ])(?:(?:м|місто|с|село|смт)\.?\s+)?(?:апостолов[а-яіїєґ]*|зеленодольськ[а-яіїєґ]*|глеюватк[а-яіїєґ]*|гречанопод[а-яіїєґ]*|гречан[іi]\s+под[а-яіїєґ]*|девладов[а-яіїєґ]*|нивотруд[а-яіїєґ]*|новолат[іi]в[а-яіїєґ]*|вакулов[а-яіїєґ]*)(?:[^а-яіїєґ]|$)/;

const OTHER_OBLAST_RE =
  /київськ|харків|одес|львів|полтав|запорі|дніпровськ|миколаїв|херсон|вінниц|житомир|черкас|чернігів|чернівец|сумськ|рівнен|тернопіль|хмельниц|івано-франк|закарпат|волин|кіровоград|кропивниц|донецьк|луганськ/;

const PLACE_KEY_RE = /адреса|проживан|зареєстр|реєстрац|народжен/i;

const placeKeyRank = (key: string) => {
  const folded = key.toLocaleLowerCase("uk-UA");
  if (/адреса|проживан|зареєстр|реєстрац/.test(folded)) return 0;
  if (/народжен/.test(folded)) return 1;
  return 2;
};

/** м. Кривий Ріг, Криворізький район або характерний населений пункт району. */
export const isKryvyiRihAreaText = (value: string) => {
  const text = foldPlace(value);
  if (!text) return false;

  if (/крив(?:ий|ого|ому|им)\s+р[іi]г/.test(text)) return true;

  const withoutStreet = text.replace(STREET_KRYVYI_RIH_RE, " ");
  if (/кривор[іi]зьк/.test(withoutStreet)) return true;

  if (OTHER_OBLAST_RE.test(text) && !/дніпропетров/.test(text)) return false;
  return DISTINCTIVE_SETTLEMENT_RE.test(text);
};

export const pickKryvyiRihPlace = (row: EjournalPreviewRow | null) => {
  if (!row) return "";
  const matches: Array<{ rank: number; value: string }> = [];
  for (const [key, raw] of Object.entries(row)) {
    if (key.startsWith("__") || key.startsWith("roster__column_")) continue;
    if (/призван/i.test(key) || !PLACE_KEY_RE.test(key)) continue;
    const value = previewValueToDisplay(raw).replace(/\s+/g, " ").trim();
    if (!value || !isKryvyiRihAreaText(value)) continue;
    matches.push({ rank: placeKeyRank(key), value });
  }
  matches.sort((left, right) => left.rank - right.rank);
  return matches[0]?.value ?? "";
};

export const indexKryvyiRihPlaces = (rows: EjournalPreviewRow[]) => {
  const index: Record<string, string> = {};
  for (const row of rows) {
    if (!isPersonnelInStaffRoster(row)) continue;
    const place = pickKryvyiRihPlace(row);
    if (!place) continue;
    for (const name of [getPersonDisplayName(row), getRosterPersonName(row)]) {
      const key = normalizeRosterMatchText(name);
      if (!key || index[key]) continue;
      index[key] = place;
    }
  }
  return index;
};
