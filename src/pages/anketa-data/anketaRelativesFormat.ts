import {
  extractAnketaPhones,
  formatAnketaPhones,
  normalizeAnketaPlaceText,
} from "./anketaPlaceFormat";

export const ANKETA_RELATIVE_LINE_PREFIXES = [
  "Сімейний стан",
  "Мати",
  "Батько",
  "Довірена особа",
  "Дитина",
] as const;

const WORD_MAP: Record<string, string> = {
  екатерина: "Катерина",
  екатерины: "Катерини",
  наталья: "Наталія",
  наталия: "Наталія",
  натальи: "Наталії",
  владимировна: "Володимирівна",
  володиміровна: "Володимирівна",
  володимировна: "Володимирівна",
  владимирович: "Володимирович",
  владимир: "Володимир",
  советская: "Радянська",
  советской: "Радянської",
  рядянська: "Радянська",
  рядянської: "Радянської",
  зацепина: "Зацепіна",
  зацепин: "Зацепін",
  лилівна: "Іллівна",
  ильинична: "Іллівна",
  ільінічна: "Іллівна",
};

const FAMILY_STATUS_MAP: Record<string, string> = {
  "цивільна дружина": "цивільна дружина",
  "цивільний чоловік": "цивільний чоловік",
  дружина: "дружина",
  чоловік: "чоловік",
  "не одружений": "не одружений",
  неодружений: "не одружений",
  "не одружена": "не одружена",
  холост: "не одружений",
  холостий: "не одружений",
  "не женат": "не одружений",
  "не замужем": "не одружена",
  женат: "одружений",
  замужем: "заміжня",
  одружений: "одружений",
  одружена: "одружена",
  заміжня: "заміжня",
};

const FAMILY_STATUS_PREFIXES = Object.keys(FAMILY_STATUS_MAP).sort(
  (left, right) => right.length - left.length,
);

const TRUSTED_ROLE_RE =
  /^(цивільна дружина|цивільний чоловік|дружина|чоловік|мати|батько|дитина)$/i;

const PATRONYMIC_RE = /(?:івна|овна|ївна|ович|евич|ійович|инична)$/i;

const collapse = (value: string) => value.replace(/\s+/g, " ").trim();

const ukrainianizeWords = (value: string) =>
  value.replace(/[А-Яа-яІіЇїЄєҐґЁёЫыЭэЪъ']+/gu, (word) => {
    const mapped = WORD_MAP[word.toLocaleLowerCase("uk-UA").replace(/ё/g, "е")];
    return mapped ?? word;
  });

const titleName = (value: string) =>
  collapse(value)
    .split(" ")
    .map((part) =>
      part
        .split("-")
        .map((chunk) =>
          chunk
            ? chunk.charAt(0).toLocaleUpperCase("uk-UA") +
              chunk.slice(1).toLocaleLowerCase("uk-UA")
            : chunk,
        )
        .join("-"),
    )
    .join(" ");

const normalizePersonNameKey = (value: string) =>
  titleName(value)
    .toLocaleLowerCase("uk-UA")
    .replace(/[.,;:]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const nameTokens = (value: string) =>
  normalizePersonNameKey(value)
    .split(" ")
    .filter((token) => token.length > 1);

const relativeNamesMatch = (left: string, right: string) => {
  const leftKey = normalizePersonNameKey(left);
  const rightKey = normalizePersonNameKey(right);
  if (!leftKey || !rightKey) return false;
  if (leftKey === rightKey) return true;
  const leftTokens = nameTokens(left);
  const rightTokens = nameTokens(right);
  if (leftTokens.length < 2 || rightTokens.length < 2) return false;
  const shared = Math.min(leftTokens.length, rightTokens.length);
  for (let index = 0; index < shared; index += 1) {
    if (leftTokens[index] !== rightTokens[index]) return false;
  }
  return true;
};

const normalizeAbsent = (value: string) => {
  const text = collapse(value);
  if (!text) return "";
  const lower = text.toLocaleLowerCase("uk-UA");
  if (/^нема\b.*не\s+спілку/i.test(lower) || (/не\s+спілку/.test(lower) && /нема/.test(lower))) {
    return "немає, не спілкується";
  }
  if (/^не\s+спілку/.test(lower)) return "не спілкується";
  if (/^(немає|нема|нет|ні|відсутн\w*|—|-+)$/i.test(lower)) return "немає";
  return text;
};

const stripDanglingAbsent = (value: string) =>
  collapse(
    value.replace(
      /(?<=\S)(?:[\s,;]+)(?:немає|нема|нет)(?!\s+спілку)[\s,;]*$/gi,
      " ",
    ),
  );

const padDatePart = (value: string) => value.padStart(2, "0");

const expandShortYear = (year: string) => {
  if (year.length !== 2) return year;
  return Number(year) >= 50 ? `19${year}` : `20${year}`;
};

const extractDeathNote = (value: string) => {
  const text = collapse(value);
  const feminine = /померл[аи]/i.test(text);
  const masculine = !feminine && /(?:^|\s)помер(?:\s|$)/i.test(text);
  if (!feminine && !masculine) return { note: "", rest: text };
  const yearMatch = text.match(
    /померл[аи]?\s+(?:(\d{1,2})(?:-го)?\s+(?:о|у|в)?\s*)?(?:(\d{1,2})[./.]\s*)?(\d{4})\s*р(?:ік|оці|оку|\.)?/i,
  );
  const fullDate = text.match(
    /померл[аи]?[^\d]*(\d{1,2})\s*[./.]\s*(\d{1,2})\s*[./.]\s*(\d{2,4})/i,
  );
  let note = feminine ? "померла" : "помер";
  if (fullDate) {
    const year = expandShortYear(fullDate[3]!);
    note = `${note} ${padDatePart(fullDate[1]!)}.${padDatePart(fullDate[2]!)}.${year}`;
  } else if (yearMatch?.[3] && feminine) {
    note = `померла у ${yearMatch[3]} р.`;
  } else if (yearMatch?.[3]) {
    note = `помер у ${yearMatch[3]} р.`;
  }
  const rest = text
    .replace(
      /(?:мати|батько)?\s*померл[аи]?\s+(?:\d{1,2}(?:-го)?\s+(?:о|у|в)?\s*)?(?:\d{1,2}\s*[./.]\s*)?\d{2,4}\s*р(?:ік|оці|оку|\.)?/gi,
      " ",
    )
    .replace(/померл[аи]/gi, " ")
    .replace(/(?:^|\s)помер(?=\s|$)/gi, " ")
    .trim();
  return { note, rest };
};

const splitNameAndPlace = (value: string) => {
  const words = collapse(value).split(" ").filter(Boolean);
  if (!words.length) return { name: "", place: "" };
  const patronymicAt = words.findIndex((word) => PATRONYMIC_RE.test(word));
  if (patronymicAt >= 1) {
    return {
      name: titleName(words.slice(0, patronymicAt + 1).join(" ")),
      place: words.slice(patronymicAt + 1).join(" "),
    };
  }
  const streetAt = words.findIndex((word, index) => {
    if (/^(вул|м|смт|буд|кв|бульвар|бульв|проспект|просп|пр|площа|пл)\.?$/i.test(word)) {
      return true;
    }
    return index >= 2 && /^\d+$/.test(word);
  });
  if (streetAt > 0) {
    return {
      name: titleName(words.slice(0, streetAt).join(" ")),
      place: words.slice(streetAt).join(" "),
    };
  }
  return { name: titleName(value), place: "" };
};

const extractBirthYear = (value: string) => {
  const text = collapse(value);
  const full = text.match(
    /(\d{1,2})\s*[./]\s*(\d{1,2})\s*[./]\s*(\d{2,4})\s*(?:р\.?\s*н\.?|року народження)?/i,
  );
  if (full) {
    return {
      birth: `${padDatePart(full[1]!)}.${padDatePart(full[2]!)}.${expandShortYear(full[3]!)} р.н.`,
      rest: collapse(text.replace(full[0], " ")),
    };
  }
  const yearOnly = text.match(
    /((?:19|20)\d{2})\s*(?:р\.?\s*н\.?|року народження)/i,
  );
  if (yearOnly) {
    return {
      birth: `${yearOnly[1]} р.н.`,
      rest: collapse(text.replace(yearOnly[0], " ")),
    };
  }
  const tokens = text.split(" ");
  const yearAt = tokens.findIndex((token, index) => {
    if (!/^(?:19[2-9]\d|20[0-2]\d)$/.test(token)) return false;
    const previous = tokens[index - 1]?.toLocaleLowerCase("uk-UA") ?? "";
    return !/^(буд\.?|кв\.?|вул\.?|№)$/i.test(previous);
  });
  if (yearAt < 0) return { birth: "", rest: text };
  const year = tokens[yearAt]!;
  tokens.splice(yearAt, 1);
  return { birth: `${year} р.н.`, rest: collapse(tokens.join(" ")) };
};

const findFamilyStatusHit = (lowered: string) => {
  for (const prefix of FAMILY_STATUS_PREFIXES) {
    let from = 0;
    while (from <= lowered.length) {
      const index = lowered.indexOf(prefix, from);
      if (index < 0) break;
      const before = index === 0 || /[\s,;:]/.test(lowered[index - 1] ?? "");
      const afterIndex = index + prefix.length;
      const after =
        afterIndex >= lowered.length ||
        /[\s,;:]/.test(lowered[afterIndex] ?? "");
      if (before && after) return { prefix, index };
      from = index + 1;
    }
  }
  return null;
};

const parseFamilyStatus = (raw: string) => {
  const ukrainian = ukrainianizeWords(
    collapse(raw).replace(/^сімейний стан:\s*/i, ""),
  );
  const lowered = ukrainian.toLocaleLowerCase("uk-UA");
  const hit = findFamilyStatusHit(lowered);
  if (!hit) {
    const spouse = ukrainian ? parsePersonParts(ukrainian) : parsePersonParts("");
    const line = spouse.name || spouse.place || spouse.phones
      ? formatPersonParts(spouse)
      : lowered;
    return {
      status: "",
      line,
      spouse,
      role: "",
    };
  }
  const status = FAMILY_STATUS_MAP[hit.prefix] ?? hit.prefix;
  const rest = collapse(
    ukrainian.replace(new RegExp(hit.prefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i"), " "),
  );
  const spouse = rest ? parsePersonParts(rest) : parsePersonParts("");
  const role = spouseRoleFromStatus(status, spouse.name);
  const line = rest ? `${status}, ${formatPersonParts(spouse)}` : status;
  return { status, line, spouse, role };
};

const takeBirth = (value: string, existing = "") => {
  const extracted = extractBirthYear(value);
  return {
    birth: existing || extracted.birth,
    rest: extracted.rest,
  };
};

type RelativePersonParts = {
  name: string;
  birth: string;
  note: string;
  place: string;
  phones: string;
  absent: string;
};

const parsePersonParts = (raw: string): RelativePersonParts => {
  const empty = {
    name: "",
    birth: "",
    note: "",
    place: "",
    phones: "",
    absent: "",
  };
  const ukrainian = ukrainianizeWords(collapse(raw));
  if (!ukrainian) return empty;
  const phones = formatAnketaPhones(extractAnketaPhones(ukrainian));
  const withoutPhones = stripDanglingAbsent(
    ukrainian
      .replace(/(?:\+?38[\s-]*)?0(?:[\s-]*\d){9}/g, " ")
      .replace(/тел[.:]?\s*/gi, " ")
      .replace(/[,;]+/g, " "),
  );
  const absent = normalizeAbsent(withoutPhones);
  if (
    absent === "немає" ||
    absent === "немає, не спілкується" ||
    absent === "не спілкується"
  ) {
    return { ...empty, phones, absent };
  }
  const { note, rest } = extractDeathNote(collapse(withoutPhones));
  const firstBirth = extractBirthYear(rest);
  const { name, place } = splitNameAndPlace(firstBirth.rest);
  const fromPlace = takeBirth(place, firstBirth.birth);
  const formattedPlace = stripDanglingAbsent(
    normalizeAnketaPlaceText(fromPlace.rest, { keepPhone: false }),
  );
  const fromFormatted = takeBirth(formattedPlace, fromPlace.birth);
  return {
    name,
    birth: fromFormatted.birth,
    note,
    place: stripDanglingAbsent(fromFormatted.rest),
    phones,
    absent: "",
  };
};

const formatPersonParts = (parts: RelativePersonParts) => {
  if (parts.absent) {
    return [parts.absent, parts.phones ? `тел: ${parts.phones}` : ""]
      .filter(Boolean)
      .join(", ");
  }
  return [
    parts.name,
    parts.birth,
    parts.note,
    parts.place,
    parts.phones ? `тел: ${parts.phones}` : "",
  ]
    .filter(Boolean)
    .join(", ");
};

const phonesOverlap = (left: string, right: string) => {
  if (!left || !right) return false;
  const leftSet = new Set(left.split(", ").filter(Boolean));
  return right.split(", ").some((phone) => leftSet.has(phone));
};

const sameRelativePerson = (
  left: RelativePersonParts,
  right: RelativePersonParts,
) => {
  if (left.absent || right.absent) return false;
  if (relativeNamesMatch(left.name, right.name)) return true;
  return Boolean(
    phonesOverlap(left.phones, right.phones) &&
      left.name &&
      right.name &&
      nameTokens(left.name)[0] === nameTokens(right.name)[0],
  );
};

const spouseRoleFromStatus = (status: string, spouseName: string) => {
  if (/цивільна дружина/i.test(status)) return "цивільна дружина";
  if (/цивільний чоловік/i.test(status)) return "цивільний чоловік";
  if (/^дружина/i.test(status)) return "дружина";
  if (/^чоловік/i.test(status)) return "чоловік";
  if (/одружений|одружена|заміжня/i.test(status)) {
    if (/(?:івна|ївна|овна|инична)$/i.test(spouseName)) return "дружина";
    if (/(?:ович|евич|ійович)$/i.test(spouseName)) return "чоловік";
    return "дружина";
  }
  return "";
};

const trustedPersonRole = (raw: string) => {
  const text = collapse(raw).replace(/\.+$/, "");
  if (TRUSTED_ROLE_RE.test(text)) return text.toLocaleLowerCase("uk-UA");
  return "";
};

const resolveTrustedPersonLine = (
  raw: string,
  family: ReturnType<typeof parseFamilyStatus>,
  mother: RelativePersonParts,
  father: RelativePersonParts,
  child: RelativePersonParts,
) => {
  const alreadyRole = trustedPersonRole(raw);
  if (alreadyRole) return alreadyRole;
  const trusted = parsePersonParts(raw);
  if (!trusted.name && !trusted.phones) return formatPersonParts(trusted);
  if (sameRelativePerson(trusted, mother)) return "мати";
  if (sameRelativePerson(trusted, father)) return "батько";
  if (family.role && sameRelativePerson(trusted, family.spouse)) {
    return family.role;
  }
  if (sameRelativePerson(trusted, child)) return "дитина";
  return formatPersonParts(trusted);
};

const parseRelativeSections = (raw: string) => {
  const text = String(raw ?? "").replace(/\r\n/g, "\n").trim();
  const values: Record<(typeof ANKETA_RELATIVE_LINE_PREFIXES)[number], string> =
    {
      "Сімейний стан": "",
      Мати: "",
      Батько: "",
      "Довірена особа": "",
      Дитина: "",
    };
  const matcher = new RegExp(
    `(^|\\n)\\s*(${ANKETA_RELATIVE_LINE_PREFIXES.join("|")}):\\s*`,
    "gi",
  );
  const hits = [...text.matchAll(matcher)];
  if (!hits.length) return values;
  for (let index = 0; index < hits.length; index += 1) {
    const hit = hits[index]!;
    const label = ANKETA_RELATIVE_LINE_PREFIXES.find(
      (prefix) => prefix.toLowerCase() === hit[2]!.toLowerCase(),
    );
    if (!label) continue;
    const start = (hit.index ?? 0) + hit[0].length;
    const end = hits[index + 1]?.index ?? text.length;
    values[label] = text.slice(start, end).replace(/\n+/g, " ").trim();
  }
  return values;
};

/** Канон родичів: ПІБ, примітка, адреса, телефон в кінці; українською. */
export const normalizeAnketaRelativesText = (raw: string) => {
  const sections = parseRelativeSections(raw);
  const hasAny = ANKETA_RELATIVE_LINE_PREFIXES.some((prefix) =>
    sections[prefix].trim(),
  );
  if (!hasAny) return String(raw ?? "").trim();

  const family = parseFamilyStatus(sections["Сімейний стан"]);
  const mother = parsePersonParts(sections["Мати"]);
  const father = parsePersonParts(sections["Батько"]);
  const child = parsePersonParts(sections["Дитина"]);
  const trustedLine = resolveTrustedPersonLine(
    sections["Довірена особа"],
    family,
    mother,
    father,
    child,
  );
  const lines = [
    `Сімейний стан: ${family.line}`.trim(),
    `Мати: ${formatPersonParts(mother)}`.trim(),
    `Батько: ${formatPersonParts(father)}`.trim(),
    `Довірена особа: ${trustedLine}`.trim(),
    `Дитина: ${formatPersonParts(child)}`.trim(),
  ];
  const usedPhones = new Set(
    extractAnketaPhones(lines.join("\n")).map(
      (phone) => formatAnketaPhones([phone]),
    ),
  );
  const leftoverPhones = formatAnketaPhones(extractAnketaPhones(raw))
    .split(", ")
    .filter((phone) => phone && !usedPhones.has(phone));
  const trustedIsRole = Boolean(trustedPersonRole(trustedLine));
  if (leftoverPhones.length && !trustedIsRole) {
    const extra = leftoverPhones.join(", ");
    const trusted = lines[3] ?? "Довірена особа:";
    if (/тел:/.test(trusted)) {
      lines[3] = `${trusted}, ${extra}`;
    } else if (trusted === "Довірена особа:") {
      lines[3] = `Довірена особа: тел: ${extra}`;
    } else {
      lines[3] = `${trusted}, тел: ${extra}`;
    }
  }

  return lines.join("\n");
};
