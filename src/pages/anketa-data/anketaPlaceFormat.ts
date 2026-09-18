const PHONE_RE = /(?:\+?38[\s-]*)?0(?:[\s-]*\d){9}/g;
const NAME_PART = "[А-ЯІЇЄҐа-яіїєґ'`’\\-]+";

const titleName = (value: string) => {
  const text = value.replace(/\s+/g, " ").trim();
  if (!text) return "";
  return text
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
};

const STREET_TYPE_CANON: Record<string, string> = {
  вул: "вул.",
  "вул.": "вул.",
  бульвар: "бульвар",
  бульв: "бульв.",
  "бульв.": "бульв.",
  проспект: "проспект",
  просп: "просп.",
  "просп.": "просп.",
  пр: "пр.",
  "пр.": "пр.",
  площа: "площа",
  пл: "пл.",
  "пл.": "пл.",
  провулок: "провулок",
  пров: "пров.",
  "пров.": "пров.",
};

const PLACE_KEEP_LOWER = new Set([
  "буд",
  "буд.",
  "кв",
  "кв.",
  "м",
  "м.",
  "с",
  "с.",
  "смт",
  "тел",
  "р-н",
  "обл",
]);

const titleStreetRemainder = (value: string) =>
  value.replace(/[А-Яа-яІіЇїЄєҐґЁё'`’.\-]+/gu, (word) => {
    const lower = word.toLocaleLowerCase("uk-UA");
    const type = STREET_TYPE_CANON[lower];
    if (type) return type;
    if (PLACE_KEEP_LOWER.has(lower)) return lower;
    if (/^\d/.test(word)) return word;
    return titleName(word);
  });

const collapsePlace = (value: string) =>
  value
    .replace(/[;|/]+/g, ",")
    .replace(/\s*,\s*/g, ", ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^,|,$/g, "")
    .trim();

const formatPhones = (phones: string[]) => {
  const unique: string[] = [];
  for (const raw of phones) {
    const digits = raw.replace(/\D/g, "");
    const local =
      digits.length === 12 && digits.startsWith("38")
        ? digits.slice(2)
        : digits;
    const pretty = local.length === 10 ? `+38${local}` : raw.trim();
    if (pretty && !unique.includes(pretty)) unique.push(pretty);
  }
  return unique.join(", ");
};

export const extractAnketaPhones = (raw: string) =>
  [...String(raw ?? "").matchAll(PHONE_RE)].map((match) => match[0]);

export const formatAnketaPhones = (phones: string[]) => formatPhones(phones);

const looksLikeCivilPlace = (value: string) =>
  /(?:смт|село|місто|селище|с\.|м\.|район|р-н|область|\bобл|вул\.?|буд\.?|кв\.?)/i.test(
    value,
  );

const tidyStreet = (value: string) => {
  let text = value
    .replace(/улиц[аяеи]\s*/gi, "вул. ")
    .replace(/вул\.\s*/gi, "вул. ")
    .replace(/(?<![вВ])ул\.\s*/gi, "вул. ")
    .replace(/буд\.\s*/gi, "буд. ")
    .replace(/кв\.\s*/gi, "кв. ")
    .replace(/кв\s+/gi, "кв. ")
    .replace(/\s+/g, " ")
    .trim();
  if (!/вул\./i.test(text)) {
    text = text.replace(
      /^([А-ЯІЇЄҐа-яіїєґ'`’\-]+)\s+(\d+)\b/u,
      "вул. $1 буд. $2",
    );
  }
  return titleStreetRemainder(
    text
      .replace(/([^\d\s,])\s+(буд\.)/gi, "$1, $2")
      .replace(/(буд\.\s*\d+)\s+([А-ЯІЇЄҐA-Za-z])\b/gu, "$1$2")
      .replace(/(буд\.\s*\d+[А-ЯІЇЄҐA-Za-z]?)\s+(\d+)\s+(кв\.)/gi, "$1, $2 $3")
      .replace(/(буд\.\s*\d+[А-ЯІЇЄҐA-Za-z]?)\s+(кв\.)/i, "$1, $2")
      .replace(/\s+/g, " ")
      .trim(),
  );
};

/** Канон: с./м./смт Назва, Район р-н, Область обл, тел. */
export const normalizeAnketaPlaceText = (
  raw: string,
  options?: { keepPhone?: boolean },
): string => {
  const original = collapsePlace(String(raw ?? ""));
  if (!original) return "";

  const phones = [...original.matchAll(PHONE_RE)].map((match) => match[0]);
  let text = original.replace(PHONE_RE, " ");

  const districtMatch = text.match(
    new RegExp(
      `(${NAME_PART}(?:ський|цький|зький))\\s+(?:район|р-н\\.?)`,
      "i",
    ),
  );
  const district = districtMatch?.[1]
    ? `${titleName(districtMatch[1])} р-н`
    : "";
  if (districtMatch) text = text.replace(districtMatch[0], " ");

  const regionMatch = text.match(
    new RegExp(
      `(${NAME_PART}(?:ська|цька|зька))\\s+(?:область|обл\\.?)`,
      "i",
    ),
  );
  const region = regionMatch?.[1] ? `${titleName(regionMatch[1])} обл` : "";
  if (regionMatch) text = text.replace(regionMatch[0], " ");

  let settlement = "";
  const smtMatch = text.match(
    new RegExp(
      `(?:смт|с\\.?\\s*м\\.?\\s*т\\.?|селище міського типу)\\s*\\.?\\s*(${NAME_PART}(?:\\s+${NAME_PART}){0,2})`,
      "i",
    ),
  );
  const villageMatch = text.match(
    new RegExp(
      `(?:^|[\\s,])(?:село|с\\.)\\s*(${NAME_PART}(?:\\s+(?!район|р-н|обл)${NAME_PART}){0,2})`,
      "i",
    ),
  );
  const cityMatch = text.match(
    new RegExp(
      `(?:місто|м\\.)\\s*(Кривий\\s+Ріг|Біла\\s+Церква|Кам['’\`]?янець-Подільський|${NAME_PART})`,
      "i",
    ),
  );
  if (smtMatch?.[1]) {
    settlement = `смт ${titleName(smtMatch[1])}`;
    text = text.replace(smtMatch[0], " ");
  } else if (villageMatch?.[1]) {
    settlement = `с. ${titleName(villageMatch[1])}`;
    text = text.replace(villageMatch[0], " ");
  } else if (cityMatch?.[1]) {
    settlement = `м. ${titleName(cityMatch[1])}`;
    text = text.replace(cityMatch[0], " ");
  }

  const rest = tidyStreet(collapsePlace(text));
  const parts = [settlement, district, region, rest].filter(Boolean);
  const keepPhone = options?.keepPhone !== false;
  const phonePart =
    keepPhone && phones.length ? `тел: ${formatPhones(phones)}` : "";
  const assembled = [...parts, phonePart].filter(Boolean).join(", ");
  if (settlement || district || region || (keepPhone && phones.length)) {
    return assembled;
  }
  if (!looksLikeCivilPlace(original)) return original;
  return tidyStreet(
    collapsePlace(
      original
        .replace(/с\.\s*/gi, "с. ")
        .replace(/м\.\s*/gi, "м. ")
        .replace(/смт\s*\.?/gi, "смт ")
        .replace(/район/gi, "р-н")
        .replace(/область/gi, "обл"),
    ),
  );
};
