const DATE_RE =
  /(\d{1,2})\s*[./.]\s*(\d{1,2})\s*[./.]\s*(\d{2,4})/;
const DATE_GLOBAL_RE = new RegExp(DATE_RE.source, "g");
const MILITARY_ID_RE =
  /(?:^|[^А-ЯІЇЄҐA-Z0-9])([А-ЯІЇЄҐA-Z]{2})\s*(\d{5,8})(?!\d)/iu;

export const looksLikeConscriptionOffice = (value: string) =>
  /ртцк|тцк\s*та\s*сп|\bтцк\b|військкомат|військомат|\bмвк\b|\bовк\b/i.test(
    value,
  );

export const normalizeAnketaConscriptionDate = (raw: string) => {
  const match = String(raw ?? "").match(DATE_RE);
  if (!match) return "";
  const day = match[1]!.padStart(2, "0");
  const month = match[2]!.padStart(2, "0");
  let year = match[3]!;
  if (year.length === 2) {
    year = Number(year) >= 50 ? `19${year}` : `20${year}`;
  }
  return `${day}.${month}.${year}`;
};

const collapse = (value: string) => value.replace(/\s+/g, " ").trim();

/** Розбирає рядок «РТЦК … дата» (і опційно серію квитка) на колонки анкети. */
export const parseAnketaConscriptionText = (raw: string) => {
  const text = collapse(String(raw ?? ""));
  if (!text) return { when: "", by: "", militaryId: "" };

  const when = normalizeAnketaConscriptionDate(text);
  const idMatch = text.match(MILITARY_ID_RE);
  const militaryId = idMatch
    ? `${idMatch[1]} ${idMatch[2]}`.replace(/\s+/g, " ")
    : "";

  let by = text
    .replace(DATE_GLOBAL_RE, " ")
    .replace(MILITARY_ID_RE, " ")
    .replace(
      /(?:мобілізований|призваний|коли|ким)[:\s]*/gi,
      " ",
    )
    .replace(/[.,;]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  by = by.replace(/(.+?)\s+\1$/u, "$1").trim();

  if (/^[А-ЯІЇЄҐA-Z]{2}\s*\d{5,8}$/iu.test(by)) by = "";

  return { when, by, militaryId };
};
