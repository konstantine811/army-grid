import JSZip from "jszip";
import { stripRedColorInWordZip } from "../documents/wordXml";
import type { PreAnketaFieldKey, PreAnketaFormFields } from "./preAnketaFields";

const PRE_ANKETA_TEMPLATE_URL = `${import.meta.env.BASE_URL}templates/pre-anketa-form.docx`;

const PARAGRAPH_TARGETS: Partial<Record<PreAnketaFieldKey, number>> = {
  callsign: 3,
  fullName: 7,
  rank: 14,
  rnokpp: 17,
  iban: 20,
  passport: 23,
  driverLicense: 29,
  birthDate: 32,
  birthPlace: 35,
  registrationAddress: 39,
  actualAddress: 45,
  phone1: 51,
  phone2: 52,
  education: 55,
  work: 61,
  servedBefore2022: 68,
  arrivedFromUnit: 71,
  conscription: 74,
  ubdCertificate: 80,
  familyStatus: 87,
  mother: 97,
  father: 107,
  trustedPerson: 117,
  children: 127,
  sports: 134,
  sportsRank: 137,
  signatureDate: 142,
};

const CHECKBOX_PARAGRAPHS: Partial<Record<PreAnketaFieldKey, number>> = {
  serviceContract: 64,
  serviceMobilized: 65,
  criminalCharged: 83,
  criminalNotCharged: 84,
};

const escapeXml = (value: string) =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");

const decodeXml = (value: string) =>
  value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");

const paragraphText = (paragraph: string) => {
  const texts: string[] = [];
  paragraph.replace(/<w:t\b[^>]*>([^<]*)<\/w:t>/g, (_, text: string) => {
    texts.push(decodeXml(text));
    return _;
  });
  return texts.join("");
};

const setParagraphText = (paragraph: string, nextText: string) => {
  const escaped = escapeXml(nextText);
  if (!/<w:t\b/.test(paragraph)) {
    return paragraph.replace(
      /<\/w:p>/,
      `<w:r><w:rPr><w:rFonts w:cstheme="minorHAnsi"/></w:rPr><w:t xml:space="preserve">${escaped}</w:t></w:r></w:p>`,
    );
  }
  let used = false;
  return paragraph.replace(/<w:t\b([^>]*)>[^<]*<\/w:t>/g, (run, attrs: string) => {
    if (used) return run.replace(/>[^<]*</, "><");
    used = true;
    const nextAttrs = /xml:space=/.test(attrs)
      ? attrs.replace(/xml:space="[^"]*"/, 'xml:space="preserve"')
      : `${attrs} xml:space="preserve"`;
    return `<w:t${nextAttrs}>${escaped}</w:t>`;
  });
};

const splitParagraphs = (documentXml: string) =>
  [...documentXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map((match) => match[0]);

const joinParagraphs = (documentXml: string, paragraphs: string[]) => {
  let index = 0;
  return documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, () => paragraphs[index++] ?? "");
};

const formatPhoneParagraph = (current: string, value: string) => {
  const trimmed = value.trim();
  if (!trimmed) return current;
  if (/^\+380\s*$/.test(current.trim())) {
    return trimmed.startsWith("+380")
      ? trimmed
      : `+380 ${trimmed.replace(/^\+?380?/, "")}`;
  }
  return trimmed;
};

const formatCheckboxParagraph = (current: string, checked: boolean) => {
  if (!checked) return current.replace(/☑/g, "☐");
  return current.includes("☑")
    ? current
    : current.replace(/☐/g, "☑");
};

const formatTodayForSignature = () => {
  const now = new Date();
  const day = String(now.getDate()).padStart(2, "0");
  const month = String(now.getMonth() + 1).padStart(2, "0");
  return `${day}.${month}.${now.getFullYear()}`;
};

export const fillPreAnketaDocumentXml = (
  documentXml: string,
  fields: PreAnketaFormFields,
) => {
  const paragraphs = splitParagraphs(documentXml);

  for (const [fieldKey, paragraphIndex] of Object.entries(PARAGRAPH_TARGETS)) {
    const key = fieldKey as PreAnketaFieldKey;
    const value = String(fields[key] ?? "").trim();
    if (!value || paragraphIndex >= paragraphs.length) continue;
    const current = paragraphText(paragraphs[paragraphIndex] ?? "");
    const next =
      key === "phone1" || key === "phone2"
        ? formatPhoneParagraph(current, value)
        : value;
    paragraphs[paragraphIndex] = setParagraphText(
      paragraphs[paragraphIndex] ?? "",
      next,
    );
  }

  for (const [fieldKey, paragraphIndex] of Object.entries(CHECKBOX_PARAGRAPHS)) {
    const key = fieldKey as PreAnketaFieldKey;
    if (paragraphIndex >= paragraphs.length) continue;
    const checked = Boolean(String(fields[key] ?? "").trim());
    paragraphs[paragraphIndex] = setParagraphText(
      paragraphs[paragraphIndex] ?? "",
      formatCheckboxParagraph(
        paragraphText(paragraphs[paragraphIndex] ?? ""),
        checked,
      ),
    );
  }

  const signatureDate =
    fields.signatureDate.trim() || formatTodayForSignature();
  if (paragraphs[142]) {
    paragraphs[142] = setParagraphText(paragraphs[142], signatureDate);
  }
  if (paragraphs[158]) {
    const current = paragraphText(paragraphs[158]);
    paragraphs[158] = setParagraphText(
      paragraphs[158],
      current.replace(/2026/, signatureDate.slice(-4) || "2026"),
    );
  }

  const rosterLine = [
    fields.rank.trim(),
    fields.fullName.trim(),
    fields.birthDate.trim(),
  ]
    .filter(Boolean)
    .join(", ");
  if (rosterLine && paragraphs[150]) {
    paragraphs[150] = setParagraphText(paragraphs[150], rosterLine);
  }

  return joinParagraphs(documentXml, paragraphs);
};

export const createPreAnketaWordBlob = async (fields: PreAnketaFormFields) => {
  const response = await fetch(PRE_ANKETA_TEMPLATE_URL);
  if (!response.ok) {
    throw new Error("Не знайшов шаблон анкети.");
  }
  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) {
    throw new Error("Шаблон анкети пошкоджений.");
  }

  const nextXml = fillPreAnketaDocumentXml(documentXml, fields);
  zip.file("word/document.xml", nextXml);
  await stripRedColorInWordZip(zip);
  return zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
};
