import JSZip from "jszip";
import {
  form12FighterSignName,
  form12PleaText,
  splitForm12Signatory,
  type Form12ReportFields,
} from "./form12Report";
import { stripRedColorInWordZip } from "./wordXml";

const FORM12_TEMPLATE_URL = `${import.meta.env.BASE_URL}templates/form12-report.docx`;

const SAMPLE = {
  plea: "Прошу Вашого клопотання перед вищим командуванням на надання мені, солдату ТРИГУБУ Сергію Олександровичу, оператору безпілотних літальних апаратів 3 піхотної роти 1 піхотного батальйону, довідки Ф-12 (довідки про безпосередню участь у бойових діях).",
  position:
    "Оператор безпілотних літальних апаратів 3 піхотної роти 1 піхотного батальйонувійськової частини А4862",
  rank: "солдат",
  fighterDate: "14.08.26",
  fighterName: "Сергій ТРИГУБ",
  signerName: "Андрій КІЯНЕНКО",
  signerDate: "14.08.2026",
  approvalName: "Олег АДАМОВ",
  actingTitle: "Тимчасово виконуючий обов’язки",
  battalionLine: "командира 1 піхотного батальйону",
  unitLine: "військової частини А4862",
  approvalUnit: "командира військової частини А4862",
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

const normalizeText = (value: string) =>
  value
    .replace(/[’']/g, "'")
    .replace(/\s+/g, " ")
    .trim();

const withFallback = (value: string, fallback = "") =>
  value.trim() || fallback;

const joinNameForWord = (value: string) =>
  value.trim().replace(/\s+/g, "\u00A0");

const formatSignatoryName = (fullName: string) => {
  const parts = fullName.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "";
  if (parts.length === 1) return parts[0].toLocaleUpperCase("uk-UA");
  const surname = parts[parts.length - 1].toLocaleUpperCase("uk-UA");
  const given = parts
    .slice(0, -1)
    .map(
      (part) =>
        part.charAt(0).toLocaleUpperCase("uk-UA") +
        part.slice(1).toLocaleLowerCase("uk-UA"),
    )
    .join("\u00A0");
  return `${given}\u00A0${surname}`;
};

const shortDate = (value: string) => {
  const match = value.trim().match(/^(\d{1,2}\.\d{1,2}\.)(\d{2})(\d{2})$/);
  return match ? `${match[1]}${match[3]}` : value.trim();
};

const positionWithUnit = (position: string) => {
  const text = position.trim() || "______";
  if (/військової частини/i.test(text)) return text;
  return `${text} військової частини А4862`;
};

const writeTextNode = (attrs: string, text: string) => {
  const needsPreserve =
    text.startsWith(" ") ||
    text.endsWith(" ") ||
    /\s{2}/.test(text) ||
    text.length === 0;
  let nextAttrs = attrs;
  if (needsPreserve) {
    if (/xml:space=/.test(nextAttrs)) {
      nextAttrs = nextAttrs.replace(
        /xml:space="[^"]*"/,
        'xml:space="preserve"',
      );
    } else {
      nextAttrs = `${nextAttrs} xml:space="preserve"`;
    }
  }
  return `<w:t${nextAttrs}>${escapeXml(text)}</w:t>`;
};

const listParagraphTexts = (paragraph: string) => {
  const texts: string[] = [];
  paragraph.replace(/<w:t\b([^>]*)>([^<]*)<\/w:t>/g, (_, _attrs, text: string) => {
    texts.push(decodeXml(text));
    return _;
  });
  return texts;
};

const applyParagraphTexts = (paragraph: string, texts: string[]) => {
  let index = 0;
  return paragraph.replace(
    /<w:t\b([^>]*)>([^<]*)<\/w:t>/g,
    (full, attrs: string) => {
      const next = texts[index];
      index += 1;
      if (next === undefined) return full;
      return writeTextNode(attrs, next);
    },
  );
};

const firstContentIndex = (texts: string[]) =>
  texts.findIndex((text) => text.trim().length > 0);

const setLeadingContent = (paragraph: string, value: string) => {
  const texts = listParagraphTexts(paragraph);
  const start = firstContentIndex(texts);
  if (start < 0) return paragraph;
  texts[start] = value;
  for (let index = start + 1; index < texts.length; index += 1) {
    texts[index] = "";
  }
  return applyParagraphTexts(paragraph, texts);
};

const NAME_TAB_POS = "6379";
const NAME_RUN_PR =
  "<w:rPr><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/></w:rPr>";

const paragraphOpenTag = (paragraph: string) =>
  paragraph.match(/<w:p\b[^>]*>/)?.[0] ?? "<w:p>";

const paragraphPPr = (paragraph: string) =>
  paragraph.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? "<w:pPr></w:pPr>";

const paragraphDrawingRuns = (paragraph: string) =>
  [
    ...paragraph.matchAll(
      /<w:r\b[^>]*>[\s\S]*?<w:drawing>[\s\S]*?<\/w:drawing>[\s\S]*?<\/w:r>/g,
    ),
  ]
    .map((match) => match[0])
    .join("");

const removeDrawingByRelId = (documentXml: string, relId: string) =>
  documentXml.replace(
    new RegExp(
      `<w:r\\b[^>]*>(?:(?!<\\/w:r>)[\\s\\S])*?<w:drawing>(?:(?!<\\/w:drawing>)[\\s\\S])*?r:embed="${relId}"(?:(?!<\\/w:drawing>)[\\s\\S])*?<\\/w:drawing>(?:(?!<\\/w:r>)[\\s\\S])*?<\\/w:r>`,
      "g",
    ),
    "",
  );

const withNameTab = (pPr: string) => {
  const cleaned = pPr
    .replace(/<w:jc\b[^/]*\/>/g, "")
    .replace(/<w:tabs>[\s\S]*?<\/w:tabs>/g, "");
  return cleaned.replace(
    "<w:pPr>",
    `<w:pPr><w:tabs><w:tab w:val="left" w:pos="${NAME_TAB_POS}"/></w:tabs>`,
  );
};

const setNameColumn = (paragraph: string, leftText: string, name: string) => {
  const left = leftText.trim();
  const leftRun = left
    ? `<w:r>${NAME_RUN_PR}<w:t xml:space="preserve">${escapeXml(left)}</w:t></w:r>`
    : "";
  const tabRun = `<w:r>${NAME_RUN_PR}<w:tab/></w:r>`;
  const nameRun = `<w:r>${NAME_RUN_PR}<w:t xml:space="preserve">${escapeXml(joinNameForWord(name))}</w:t></w:r>`;
  return `${paragraphOpenTag(paragraph)}${withNameTab(paragraphPPr(paragraph))}${paragraphDrawingRuns(paragraph)}${leftRun}${tabRun}${nameRun}</w:p>`;
};

const setFighterDate = (paragraph: string, value: string) => {
  const texts = listParagraphTexts(paragraph);
  if (!texts.length) return paragraph;
  const current = texts[0];
  const date = shortDate(value) || SAMPLE.fighterDate;
  const suffix = current.startsWith(SAMPLE.fighterDate)
    ? current.slice(SAMPLE.fighterDate.length)
    : current.replace(/^\d{1,2}\.\d{1,2}\.\d{2,4}/, "");
  texts[0] = `${date}${suffix}`;
  return applyParagraphTexts(paragraph, texts);
};

const setPlainDate = (paragraph: string, value: string) => {
  const texts = listParagraphTexts(paragraph);
  const start = firstContentIndex(texts);
  if (start < 0) return paragraph;
  texts[start] = withFallback(value, texts[start]);
  for (let index = start + 1; index < texts.length; index += 1) {
    texts[index] = "";
  }
  return applyParagraphTexts(paragraph, texts);
};

const ZIP_FILE_OPTIONS = { createFolders: false as const };

const writeZipFile = (zip: JSZip, path: string, data: string | Uint8Array) => {
  zip.file(path, data, ZIP_FILE_OPTIONS);
};

const dataUrlParts = (dataUrl: string) => {
  const [header, base64 = ""] = dataUrl.split(",", 2);
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return { bytes, isPng: header.includes("image/png") };
};

const replaceRelImage = async (
  zip: JSZip,
  relId: string,
  signatureData: string,
  fileName: string,
) => {
  const { bytes, isPng } = dataUrlParts(signatureData);
  const relsPath = "word/_rels/document.xml.rels";
  const typesPath = "[Content_Types].xml";
  const rels = await zip.file(relsPath)?.async("string");
  const types = await zip.file(typesPath)?.async("string");
  if (!rels || !types) return;

  const ext = isPng ? "png" : "jpeg";
  const path = `word/media/${fileName}.${ext}`;
  writeZipFile(zip, path, bytes);
  const nextRels = rels.replace(
    new RegExp(`(<Relationship Id="${relId}"[^>]*Target=")media/[^"]+(")`),
    `$1media/${fileName}.${ext}$2`,
  );
  writeZipFile(zip, relsPath, nextRels);
  if (isPng && !types.includes('Extension="png"')) {
    writeZipFile(
      zip,
      typesPath,
      types.replace(
        '<Default Extension="jpeg"',
        '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg"',
      ),
    );
  }
};

export const createForm12WordBlob = async (fields: Form12ReportFields) => {
  const response = await fetch(FORM12_TEMPLATE_URL);
  if (!response.ok) {
    throw new Error("Не знайшов шаблон Форми 12.");
  }
  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) {
    throw new Error("Шаблон Форми 12 пошкоджений.");
  }

  const signer =
    fields.signatories.find((item) => item.blockType === "SIGNER") ?? null;
  const approval =
    fields.signatories.find((item) => item.blockType === "APPROVAL") ?? null;
  const signerParts = splitForm12Signatory(signer, 3);
  const approvalParts = splitForm12Signatory(approval, 2);
  const signerName = formatSignatoryName(signerParts.fullName);
  const approvalName = formatSignatoryName(approvalParts.fullName);
  const fighterName = form12FighterSignName(fields.fullName);
  const fighterRank = withFallback(fields.rank, SAMPLE.rank).toLocaleLowerCase(
    "uk-UA",
  );

  let actingHits = 0;
  let battalionHits = 0;

  let filled = documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const texts = listParagraphTexts(paragraph);
    if (!texts.length) return paragraph;
    const joined = texts.join("");
    const trimmed = normalizeText(joined);

    if (trimmed === normalizeText(SAMPLE.plea)) {
      return setLeadingContent(paragraph, form12PleaText(fields));
    }
    if (trimmed === normalizeText(SAMPLE.position)) {
      return setLeadingContent(paragraph, positionWithUnit(fields.staffPosition));
    }
    if (joined.includes(SAMPLE.signerName)) {
      return setNameColumn(
        paragraph,
        signerParts.rank,
        withFallback(signerName, SAMPLE.signerName),
      );
    }
    if (joined.includes(SAMPLE.approvalName)) {
      return setNameColumn(
        paragraph,
        approvalParts.rank,
        withFallback(approvalName, SAMPLE.approvalName),
      );
    }
    if (
      trimmed === SAMPLE.fighterName ||
      trimmed.endsWith(SAMPLE.fighterName)
    ) {
      return setNameColumn(
        paragraph,
        "",
        withFallback(fighterName, SAMPLE.fighterName),
      );
    }
    if (trimmed.startsWith(SAMPLE.fighterDate)) {
      return setFighterDate(paragraph, withFallback(fields.date));
    }
    if (trimmed === SAMPLE.signerDate) {
      return signerParts.date
        ? setPlainDate(paragraph, signerParts.date)
        : paragraph;
    }
    if (trimmed === SAMPLE.rank) {
      return setLeadingContent(paragraph, `${fighterRank}`);
    }
    if (trimmed === normalizeText(SAMPLE.actingTitle)) {
      actingHits += 1;
      const line =
        actingHits === 3
          ? approvalParts.titleLines[0]
          : signerParts.titleLines[0];
      return line.trim() ? setLeadingContent(paragraph, line) : paragraph;
    }
    if (trimmed === normalizeText(SAMPLE.battalionLine)) {
      battalionHits += 1;
      const line = signerParts.titleLines[1];
      const next = line.trim();
      if (!next) return paragraph;
      return setLeadingContent(
        paragraph,
        joined.endsWith(" ") ? `${next} ` : next,
      );
    }
    if (trimmed === normalizeText(SAMPLE.unitLine)) {
      const line = signerParts.titleLines[2];
      return line.trim() ? setLeadingContent(paragraph, line) : paragraph;
    }
    if (trimmed === normalizeText(SAMPLE.approvalUnit)) {
      const line = approvalParts.titleLines[1];
      return line.trim() ? setLeadingContent(paragraph, line) : paragraph;
    }
    return paragraph;
  });

  if (fields.signatureData) {
    await replaceRelImage(zip, "rId7", fields.signatureData, "fighter-sign");
  } else {
    filled = removeDrawingByRelId(filled, "rId7");
  }
  if (signer?.signatureData) {
    await replaceRelImage(zip, "rId8", signer.signatureData, "commander-sign");
  } else {
    filled = removeDrawingByRelId(filled, "rId8");
  }

  writeZipFile(zip, "word/document.xml", filled);
  await stripRedColorInWordZip(zip);
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) zip.remove(name);
  }
  return zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
};
