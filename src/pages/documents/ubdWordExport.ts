import JSZip from "jszip";
import { loadUbdTemplate } from "./ubdTemplateStore";
import { stripRedColorInWordZip } from "./wordXml";
import { capitalizeReportPosition } from "./reportPosition";

type WordSignatory = {
  blockType: "SIGNER" | "APPROVAL";
  title: string;
  rank: string;
  fullName: string;
  signatureData?: string | null;
};

export type UbdWordFields = {
  commander: string;
  fullName: string;
  rank: string;
  staffPosition: string;
  birthDate: string;
  rnokpp: string;
  taskPeriod: string;
  taskPlace: string;
  basis: string;
  signatories: WordSignatory[];
};

export const UBD_TEMPLATE_FIELDS = [
  { key: "COMMANDER", label: "Адресат (Командиру …)" },
  { key: "RANK", label: "Військове звання" },
  { key: "FULL_NAME_1", label: "Прізвище" },
  { key: "FULL_NAME_2", label: "Ім’я та по батькові" },
  { key: "POSITION", label: "Посада згідно штату" },
  { key: "BIRTH_DATE", label: "Дата народження" },
  { key: "RNOKPP", label: "РНОКПП" },
  { key: "TASK_PERIOD", label: "Період виконання завдань" },
  { key: "TASK_PLACE", label: "Місце виконання завдань" },
  { key: "BASIS", label: "Підстава" },
  { key: "SIGNER_TITLE_1", label: "Підписант, рядок 1" },
  { key: "SIGNER_TITLE_2", label: "Підписант, рядок 2" },
  { key: "SIGNER_TITLE_3", label: "Підписант, рядок 3" },
  { key: "SIGNER_RANK", label: "Звання підписанта" },
  { key: "SIGNER_NAME", label: "ПІБ підписанта" },
  { key: "APPROVER_TITLE_1", label: "Затверджую, рядок 1" },
  { key: "APPROVER_TITLE_2", label: "Затверджую, рядок 2" },
  { key: "APPROVER_RANK", label: "Звання затверджуючого" },
  { key: "APPROVER_NAME", label: "ПІБ затверджуючого" },
] as const;

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

const isDateLine = (value: string) => /^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value);

const personNameLines = (fullName: string) => {
  const parts = fullName
    .trim()
    .replace(/\s*\([^)]+\)\s*$/, "")
    .split(/\s+/)
    .filter(Boolean);
  if (!parts.length) return ["", ""];
  if (parts.length === 1) return [parts[0].toLocaleUpperCase("uk-UA"), ""];
  return [
    parts[0].toLocaleUpperCase("uk-UA"),
    parts
      .slice(1)
      .map((part) => part.charAt(0).toLocaleUpperCase("uk-UA") + part.slice(1))
      .join(" "),
  ];
};

const splitSignatory = (
  signatory: WordSignatory | null,
  lineCount: number,
  options?: { harvestRankFromTitle?: boolean },
) => {
  if (!signatory) {
    return {
      titleLines: Array.from({ length: lineCount }, () => ""),
      allTitleLines: [] as string[],
      rank: "",
      date: "",
      fullName: "",
    };
  }
  const harvestRankFromTitle = options?.harvestRankFromTitle !== false;
  const rawLines = signatory.title
    .replace(/^ЗАТВЕРДЖУЮ[\s:–—-]*/iu, "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const date = rawLines.find(isDateLine) ?? "";
  const rankFromTitle = harvestRankFromTitle
    ? [...rawLines]
        .reverse()
        .find(
          (line) =>
            !isDateLine(line) &&
            !/командир|тимчасово|військової частини|затверджую|призначено/i.test(
              line,
            ),
        )
    : "";
  const rank = signatory.rank.trim() || rankFromTitle || "";
  const titleLines = rawLines.filter((line) => line !== date && line !== rank);
  const padded = [...titleLines];
  while (padded.length < lineCount) padded.push("");
  if (padded.length > lineCount) {
    padded.splice(
      lineCount - 1,
      padded.length,
      padded.slice(lineCount - 1).join(" "),
    );
  }
  return {
    titleLines: padded.slice(0, lineCount),
    allTitleLines: titleLines,
    rank,
    date,
    fullName: formatSignatoryName(signatory.fullName),
  };
};

const escapeXml = (value: string) =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");

const RUN_FONTS =
  '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>';

const textParagraph = (text: string, align: "left" | "right" | "both" = "left") => {
  const safe = escapeXml(text);
  const space = text !== text.trim() || text.startsWith(" ") || text.endsWith(" ")
    ? ` xml:space="preserve"`
    : "";
  return `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="${align}"/><w:rPr>${RUN_FONTS}<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr><w:r><w:rPr>${RUN_FONTS}<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t${space}>${safe}</w:t></w:r></w:p>`;
};

/** «ЗАТВЕРДЖУЮ»: текст посади зліва, звання + ПІБ справа. */
const buildApprovalTwoColumnTable = ({
  titleLines,
  rank,
  fullName,
}: {
  titleLines: string[];
  rank: string;
  fullName: string;
}) => {
  const leftLines = ["ЗАТВЕРДЖУЮ", ...titleLines.map((line) => line.trim())].filter(
    Boolean,
  );
  const rightLines = [rank.trim(), fullName.trim()].filter(Boolean);
  const leftXml = leftLines.map((line) => textParagraph(line, "left")).join("");
  const rightXml = rightLines.length
    ? rightLines
        .map((line, index) =>
          textParagraph(
            line,
            index === rightLines.length - 1 ? "right" : "right",
          ),
        )
        .join("")
    : textParagraph("");

  return `<w:tbl><w:tblPr><w:tblW w:w="9640" w:type="dxa"/><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="5200"/><w:gridCol w:w="4440"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5200" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>${leftXml || textParagraph("")}</w:tc><w:tc><w:tcPr><w:tcW w:w="4440" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr>${rightXml}</w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;
};

const paragraphPlainText = (paragraphXml: string) =>
  [...paragraphXml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
    .map((match) => match[1])
    .join("")
    .replaceAll("&amp;", "&")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">");

/** Замінює блок ЗАТВЕРДЖУЮ + title + rank/name на таблицю 2 колонки. */
const replaceApprovalBlockWithTable = (
  xml: string,
  tableXml: string,
) => {
  const paragraphs = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
  let start = -1;
  let end = -1;
  for (const match of paragraphs) {
    const text = paragraphPlainText(match[0]);
    const index = match.index ?? -1;
    if (index < 0) continue;
    if (start < 0 && /ЗАТВЕРДЖУЮ/.test(text)) {
      start = index;
      end = index + match[0].length;
      continue;
    }
    if (start >= 0) {
      if (
        text.includes("{{APPROVER_TITLE_1}}") ||
        text.includes("{{APPROVER_TITLE_2}}") ||
        text.includes("{{APPROVER_RANK}}") ||
        text.includes("{{APPROVER_NAME}}")
      ) {
        end = index + match[0].length;
        continue;
      }
      // Stop once we left the approval placeholders block.
      if (
        text.includes("{{SIGNER_") ||
        text.includes("ВІДКРИТА") ||
        text.includes("Додаток")
      ) {
        break;
      }
      // Keep consuming short empty / leftover approval lines right after start.
      if (!text.trim()) {
        end = index + match[0].length;
        continue;
      }
      break;
    }
  }
  if (start < 0 || end < 0) return xml;
  return `${xml.slice(0, start)}${tableXml}${xml.slice(end)}`;
};

const applyRightNameTab = (
  xml: string,
  rankToken: string,
  nameToken: string,
  tabPos: number,
) => {
  const textIndex = xml.indexOf(rankToken);
  if (textIndex < 0) return xml;
  const paragraphStart = xml.lastIndexOf("<w:p ", textIndex);
  const paragraphEnd = xml.indexOf("</w:p>", textIndex);
  if (paragraphStart < 0 || paragraphEnd < 0) return xml;

  const paragraph = xml.slice(paragraphStart, paragraphEnd + 6);
  const pPrClose = paragraph.indexOf("</w:pPr>");
  if (pPrClose < 0) return xml;

  let pPr = paragraph.slice(0, pPrClose + "</w:pPr>".length);
  const rightTabs = `<w:tabs><w:tab w:val="right" w:pos="${tabPos}"/></w:tabs>`;
  pPr = pPr.includes("<w:tabs>")
    ? pPr.replace(/<w:tabs>[\s\S]*?<\/w:tabs>/, rightTabs)
    : pPr.replace("<w:pPr>", `<w:pPr>${rightTabs}`);
  pPr = pPr.replace(/<w:jc w:val="both"\s*\/>/, `<w:jc w:val="left"/>`);

  const rPrMatch = paragraph.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  const rPr =
    rPrMatch?.[0] ??
    `<w:rPr>${RUN_FONTS}<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>`;

  const rebuilt = `${pPr}<w:r>${rPr}<w:t>${rankToken}</w:t></w:r><w:r>${rPr}<w:tab/><w:t>${nameToken}</w:t></w:r></w:p>`;
  return `${xml.slice(0, paragraphStart)}${rebuilt}${xml.slice(paragraphEnd + 6)}`;
};

const fillPlaceholders = (xml: string, values: Record<string, string>) => {
  let next = xml;
  for (const [key, value] of Object.entries(values)) {
    const token = `{{${key}}}`;
    const safe = escapeXml(value);
    if (next.includes(token)) {
      next = next.replaceAll(token, safe);
      continue;
    }
    const pattern = token
      .split("")
      .map((char) => char.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
      .join("(?:<[^>]+>)*");
    next = next.replace(new RegExp(pattern, "g"), safe);
  }
  return next;
};

const ZIP_FILE_OPTIONS = { createFolders: false as const };

const writeZipFile = (zip: JSZip, path: string, data: string | Uint8Array) => {
  zip.file(path, data, ZIP_FILE_OPTIONS);
};

const generateDocxBlob = (zip: JSZip) => {
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

const dataUrlParts = (dataUrl: string) => {
  const [header, base64 = ""] = dataUrl.split(",", 2);
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  const isPng = header.includes("image/png");
  return { bytes, isPng };
};

const replaceSignature = async (zip: JSZip, signatureData: string) => {
  const { bytes, isPng } = dataUrlParts(signatureData);
  const relsPath = "word/_rels/document.xml.rels";
  const typesPath = "[Content_Types].xml";
  const rels = await zip.file(relsPath)?.async("string");
  const types = await zip.file(typesPath)?.async("string");
  if (!rels || !types) return;

  const jpegName = "word/media/image1.jpeg";
  const pngName = "word/media/image1.png";
  if (isPng) {
    zip.remove(jpegName);
    writeZipFile(zip, pngName, bytes);
    writeZipFile(
      zip,
      relsPath,
      rels.replace("media/image1.jpeg", "media/image1.png"),
    );
    writeZipFile(
      zip,
      typesPath,
      types.includes('Extension="png"')
        ? types
        : types.replace(
            '<Default Extension="jpeg"',
            '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg"',
          ),
    );
    return;
  }
  writeZipFile(zip, jpegName, bytes);
};

const valuesFromFields = (fields: UbdWordFields) => {
  const signer =
    fields.signatories.find((item) => item.blockType === "SIGNER") ?? null;
  const approval =
    fields.signatories.find((item) => item.blockType === "APPROVAL") ?? null;
  const signerParts = splitSignatory(signer, 3);
  const approvalParts = splitSignatory(approval, 8, {
    harvestRankFromTitle: false,
  });
  const [surname, givenNames] = personNameLines(fields.fullName);

  return {
    COMMANDER: fields.commander.trim(),
    RANK: fields.rank.trim(),
    FULL_NAME_1: surname,
    FULL_NAME_2: givenNames,
    POSITION: capitalizeReportPosition(fields.staffPosition),
    BIRTH_DATE: fields.birthDate.trim(),
    RNOKPP: fields.rnokpp.trim(),
    TASK_PERIOD: fields.taskPeriod.trim(),
    TASK_PLACE: fields.taskPlace.trim(),
    BASIS: fields.basis.trim(),
    SIGNER_TITLE_1: signerParts.titleLines[0] ?? "",
    SIGNER_TITLE_2: signerParts.titleLines[1] ?? "",
    SIGNER_TITLE_3: signerParts.titleLines[2] ?? "",
    SIGNER_RANK: signerParts.rank,
    SIGNER_NAME: signerParts.fullName,
    SIGNER_DATE: signerParts.date,
    APPROVER_TITLE_1: approvalParts.titleLines[0] ?? "",
    APPROVER_TITLE_2: approvalParts.titleLines[1] ?? "",
    APPROVER_RANK: approvalParts.rank,
    APPROVER_NAME: approvalParts.fullName,
    approvalTitleLines: approvalParts.allTitleLines,
    approvalRank: approvalParts.rank,
    approvalName: approvalParts.fullName,
  };
};

export const createUbdWordBlob = async (fields: UbdWordFields) => {
  const template = await loadUbdTemplate();
  const zip = await JSZip.loadAsync(template);
  const documentPath = "word/document.xml";
  const xml = await zip.file(documentPath)?.async("string");
  if (!xml) {
    throw new Error("У шаблоні немає word/document.xml.");
  }
  const values = valuesFromFields(fields);
  const {
    SIGNER_DATE,
    approvalTitleLines,
    approvalRank,
    approvalName,
    APPROVER_TITLE_1: _a1,
    APPROVER_TITLE_2: _a2,
    APPROVER_RANK: _ar,
    APPROVER_NAME: _an,
    ...tokens
  } = values;

  const approvalTable = buildApprovalTwoColumnTable({
    titleLines: approvalTitleLines,
    rank: approvalRank,
    fullName: approvalName,
  });
  let filled = replaceApprovalBlockWithTable(xml, approvalTable);
  filled = applyRightNameTab(
    filled,
    "{{SIGNER_RANK}}",
    "{{SIGNER_NAME}}",
    9640,
  );
  filled = fillPlaceholders(filled, tokens);
  if (SIGNER_DATE) {
    const safeDate = escapeXml(SIGNER_DATE);
    filled = filled.replaceAll("{{SIGNER_DATE}}", safeDate);
    filled = filled.replaceAll("14.07.2026", safeDate);
  }
  writeZipFile(zip, documentPath, filled);

  const signer = fields.signatories.find((item) => item.blockType === "SIGNER");
  if (signer?.signatureData) {
    await replaceSignature(zip, signer.signatureData);
  }

  await stripRedColorInWordZip(zip);
  return generateDocxBlob(zip);
};

/** Підставляє блок «ЗАТВЕРДЖУЮ» як таблицю: текст зліва, звання/ПІБ справа. */
export const injectUbdApprovalTwoColumnBlock = (
  xml: string,
  approval: WordSignatory | null,
) => {
  const parts = splitSignatory(approval, 8, { harvestRankFromTitle: false });
  return replaceApprovalBlockWithTable(
    xml,
    buildApprovalTwoColumnTable({
      titleLines: parts.allTitleLines,
      rank: parts.rank,
      fullName: parts.fullName,
    }),
  );
};
