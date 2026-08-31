import JSZip from "jszip";
import type { ZhbdCertificateFields } from "./zhbdCertificateReport";
import { formatGivenSurname } from "./ubdRestoreReport";
import { stripRedColorInWordZip } from "./wordXml";

const TEMPLATE_URL = `${import.meta.env.BASE_URL}templates/zhbd-certificate.docx`;

const SAMPLE = {
  body:
    "Надана молодшому сержанту АЛЄКСЄЄВ Дмитро Юрійович, в тому що виконував (виконує) бойові (спеціальні) завдання на посаді командира 2 штурмового відділення 1 штурмового взводу 3 штурмової роти 2 штурмового батальйону військової частини А4862 в період з 01.01.2026 по теперішній час (з 01.01.2026 по 30.04.2026).",
  basisOrders:
    "Бойові розпорядження командира військової частини А4862 від 01.01.2026. № 4862/6/01/001/дск, від 01.04.2026 № 4862/829/04/089/дск.",
  basisJournal:
    "Журнал ведення бойових дій військової частини А4862 (інв. № 171/дск. від 01.01.2026 року)",
  signerTitle1: "Начальник групи персоналу ",
  signerTitle2: "2 штурмового батальйону ",
  signerTitle3: "військової частини А4862",
  signerLine: "молодший лейтенантОлександр РАКІН",
  signerDate: "23.08.2026",
  headerDate: "“__” ________ 2026р.",
  documentNumber: "№ ______",
};

const ZIP_FILE_OPTIONS = { createFolders: false as const };

/**
 * Floating signature (Word «Перед текстом»).
 * Height ≈ 26.4mm; width follows image aspect (up to ~59mm).
 */
const SIGNATURE_LAYOUT = {
  relId: "rId8",
  targetCy: 951_269,
  maxCx: 2_137_907,
} as const;

const APPROVAL_SIGNATURE_LAYOUT = {
  relId: "rId9",
  targetCy: 951_269,
  maxCx: 2_137_907,
} as const;

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

const withFallback = (value: string, fallback = "______") =>
  value.trim() || fallback;

const isDateLine = (value: string) => /^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value);

const RANK_TAIL =
  /(головний майстер-сержант|старший майстер-сержант|майстер-сержант|штаб-сержант|головний сержант|старший сержант|молодший сержант|старший лейтенант|молодший лейтенант|старший солдат|підполковник|полковник|лейтенант|сержант|капітан|майор|солдат|рекрут)$/iu;

const splitSigner = (signatory: {
  title: string;
  rank: string;
  fullName: string;
} | null) => {
  if (!signatory) {
    return {
      titleLines: ["", "", ""] as [string, string, string],
      rank: "",
      date: "",
    };
  }
  const lines = signatory.title
    .split(/\r?\n/)
    .map((line) => line.trim())
    .flatMap((line) => {
      const match = line.match(/^(.*?)(?:\s+)(\d{1,2}\.\d{1,2}\.\d{4})$/);
      if (match?.[1]?.trim()) return [match[1].trim(), match[2]];
      return [line];
    })
    .filter(Boolean);
  const date = lines.find(isDateLine) ?? "";
  let rank = signatory.rank.trim();
  const titles: string[] = [];
  for (const line of lines) {
    if (isDateLine(line)) continue;
    const tail = line.match(RANK_TAIL);
    if (tail && !rank) {
      rank = tail[1];
      const head = line.slice(0, -tail[1].length).trim();
      if (head) titles.push(head);
      continue;
    }
    if (
      rank &&
      line.toLocaleLowerCase("uk-UA") === rank.toLocaleLowerCase("uk-UA")
    ) {
      continue;
    }
    titles.push(line);
  }
  while (titles.length < 3) titles.push("");
  if (titles.length > 3) {
    titles.splice(2, titles.length, titles.slice(2).join(" "));
  }
  return {
    titleLines: [titles[0], titles[1], titles[2]] as [string, string, string],
    rank,
    date,
  };
};

const paragraphJoinedText = (paragraph: string) => {
  const texts: string[] = [];
  paragraph.replace(/<w:t\b[^>]*>([^<]*)<\/w:t>/g, (_, text: string) => {
    texts.push(decodeXml(text));
    return _;
  });
  return texts.join("");
};

const replaceParagraphTexts = (
  xml: string,
  replace: (text: string) => string,
) =>
  xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const joined = paragraphJoinedText(paragraph);
    if (!joined.length) return paragraph;
    const next = replace(joined);
    if (next === joined) return paragraph;
    const escaped = escapeXml(next);
    let used = false;
    return paragraph.replace(/<w:t\b[^>]*>[^<]*<\/w:t>/g, (run) => {
      if (used) return run.replace(/>[^<]*</, "><");
      used = true;
      return run.replace(/>[^<]*</, `>${escaped}<`);
    });
  });

const RUN_FONTS =
  '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';

const textParagraph = (text: string, align: "left" | "right" = "left") => {
  const body = text.trim()
    ? `<w:r><w:rPr>${RUN_FONTS}<w:noProof/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`
    : `<w:r><w:rPr>${RUN_FONTS}<w:noProof/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:r>`;
  return `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="${align}"/><w:rPr>${RUN_FONTS}<w:noProof/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr>${body}</w:p>`;
};

const frontOfTextSignatureDrawing = (
  relId: string,
  cx: number,
  cy: number,
  docPrId: number,
) =>
  `<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251662336" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${docPrId}" name="Підпис"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="0"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="Підпис"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

const signatureImageParagraph = (drawingRun: string) =>
  `<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:rPr>${RUN_FONTS}<w:noProof/><w:sz w:val="28"/></w:rPr></w:pPr>${drawingRun}</w:p>`;

const emptyParagraph = () =>
  `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>${RUN_FONTS}<w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:rPr>${RUN_FONTS}<w:sz w:val="28"/></w:rPr></w:r></w:p>`;

/** «ЗАТВЕРДЖУЮ» — 2 columns to the edges (left titles, right name). */
const buildApprovalBlock = ({
  titleLines,
  rank,
  fullName,
  signatureRun,
}: {
  titleLines: [string, string, string];
  rank: string;
  fullName: string;
  signatureRun: string;
}) => {
  const displayName = (
    formatGivenSurname(fullName) || withFallback(fullName)
  ).replace(/\s+/g, "\u00A0");
  const leftLines = [
    "ЗАТВЕРДЖУЮ",
    ...titleLines.map((line) => line.trim()).filter(Boolean),
    rank.trim(),
  ].filter(Boolean);
  const leftXml = leftLines.map((line) => textParagraph(line, "left")).join("");
  const rightXml = [
    signatureRun ? signatureImageParagraph(signatureRun) : "",
    textParagraph(displayName, "right"),
  ]
    .filter(Boolean)
    .join("");

  return `${emptyParagraph()}${emptyParagraph()}<w:tbl><w:tblPr><w:tblW w:w="9640" w:type="dxa"/><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="5200"/><w:gridCol w:w="4440"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5200" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>${leftXml || textParagraph("")}</w:tc><w:tc><w:tcPr><w:tcW w:w="4440" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr>${rightXml || textParagraph("")}</w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;
};

/** Insert approval block after the signer footer (bottom of the certificate). */
const insertApprovalAfterSigner = (xml: string, approvalXml: string) => {
  if (!approvalXml.trim()) return xml;
  // Prefer right after the last table (signer footer we just injected).
  const tables = [...xml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)];
  const lastTable = tables.at(-1);
  if (lastTable && typeof lastTable.index === "number") {
    const end = lastTable.index + lastTable[0].length;
    return `${xml.slice(0, end)}${approvalXml}${xml.slice(end)}`;
  }
  // Fallback: append before </w:body>.
  const bodyClose = xml.lastIndexOf("</w:body>");
  if (bodyClose >= 0) {
    return `${xml.slice(0, bodyClose)}${approvalXml}${xml.slice(bodyClose)}`;
  }
  return `${xml}${approvalXml}`;
};

/** Two columns: titles/rank/date | floating signature (center, wide) + name + 3 blank lines. */
const buildSignerFooterTable = ({
  titleLines,
  rank,
  fullName,
  date,
  signatureRun,
}: {
  titleLines: [string, string, string];
  rank: string;
  fullName: string;
  date: string;
  signatureRun: string;
}) => {
  const leftLines = [
    ...titleLines.map((line) => line.trim()).filter(Boolean),
    rank.trim(),
    date.trim(),
  ].filter(Boolean);
  const leftXml = leftLines.map((line) => textParagraph(line, "left")).join("");
  const displayName = (
    formatGivenSurname(fullName) || withFallback(fullName)
  ).replace(/\s+/g, "\u00A0");
  const rightXml = [
    signatureRun ? signatureImageParagraph(signatureRun) : textParagraph(""),
    textParagraph(displayName, "right"),
  ].join("");

  return `<w:tbl><w:tblPr><w:tblW w:w="9640" w:type="dxa"/><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="5200"/><w:gridCol w:w="4440"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5200" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>${leftXml || textParagraph("")}</w:tc><w:tc><w:tcPr><w:tcW w:w="4440" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr>${rightXml}</w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;
};

const replaceSignerFooterWithTable = (xml: string, tableXml: string) => {
  const matches = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
  let start = -1;
  let end = -1;
  for (const match of matches) {
    const text = paragraphJoinedText(match[0]);
    const trimmed = text.trim();
    if (
      start < 0 &&
      (text === SAMPLE.signerTitle1 || trimmed === SAMPLE.signerTitle1.trim())
    ) {
      start = match.index ?? -1;
    }
    if (
      start >= 0 &&
      (trimmed === SAMPLE.signerDate || text === SAMPLE.signerDate)
    ) {
      end = (match.index ?? 0) + match[0].length;
    }
  }
  if (start < 0 || end < 0 || end <= start) return xml;
  return `${xml.slice(0, start)}${tableXml}${xml.slice(end)}`;
};

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

const readU32BE = (bytes: Uint8Array, offset: number) =>
  ((bytes[offset] << 24) |
    (bytes[offset + 1] << 16) |
    (bytes[offset + 2] << 8) |
    bytes[offset + 3]) >>>
  0;

const readU16BE = (bytes: Uint8Array, offset: number) =>
  (bytes[offset] << 8) | bytes[offset + 1];

const readRasterSize = (
  bytes: Uint8Array,
): { width: number; height: number } | null => {
  if (
    bytes.length >= 24 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) {
    const width = readU32BE(bytes, 16);
    const height = readU32BE(bytes, 20);
    return width > 0 && height > 0 ? { width, height } : null;
  }
  if (bytes.length > 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = bytes[offset + 1];
      if (marker === 0xc0 || marker === 0xc1 || marker === 0xc2) {
        const height = readU16BE(bytes, offset + 5);
        const width = readU16BE(bytes, offset + 7);
        return width > 0 && height > 0 ? { width, height } : null;
      }
      if (marker === 0xd9 || marker === 0xda) break;
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd9)) {
        offset += 2;
        continue;
      }
      const length = readU16BE(bytes, offset + 2);
      if (length < 2) break;
      offset += 2 + length;
    }
  }
  return null;
};

const extentForImage = (imageWidth: number, imageHeight: number) => {
  const ratio =
    imageWidth > 0 && imageHeight > 0 ? imageWidth / imageHeight : 1.294;
  let cy: number = SIGNATURE_LAYOUT.targetCy;
  let cx = Math.max(1, Math.round(cy * ratio));
  if (cx > SIGNATURE_LAYOUT.maxCx) {
    cx = SIGNATURE_LAYOUT.maxCx;
    cy = Math.max(1, Math.round(cx / ratio));
  }
  return { cx, cy };
};

const addSignatureImage = async (
  zip: JSZip,
  signatureData: string,
  relId: string,
) => {
  const { bytes, isPng } = dataUrlParts(signatureData);
  const relsPath = "word/_rels/document.xml.rels";
  const typesPath = "[Content_Types].xml";
  const rels = await zip.file(relsPath)?.async("string");
  const types = await zip.file(typesPath)?.async("string");
  if (!rels || !types) return bytes;

  const ext = isPng ? "png" : "jpeg";
  const fileName = `${relId}-sign.${ext}`;
  writeZipFile(zip, `word/media/${fileName}`, bytes);

  if (!rels.includes(`Id="${relId}"`)) {
    writeZipFile(
      zip,
      relsPath,
      rels.replace(
        "</Relationships>",
        `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${fileName}"/></Relationships>`,
      ),
    );
  } else {
    writeZipFile(
      zip,
      relsPath,
      rels.replace(
        new RegExp(`(<Relationship Id="${relId}"[^>]*Target=")media/[^"]+(")`),
        `$1media/${fileName}$2`,
      ),
    );
  }

  if (!isPng && !types.includes('Extension="jpeg"')) {
    writeZipFile(
      zip,
      typesPath,
      types.replace(
        "</Types>",
        '<Default Extension="jpeg" ContentType="image/jpeg"/></Types>',
      ),
    );
  }
  return bytes;
};

export const createZhbdCertificateWordBlob = async (
  fields: ZhbdCertificateFields,
) => {
  const response = await fetch(TEMPLATE_URL);
  if (!response.ok) {
    throw new Error("Не знайшов шаблон довідки ЖБД.");
  }
  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) {
    throw new Error("Шаблон довідки ЖБД пошкоджений.");
  }

  const signer =
    fields.signatories.find(
      (item) => String(item.blockType).toUpperCase() === "SIGNER",
    ) ??
    fields.signatories.find(
      (item) => String(item.blockType).toUpperCase() !== "APPROVAL",
    ) ??
    null;
  const approval =
    fields.signatories.find(
      (item) => String(item.blockType).toUpperCase() === "APPROVAL",
    ) ?? null;
  const signerParts = splitSigner(signer);
  const approvalParts = splitSigner(approval);
  const signerDate = signerParts.date || fields.date.trim();

  // Keep SAMPLE signer markers so we can swap the footer for a 2-column table.
  let filled = replaceParagraphTexts(documentXml, (text) => {
    const trimmed = text.trim();
    if (trimmed === SAMPLE.body.trim() || text === SAMPLE.body) {
      return withFallback(fields.bodyParagraph, SAMPLE.body);
    }
    if (trimmed === SAMPLE.basisOrders.trim() || text === SAMPLE.basisOrders) {
      return withFallback(fields.basisOrders, SAMPLE.basisOrders);
    }
    if (
      trimmed === SAMPLE.basisJournal.trim() ||
      text === SAMPLE.basisJournal
    ) {
      return withFallback(fields.basisJournal, SAMPLE.basisJournal);
    }
    if (trimmed === SAMPLE.headerDate || text === SAMPLE.headerDate) {
      return withFallback(fields.headerDate, SAMPLE.headerDate);
    }
    if (trimmed === SAMPLE.documentNumber || text === SAMPLE.documentNumber) {
      return withFallback(fields.documentNumber, SAMPLE.documentNumber);
    }
    return text;
  });

  let signatureRun = "";
  if (signer?.signatureData) {
    const bytes = await addSignatureImage(
      zip,
      signer.signatureData,
      SIGNATURE_LAYOUT.relId,
    );
    const size = bytes ? readRasterSize(bytes) : null;
    const { cx, cy } = extentForImage(size?.width ?? 0, size?.height ?? 0);
    signatureRun = frontOfTextSignatureDrawing(
      SIGNATURE_LAYOUT.relId,
      cx,
      cy,
      1746884461,
    );
  }

  const table = buildSignerFooterTable({
    titleLines: signerParts.titleLines,
    rank: signerParts.rank || signer?.rank || "",
    fullName: signer?.fullName ?? "",
    date: signerDate,
    signatureRun,
  });

  let approvalSignatureRun = "";
  if (approval?.signatureData) {
    const bytes = await addSignatureImage(
      zip,
      approval.signatureData,
      APPROVAL_SIGNATURE_LAYOUT.relId,
    );
    const size = bytes ? readRasterSize(bytes) : null;
    const { cx, cy } = extentForImage(size?.width ?? 0, size?.height ?? 0);
    approvalSignatureRun = frontOfTextSignatureDrawing(
      APPROVAL_SIGNATURE_LAYOUT.relId,
      cx,
      cy,
      1746884462,
    );
  }
  const approvalXml = approval
    ? buildApprovalBlock({
        titleLines: approvalParts.titleLines,
        rank: approvalParts.rank || approval.rank || "",
        fullName: approval.fullName ?? "",
        signatureRun: approvalSignatureRun,
      })
    : "";

  // Keep signer table + approval as one replacement so «ЗАТВЕРДЖУЮ» cannot be dropped.
  filled = replaceSignerFooterWithTable(filled, `${table}${approvalXml}`);
  if (approvalXml && !filled.includes("ЗАТВЕРДЖУЮ")) {
    filled = insertApprovalAfterSigner(filled, approvalXml);
  }

  zip.file("word/document.xml", filled, { createFolders: false });
  await stripRedColorInWordZip(zip);
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) zip.remove(name);
  }
  return zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
};
