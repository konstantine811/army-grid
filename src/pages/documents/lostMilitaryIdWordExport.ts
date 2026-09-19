import JSZip from "jszip";
import {
  AlignmentType,
  Document,
  Footer,
  ImageRun,
  Packer,
  PageNumber,
  Paragraph,
  TabStopType,
  TextRun,
} from "docx";
import {
  buildLostMilitaryIdReportText,
  buildLostMilitaryIdOrderText,
  buildLostMilitaryIdPersonCard,
  buildLostMilitaryIdActCircumstances,
  buildLostMilitaryIdActConclusions,
  buildLostMilitaryIdActProposals,
  buildLostMilitaryIdActAttachments,
  buildLostMilitaryIdActDutySection,
  buildLostMilitaryIdPersonExplanation,
  actApprovalDateLine,
  approvalFooterBlock,
  buildLostMilitaryIdReportDateLine,
  declinedPerson,
  investigatorFooterBlock,
  instrumentalInvestigatorLine,
  lostMilitaryIdServicemanUnitPhrase,
  normalizeMilitaryUnitPhrase,
  orderFooterBlock,
  reporterFooterBlock,
  type LostMilitaryIdFields,
} from "./lostMilitaryIdReport";
import { formatNominativeGivenSurname } from "./lostMilitaryIdCases";

const FONT = "Times New Roman";

const run = (text: string, opts?: { bold?: boolean; size?: number }) =>
  new TextRun({
    text,
    bold: opts?.bold,
    font: FONT,
    size: (opts?.size ?? 28) * 0.5 * 2,
  });

const para = (
  text: string,
  opts?: {
    bold?: boolean;
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
    indent?: boolean;
    spacingAfter?: number;
  },
) =>
  new Paragraph({
    alignment: opts?.align ?? AlignmentType.BOTH,
    spacing: { after: opts?.spacingAfter ?? 200, line: 276 },
    indent: opts?.indent ? { firstLine: 567 } : undefined,
    children: [run(text, { bold: opts?.bold })],
  });

const empty = () => new Paragraph({ children: [] });

const RIGHT_TAB = { type: TabStopType.RIGHT, position: 9000 };

type SignatoryFooterParts = {
  titleLines: string[];
  rank: string;
  name: string;
  signatureData?: string;
};

/** ЗАТВЕРДЖУЮ: посада, звання і ПІБ — одна колонка справа. */
const buildSignatoryRightBlock = (footer: SignatoryFooterParts): Paragraph[] => {
  const paragraphs: Paragraph[] = [];
  for (const line of footer.titleLines) {
    if (line.trim()) {
      paragraphs.push(
        para(line, { align: AlignmentType.RIGHT, spacingAfter: 40 }),
      );
    }
  }
  const signatureRun = footer.signatureData
    ? dataUrlToImageRun(footer.signatureData)
    : null;
  if (signatureRun) {
    paragraphs.push(
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { after: 40, line: 276 },
        children: [signatureRun],
      }),
    );
  }
  paragraphs.push(
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { after: 40, line: 276 },
      tabStops: [RIGHT_TAB],
      children: [
        run(footer.rank.trim() || "________________"),
        new TextRun({ text: "\t", font: FONT, size: 28 }),
        run(footer.name.trim() || "________________"),
      ],
    }),
  );
  return paragraphs;
};

/** Два рядки без таблиці: посада зліва, звання і ПІБ справа (таб-стоп). */
const buildSignatoryPositionNameParagraphs = (
  footer: SignatoryFooterParts,
): Paragraph[] => {
  const titles = footer.titleLines.map((line) => line.trim()).filter(Boolean);
  const line1 = titles[0] ?? "";
  const line2 = titles[1] ?? "";
  const rank = footer.rank.trim() || "________________";
  const name = footer.name.trim() || "________________";
  const signatureRun = footer.signatureData
    ? dataUrlToImageRun(footer.signatureData)
    : null;

  const row2Children: (TextRun | ImageRun)[] = [
    run(line2 || " "),
    new TextRun({ text: "\t", font: FONT, size: 28 }),
  ];
  if (signatureRun) {
    row2Children.push(signatureRun);
    row2Children.push(new TextRun({ text: " ", font: FONT, size: 28 }));
  }
  row2Children.push(run(name));

  return [
    new Paragraph({
      spacing: { after: 40, line: 276 },
      tabStops: [RIGHT_TAB],
      children: [
        run(line1 || " "),
        new TextRun({ text: "\t", font: FONT, size: 28 }),
        run(rank),
      ],
    }),
    new Paragraph({
      spacing: { after: 80, line: 276 },
      tabStops: [RIGHT_TAB],
      children: row2Children,
    }),
  ];
};

const dataUrlToImageRun = (dataUrl: string) => {
  const match = dataUrl.trim().match(/^data:([^;]+);base64,(.+)$/);
  if (!match) return null;
  const binary = atob(match[2]);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new ImageRun({
    type: match[1].includes("png") ? "png" : "jpg",
    data: bytes,
    transformation: { width: 110, height: 46 },
  });
};

const splitBlocks = (text: string) =>
  text
    .split(/\n{2,}/)
    .map((block) => block.trim())
    .filter(Boolean);

const pageFooter = () =>
  new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            children: [PageNumber.CURRENT],
            font: FONT,
            size: 20,
          }),
        ],
      }),
    ],
  });

const buildReportDocument = (fields: LostMilitaryIdFields) => {
  const footer = reporterFooterBlock(fields);
  const body = splitBlocks(buildLostMilitaryIdReportText(fields));
  return new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1134, right: 851, bottom: 1134, left: 1418 },
          },
        },
        footers: { default: pageFooter() },
        children: [
          para(fields.addressee.trim() || "Командиру військової частини А4862", {
            align: AlignmentType.RIGHT,
            spacingAfter: 280,
          }),
          para("РАПОРТ", {
            bold: true,
            align: AlignmentType.CENTER,
            spacingAfter: 280,
          }),
          ...body.map((block) => para(block, { indent: true })),
          empty(),
          ...buildSignatoryPositionNameParagraphs({
            titleLines: footer.titleLines,
            rank: footer.rank,
            name: footer.name,
            signatureData: footer.signatureData,
          }),
          para(buildLostMilitaryIdReportDateLine(fields), {
            align: AlignmentType.RIGHT,
            spacingAfter: 0,
          }),
        ],
      },
    ],
  });
};

const buildOrderDocument = (fields: LostMilitaryIdFields) => {
  const unit = normalizeMilitaryUnitPhrase(fields.militaryUnit);
  const commander = orderFooterBlock(fields);
  const number = fields.orderNumber.trim() || "______";
  const date = fields.orderDate.trim() || fields.reportDate.trim() || "____.____.______";
  const body = splitBlocks(buildLostMilitaryIdOrderText(fields));
  return new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1134, right: 851, bottom: 1134, left: 1418 },
          },
        },
        footers: { default: pageFooter() },
        children: [
          para("НАКАЗ", { bold: true, align: AlignmentType.CENTER, spacingAfter: 80 }),
          para(`командира ${unit}`, {
            align: AlignmentType.CENTER,
            spacingAfter: 40,
          }),
          para("(з адміністративно-господарської діяльності)", {
            align: AlignmentType.CENTER,
            spacingAfter: 200,
          }),
          para(`${date}                                                                 № ${number}`, {
            align: AlignmentType.BOTH,
          }),
          para("Про призначення службового розслідування", {
            bold: true,
            align: AlignmentType.CENTER,
          }),
          ...body.map((block) => para(block, { indent: true })),
          empty(),
          ...buildSignatoryPositionNameParagraphs({
            titleLines: commander.titleLines,
            rank: commander.rank,
            name: commander.name,
            signatureData: commander.signatureData,
          }),
        ],
      },
    ],
  });
};

const buildActDocument = (fields: LostMilitaryIdFields) => {
  const person = declinedPerson(fields);
  const unit = normalizeMilitaryUnitPhrase(fields.militaryUnit);
  const servicemanUnit = lostMilitaryIdServicemanUnitPhrase(fields);
  const investigatorLine = instrumentalInvestigatorLine(fields);
  const approval = approvalFooterBlock(fields);
  const investigatorFooter = investigatorFooterBlock(fields);
  const orderLabel =
    fields.orderNumber.trim() && fields.orderDate.trim()
      ? `наказу командира ${unit} від ${fields.orderDate} №${fields.orderNumber} «Про призначення службового розслідування»`
      : `наказу командира ${unit} «Про призначення службового розслідування»`;
  const legal = splitBlocks(buildLostMilitaryIdActCircumstances(fields));
  const conclusions = buildLostMilitaryIdActConclusions(fields);
  const proposals = buildLostMilitaryIdActProposals(fields);
  const attachments = buildLostMilitaryIdActAttachments(fields);
  const dutySection = buildLostMilitaryIdActDutySection(fields);
  const personExplanation = buildLostMilitaryIdPersonExplanation(fields);
  const actTitleTail = `службового розслідування за фактом втрати військового квитка військовослужбовцем ${unit}${servicemanUnit} ${person.rankInstrumental} ${person.instrumental}`;

  return new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 851, right: 851, bottom: 851, left: 1134 },
          },
        },
        footers: { default: pageFooter() },
        children: [
          para("ЗАТВЕРДЖУЮ", {
            bold: true,
            align: AlignmentType.RIGHT,
            spacingAfter: 40,
          }),
          ...buildSignatoryRightBlock({
            titleLines: approval.titleLines,
            rank: approval.rank,
            name: approval.name,
            signatureData: approval.signatureData,
          }),
          para(actApprovalDateLine(fields), {
            align: AlignmentType.RIGHT,
            spacingAfter: 280,
          }),
          para("АКТ", {
            bold: true,
            align: AlignmentType.CENTER,
            spacingAfter: 40,
          }),
          para(actTitleTail, {
            bold: true,
            align: AlignmentType.CENTER,
            spacingAfter: 280,
          }),
          para(
            `Відповідно до вимог статті 85 Статуту внутрішньої служби ЗС України, Порядку проведення службового розслідування у ЗС України, затвердженого наказом Міністерства оборони України від 21.11.2017 № 608 (зі змінами) та ${orderLabel}, мною, ${
              investigatorLine || "________________"
            }, було проведено службове розслідування за фактом втрати військового квитка ${person.positionInstrumental} ${unit}${servicemanUnit} ${person.rankInstrumental} ${person.instrumental}.`,
            { indent: true },
          ),
          para("1. Нормативно-правова база:", { bold: true }),
          para("1.1. Конституція України.", { indent: true, spacingAfter: 80 }),
          para("1.2. Дисциплінарний статут Збройних Сил України.", {
            indent: true,
            spacingAfter: 80,
          }),
          para("1.3. Статут внутрішньої служби Збройних Сил України.", {
            indent: true,
            spacingAfter: 80,
          }),
          para(
            "1.4. «Порядок проведення службового розслідування у Збройних Силах України», затверджений наказом Міністерства оборони України від 21.11.2017 № 608 (зі змінами), зареєстрований в Міністерстві юстиції України від 13.12.2017 № 1503/31371.",
            { indent: true, spacingAfter: 80 },
          ),
          para(
            "1.5. Наказ МОУ від 10.04.2017 № 206 «Про військовий квиток осіб рядового, сержантського і старшинського складу».",
            { indent: true },
          ),
          para(
            `2. Опис обставин події (втрати військового квитка ${person.rankInstrumental} ${person.instrumental}).`,
            { bold: true },
          ),
          ...legal.map((block) => para(block, { indent: true })),
          para(
            "3. Обставини та факти, що були встановлені під час проведення службового розслідування.",
            { bold: true },
          ),
          para(
            "3.1. Відповідно до вимог статті 16 Статуту внутрішньої служби Збройних Сил України на військовослужбовця покладено обов’язок виконувати службові обов’язки, що визначають обсяг виконання завдань, доручених йому за посадою.",
            { indent: true },
          ),
          para(
            "3.2. Відповідно до вимог статті 11 Статуту внутрішньої служби Збройних Сил України військовослужбовець зобов’язаний берегти державне майно.",
            { indent: true },
          ),
          para(
            "3.3. Відповідно до вимог статті 128 Статуту внутрішньої служби ЗС України військовослужбовець зобов’язаний сумлінно вивчати військову справу, зразково виконувати свої службові обов’язки.",
            { indent: true },
          ),
          para(
            "3.4. Відповідно до вимог статті 4 Дисциплінарного статуту Збройних Сил України на військовослужбовця покладено обов’язок додержуватися Конституції та законів України, Військової присяги, неухильно виконувати вимоги статутів, накази командирів.",
            { indent: true },
          ),
          para(`3.5. ${personExplanation}`, { indent: true }),
          para("4. Відомості про осіб, стосовно яких призначено службове розслідування.", {
            bold: true,
          }),
          para(
            `Особисті дані військовослужбовця, стосовно якого проведено службове розслідування: ${buildLostMilitaryIdPersonCard(fields)}`,
            { indent: true },
          ),
          para("5. Неправомірні дії військовослужбовця та причинний зв’язок:", {
            bold: true,
          }),
          ...dutySection.map((line) => para(line, { indent: true })),
          para("6. Висновки службового розслідування", { bold: true }),
          para(conclusions, { indent: true }),
          para("Пропозиції:", { bold: true }),
          ...proposals.split("\n").map((line) => para(line, { indent: true })),
          para("7. До акту службового розслідування додаю:", { bold: true }),
          ...attachments.map((line) =>
            para(line, { indent: true, spacingAfter: 80 }),
          ),
          empty(),
          ...buildSignatoryPositionNameParagraphs({
            titleLines: investigatorFooter.titleLines.length
              ? investigatorFooter.titleLines
              : ["Особа, яка проводила службове розслідування"],
            rank: investigatorFooter.rank,
            name:
              investigatorFooter.name ||
              formatNominativeGivenSurname(fields.investigatorFullName),
          }),
        ],
      },
    ],
  });
};

export const createLostMilitaryIdWordBlob = (fields: LostMilitaryIdFields) =>
  Packer.toBlob(buildActDocument(fields));

export const createLostMilitaryIdReportWordBlob = (fields: LostMilitaryIdFields) =>
  Packer.toBlob(buildReportDocument(fields));

export const createLostMilitaryIdOrderWordBlob = (fields: LostMilitaryIdFields) =>
  Packer.toBlob(buildOrderDocument(fields));

export const createLostMilitaryIdActWordBlob = (fields: LostMilitaryIdFields) =>
  createLostMilitaryIdWordBlob(fields);

const safeZipEntryName = (value: string) =>
  String(value || "document")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export const createLostMilitaryIdKitZip = async (fields: LostMilitaryIdFields) => {
  const [reportBlob, orderBlob, actBlob] = await Promise.all([
    createLostMilitaryIdReportWordBlob(fields),
    createLostMilitaryIdOrderWordBlob(fields),
    createLostMilitaryIdActWordBlob(fields),
  ]);

  const base = safeZipEntryName(
    fields.folderName || fields.fullName || "Втрата військового квитка",
  );
  const zip = new JSZip();
  zip.file(`${base} · Рапорт.docx`, reportBlob);
  zip.file(`${base} · Наказ.docx`, orderBlob);
  zip.file(`${base} · Акт розслідування.docx`, actBlob);

  return zip.generateAsync({
    type: "blob",
    mimeType: "application/zip",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
};
