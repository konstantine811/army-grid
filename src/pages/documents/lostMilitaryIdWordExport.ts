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
import JSZip from "jszip";
import {
  buildLostMilitaryIdActCircumstances,
  buildLostMilitaryIdActConclusions,
  buildLostMilitaryIdActProposals,
  buildLostMilitaryIdOrderText,
  buildLostMilitaryIdPersonCard,
  buildLostMilitaryIdReportText,
  declinedPerson,
  reporterFooterBlock,
  reporterHeaderBlock,
  type LostMilitaryIdFields,
} from "./lostMilitaryIdReport";
import { capitalizeReportPosition } from "./reportPosition";

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
  const header = reporterHeaderBlock(fields);
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
            spacingAfter: 80,
          }),
          empty(),
          ...header.map((line) =>
            para(line, { align: AlignmentType.RIGHT, spacingAfter: 40 }),
          ),
          empty(),
          para("РАПОРТ", {
            bold: true,
            align: AlignmentType.CENTER,
            spacingAfter: 280,
          }),
          ...body.map((block) => para(block, { indent: true })),
          empty(),
          ...footer.titleLines.map((line) =>
            para(line, { spacingAfter: 40 }),
          ),
          empty(),
          new Paragraph({
            spacing: { after: 200 },
            tabStops: [
              { type: TabStopType.CENTER, position: 4680 },
              { type: TabStopType.RIGHT, position: 9000 },
            ],
            children: [
              run(footer.rank || "________________"),
              new TextRun({ text: "\t", font: FONT, size: 28 }),
              ...(footer.signatureData
                ? [dataUrlToImageRun(footer.signatureData) ?? run("________________")]
                : [run("________________")]),
              new TextRun({ text: "\t", font: FONT, size: 28 }),
              run(footer.name || "________________"),
            ],
          }),
          para(fields.reportDate.trim() || "____.____.______", {
            spacingAfter: 0,
          }),
        ],
      },
    ],
  });
};

const buildOrderDocument = (fields: LostMilitaryIdFields) => {
  const unit = fields.militaryUnit.trim() || "А4862";
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
          para(`командира військової частини ${unit}`, {
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
          para("Командир військової частини", { spacingAfter: 40 }),
          para(`військової частини ${unit}`, { spacingAfter: 200 }),
          para("________________", { spacingAfter: 0 }),
        ],
      },
    ],
  });
};

const buildActDocument = (fields: LostMilitaryIdFields) => {
  const person = declinedPerson(fields);
  const unit = fields.militaryUnit.trim() || "А4862";
  const investigator = [
    capitalizeReportPosition(fields.investigatorPosition),
    fields.investigatorRank.trim(),
    fields.investigatorFullName.trim(),
  ]
    .filter(Boolean)
    .join(" ");
  const orderLabel =
    fields.orderNumber.trim() && fields.orderDate.trim()
      ? `наказу командира військової частини ${unit} від ${fields.orderDate} №${fields.orderNumber}`
      : `наказу командира військової частини ${unit} «Про призначення службового розслідування»`;
  const legal = splitBlocks(buildLostMilitaryIdActCircumstances(fields));
  const conclusions = buildLostMilitaryIdActConclusions(fields);
  const proposals = buildLostMilitaryIdActProposals(fields);

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
          para("ЗАТВЕРДЖУЮ", { bold: true, align: AlignmentType.RIGHT, spacingAfter: 40 }),
          para(`Командир військової частини ${unit}`, {
            align: AlignmentType.RIGHT,
            spacingAfter: 40,
          }),
          para("________________", { align: AlignmentType.RIGHT, spacingAfter: 40 }),
          para("«____» ______________ ______ року", {
            align: AlignmentType.RIGHT,
            spacingAfter: 280,
          }),
          para(
            `АКТ службового розслідування за фактом втрати військового квитка військовослужбовцем військової частини ${unit} ${person.rankInstrumental} ${person.instrumental}`,
            { bold: true, align: AlignmentType.CENTER },
          ),
          para(
            `Відповідно до вимог статті 85 Статуту внутрішньої служби ЗС України, Порядку проведення службового розслідування у ЗС України, затвердженого наказом Міністерства оборони України від 21.11.2017 № 608 (зі змінами) та ${orderLabel}, мною, ${
              investigator || "________________"
            }, було проведено службове розслідування за фактом втрати військового квитка ${person.positionInstrumental} військової частини ${unit} ${person.rankInstrumental} ${person.instrumental}.`,
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
            "1.4. «Порядок проведення службового розслідування у Збройних Сил України», затверджений наказом Міністерства оборони України від 21.11.2017 № 608 (зі змінами).",
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
            "3.3. Відповідно до вимог статті 4 Дисциплінарного статуту Збройних Сил України військовослужбовець зобов’язаний додержуватися Конституції та законів України, Військової присяги, неухильно виконувати вимоги статутів, накази командирів.",
            { indent: true },
          ),
          para(
            `3.4. ${person.rankInstrumental} ${person.nominative} пояснив, що втрата військового квитка сталася ${
              fields.lossDate.trim() || "______"
            } ${circumstancesTextSafe(fields)}. Пошуки квитка результату не дали.`,
            { indent: true },
          ),
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
          para(
            `Втрата військового квитка відбулася ${circumstancesTextSafe(fields)}. Ознак умисних дій не встановлено.`,
            { indent: true },
          ),
          para("Заперечення, заяви та клопотання: не надходило.", {
            indent: true,
          }),
          para("6. Висновки службового розслідування", { bold: true }),
          para(conclusions, { indent: true }),
          para("Пропозиції:", { bold: true }),
          ...proposals.split("\n").map((line) => para(line, { indent: true })),
          para("7. До акту службового розслідування додаю:", { bold: true }),
          para(
            `1. Витяг з ${orderLabel} про призначення службового розслідування.`,
            { indent: true, spacingAfter: 80 },
          ),
          para(
            `2. Копію рапорту про втрату військового квитка ${person.rankInstrumental} ${person.instrumental}${
              fields.reportDate.trim() ? ` від ${fields.reportDate}` : ""
            }.`,
            { indent: true, spacingAfter: 80 },
          ),
          para(
            `3. Пояснення військовослужбовця ${person.rankGenitive} ${person.genitive}.`,
            { indent: true },
          ),
          empty(),
          ...(fields.investigatorPosition
            ? capitalizeReportPosition(fields.investigatorPosition)
                .split(/\n/)
                .map((line) => para(line, { spacingAfter: 40 }))
            : [para("Особа, яка проводила службове розслідування", { spacingAfter: 40 })]),
          new Paragraph({
            spacing: { after: 0 },
            tabStops: [{ type: TabStopType.RIGHT, position: 9000 }],
            children: [
              run(fields.investigatorRank.trim() || "________________"),
              new TextRun({ text: "\t", font: FONT, size: 28 }),
              run(fields.investigatorFullName.trim() || "________________"),
            ],
          }),
        ],
      },
    ],
  });
};

const circumstancesTextSafe = (fields: LostMilitaryIdFields) => {
  if (fields.circumstanceKind === "custom") {
    return fields.customCircumstances.trim() || "за встановлених обставин";
  }
  const from = fields.fromLocation.trim();
  const to = fields.toLocation.trim();
  if (from && to) return `під час переміщення з ${from} до ${to}`;
  return "під час переміщення";
};

const safePart = (value: string) =>
  value.replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim() || "документ";

export const createLostMilitaryIdReportWordBlob = (fields: LostMilitaryIdFields) =>
  Packer.toBlob(buildReportDocument(fields));

export const createLostMilitaryIdOrderWordBlob = (fields: LostMilitaryIdFields) =>
  Packer.toBlob(buildOrderDocument(fields));

export const createLostMilitaryIdActWordBlob = (fields: LostMilitaryIdFields) =>
  Packer.toBlob(buildActDocument(fields));

export const createLostMilitaryIdKitZip = async (fields: LostMilitaryIdFields) => {
  const folder = safePart(fields.folderName);
  const [report, order, act] = await Promise.all([
    createLostMilitaryIdReportWordBlob(fields),
    createLostMilitaryIdOrderWordBlob(fields),
    createLostMilitaryIdActWordBlob(fields),
  ]);
  const zip = new JSZip();
  zip.file(`${folder} · Рапорт.docx`, report);
  zip.file(`${folder} · Наказ.docx`, order);
  zip.file(`${folder} · Акт.docx`, act);
  return zip.generateAsync({
    type: "blob",
    mimeType: "application/zip",
  });
};
