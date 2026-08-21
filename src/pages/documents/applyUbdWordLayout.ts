import { ubdWordFormat } from "./ubdWordFormat";
import { ubdWordTexts } from "./ubdWordTexts";

const applyPage = (xml: string) => {
  const { page } = ubdWordFormat;
  return xml
    .replace(
      /<w:pgSz w:w="\d+" w:h="\d+"\s*\/>/,
      `<w:pgSz w:w="${page.width}" w:h="${page.height}"/>`,
    )
    .replace(
      /<w:pgMar w:top="\d+" w:right="\d+" w:bottom="\d+" w:left="\d+" w:header="\d+" w:footer="\d+" w:gutter="\d+"\s*\/>/,
      `<w:pgMar w:top="${page.marginTop}" w:right="${page.marginRight}" w:bottom="${page.marginBottom}" w:left="${page.marginLeft}" w:header="${page.headerDistance}" w:footer="${page.footerDistance}" w:gutter="${page.gutter}"/>`,
    )
    .replace(/<w:cols w:space="\d+"\s*\/>/, `<w:cols w:space="${page.columnSpace}"/>`)
    .replace(
      /<w:docGrid w:linePitch="\d+"\s*\/>/,
      `<w:docGrid w:linePitch="${page.linePitch}"/>`,
    );
};

const applyTable = (xml: string) => {
  const { table } = ubdWordFormat;
  const widths = [...table.columns, table.phantomColumn];
  const tableWidth = widths.reduce((sum, width) => sum + width, 0);
  let columnIndex = 0;
  let rowIndex = 0;
  const rowHeights = [
    table.headerRow1Height,
    table.headerRow2Height,
    table.dataRowHeight,
  ];

  let next = xml
    .replace(
      /<w:tblW w:w="\d+" w:type="dxa"\s*\/>/,
      `<w:tblW w:w="${tableWidth}" w:type="dxa"/>`,
    )
    .replace(
      /<w:tblInd w:w="\d+" w:type="dxa"\s*\/>/,
      `<w:tblInd w:w="${table.indent}" w:type="dxa"/>`,
    )
    .replace(/<w:gridCol w:w="\d+"\s*\/>/g, () => {
      const width = widths[columnIndex] ?? table.phantomColumn;
      columnIndex += 1;
      return `<w:gridCol w:w="${width}"/>`;
    })
    .replace(
      /<w:wAfter w:w="\d+" w:type="dxa"\s*\/>/g,
      `<w:wAfter w:w="${table.phantomColumn}" w:type="dxa"/>`,
    )
    .replace(/<w:trHeight w:val="\d+"\s*\/>/g, () => {
      const height = rowHeights[rowIndex] ?? table.dataRowHeight;
      rowIndex += 1;
      return `<w:trHeight w:val="${height}"/>`;
    })
    .replace(/w:sz="4"/g, `w:sz="${table.borderSize}"`);

  const padding = `<w:tblCellMar><w:top w:w="${table.cellPaddingTop}" w:type="dxa"/><w:left w:w="${table.cellPaddingLeft}" w:type="dxa"/><w:bottom w:w="${table.cellPaddingBottom}" w:type="dxa"/><w:right w:w="${table.cellPaddingRight}" w:type="dxa"/></w:tblCellMar>`;
  if (next.includes("<w:tblCellMar>")) {
    next = next.replace(/<w:tblCellMar>[\s\S]*?<\/w:tblCellMar>/, padding);
  } else {
    next = next.replace(
      /<w:tblLayout w:type="fixed"\s*\/>/,
      `<w:tblLayout w:type="fixed"/>${padding}`,
    );
  }

  const cellWidths = [
    ...table.columns,
    ...table.columns,
    table.phantomColumn,
    ...table.columns,
  ];
  let cellIndex = 0;
  next = next.replace(/<w:tcW w:w="\d+" w:type="dxa"\s*\/>/g, () => {
    const width = cellWidths[cellIndex] ?? table.columns[0];
    cellIndex += 1;
    return `<w:tcW w:w="${width}" w:type="dxa"/>`;
  });

  return next;
};

const applyParagraphs = (xml: string) => {
  const { classification, commander, basis, annex, approval, page2, font } =
    ubdWordFormat;
  let indentIndex = 0;

  return xml
    .replace(
      /<w:spacing w:before="\d+" w:after="\d+" w:line="\d+" w:lineRule="auto"\s*\/><w:ind w:left="\d+"\s*\/>/,
      `<w:spacing w:before="${classification.spacingBefore}" w:after="${classification.spacingAfter}" w:line="${font.line}" w:lineRule="auto" /><w:ind w:left="${classification.indentLeft}" />`,
    )
    .replace(
      /<w:spacing w:before="\d+"\s*\/><w:jc w:val="right"\s*\/>/,
      `<w:spacing w:before="${commander.spacingBefore}" /><w:jc w:val="right" />`,
    )
    .replace(/<w:ind w:left="\d+" w:firstLine="\d+"\s*\/>/g, () => {
      const block = indentIndex === 0 ? basis : annex;
      indentIndex += 1;
      return `<w:ind w:left="${block.indentLeft}" w:firstLine="${block.firstLine}" />`;
    })
    .replace(
      /<w:spacing w:before="\d+" w:after="\d+" w:line="\d+" w:lineRule="auto"\s*\/><w:ind w:right="\d+"\s*\/>/,
      `<w:spacing w:before="${approval.spacingBefore}" w:after="${approval.spacingAfter}" w:line="${font.line}" w:lineRule="auto" /><w:ind w:right="${approval.indentRight}" />`,
    )
    .replace(
      /<w:spacing w:before="\d+" w:after="\d+" w:line="\d+" w:lineRule="auto"\s*\/><w:ind w:right="\d+"\s*\/><w:jc w:val="center"\s*\/>/,
      `<w:spacing w:before="${page2.titleSpacingBefore}" w:after="0" w:line="${font.line}" w:lineRule="auto" /><w:ind w:right="${page2.titleIndentRight}" /><w:jc w:val="center" />`,
    )
    .replace(
      /<w:spacing w:after="\d+" w:line="\d+" w:lineRule="auto"\s*\/><w:ind w:left="\d+"\s*\/><w:jc w:val="center"\s*\/>/g,
      `<w:spacing w:after="0" w:line="${font.line}" w:lineRule="auto" /><w:ind w:left="${page2.noteIndentLeft}" /><w:jc w:val="center" />`,
    )
    .replace(/w:line="\d+"/g, `w:line="${font.line}"`)
    .replace(/<w:sz w:val="28"\s*\/>/g, `<w:sz w:val="${font.body}" />`)
    .replace(/<w:szCs w:val="28"\s*\/>/g, `<w:szCs w:val="${font.body}" />`)
    .replace(/<w:sz w:val="24"\s*\/>/g, `<w:sz w:val="${font.tableHeader}" />`)
    .replace(/<w:szCs w:val="24"\s*\/>/g, `<w:szCs w:val="${font.tableHeader}" />`)
    .replace(/<w:color w:val="EE0000"\s*\/>/g, `<w:color w:val="${font.color}" />`)
    .replace(/<w:color w:val="FF0000"\s*\/>/g, `<w:color w:val="${font.color}" />`);
};

const applyTexts = (xml: string) => {
  const texts = ubdWordTexts;
  const [openWord, ...openRest] = texts.classification.split(" ");
  const openTail = openRest.join(" ");
  const [page2Word, ...page2Rest] = texts.page2Title.split(" ");
  const page2Tail = page2Rest.join(" ");

  return xml
    .replace(">Відкрита<", `>${openWord}<`)
    .replace(">інформація<", `>${openTail || "інформація"}<`)
    .replace(">РАПОРТ<", `>${texts.title}<`)
    .replace(">ЗАТВЕРДЖУЮ<", `>${texts.approve}<`)
    .replace(">ВІДКРИТА<", `>${page2Word}<`)
    .replace(">ІНФОРМАЦІЯ<", `>${page2Tail || "ІНФОРМАЦІЯ"}<`)
    .replace(
      "(Обмежено в розповсюдженні – лише для Збройних Сил України)",
      texts.page2Note,
    )
    .replace(
      "    Підстава: ",
      `${texts.basisSpaces}${texts.basisLabel} `,
    )
    .replace(
      /(Підстава:\s*)(?!\{\{BASIS\}\})([^<{]*бойове розпорядження[\s\S]*?від\s*\d{2}\.\d{2}\.\d{4})/i,
      `$1{{BASIS}}`,
    )
    .replace(
      "    Додаток: копія паспорта, РНОКПП, дві фотокартки",
      `${texts.annexSpaces}${texts.annex}`,
    )
    .replaceAll("військової частини А4862", texts.unitName);
};

const contentWidthTwips = () => {
  const { page, signature } = ubdWordFormat;
  return (
    signature.nameTabRight ||
    page.width - page.marginLeft - page.marginRight
  );
};

const applyRightNameTab = (xml: string, rankToken: string, nameToken: string) => {
  const textIndex = xml.indexOf(rankToken);
  if (textIndex < 0) return xml;
  const paragraphStart = xml.lastIndexOf("<w:p ", textIndex);
  const paragraphEnd = xml.indexOf("</w:p>", textIndex);
  if (paragraphStart < 0 || paragraphEnd < 0) return xml;

  const paragraph = xml.slice(paragraphStart, paragraphEnd + 6);
  const pPrClose = paragraph.indexOf("</w:pPr>");
  if (pPrClose < 0) return xml;

  let pPr = paragraph.slice(0, pPrClose + "</w:pPr>".length);
  const tabPos = contentWidthTwips();
  const rightTabs = `<w:tabs><w:tab w:val="right" w:pos="${tabPos}"/></w:tabs>`;
  pPr = pPr.includes("<w:tabs>")
    ? pPr.replace(/<w:tabs>[\s\S]*?<\/w:tabs>/, rightTabs)
    : pPr.replace("<w:pPr>", `<w:pPr>${rightTabs}`);
  pPr = pPr.replace(/<w:jc w:val="both"\s*\/>/, `<w:jc w:val="left"/>`);

  const rPrMatch = paragraph.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  const rPr =
    rPrMatch?.[0] ??
    `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>`;

  const rebuilt = `${pPr}<w:r>${rPr}<w:t>${rankToken}</w:t></w:r><w:r>${rPr}<w:tab/><w:t>${nameToken}</w:t></w:r></w:p>`;
  return `${xml.slice(0, paragraphStart)}${rebuilt}${xml.slice(paragraphEnd + 6)}`;
};

const applySignatureNameAlignment = (xml: string) =>
  applyRightNameTab(
    applyRightNameTab(xml, "{{SIGNER_RANK}}", "{{SIGNER_NAME}}"),
    "{{APPROVER_RANK}}",
    "{{APPROVER_NAME}}",
  );

const keepParagraph = (xml: string, marker: string, withNext: boolean) => {
  const textIndex = xml.indexOf(marker);
  if (textIndex < 0) return xml;
  const paragraphStart = xml.lastIndexOf("<w:p ", textIndex);
  if (paragraphStart < 0) return xml;
  const pPrOpen = xml.indexOf("<w:pPr>", paragraphStart);
  const paragraphEnd = xml.indexOf("</w:p>", textIndex);
  if (pPrOpen < 0 || pPrOpen > paragraphEnd) return xml;
  const innerStart = pPrOpen + "<w:pPr>".length;
  if (xml.slice(innerStart, paragraphEnd).includes("<w:keepLines")) return xml;
  const tags = withNext ? "<w:keepNext/><w:keepLines/>" : "<w:keepLines/>";
  return `${xml.slice(0, innerStart)}${tags}${xml.slice(innerStart)}`;
};

/** Не розривати блоки підпису й «ЗАТВЕРДЖУЮ» між сторінками. */
const applyKeepTogether = (xml: string) => {
  let next = xml;
  for (const marker of [
    `>${ubdWordTexts.approve}<`,
    "{{APPROVER_TITLE_1}}",
    "{{APPROVER_TITLE_2}}",
    "{{SIGNER_TITLE_1}}",
    "{{SIGNER_TITLE_2}}",
    "{{SIGNER_TITLE_3}}",
    "{{SIGNER_RANK}}",
  ]) {
    next = keepParagraph(next, marker, true);
  }
  for (const marker of ["{{APPROVER_RANK}}", "{{SIGNER_DATE}}"]) {
    next = keepParagraph(next, marker, false);
  }
  return next;
};

const PAGE_BREAK_PARAGRAPH = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

const ensurePage2Break = (xml: string) => {
  if (!ubdWordFormat.page2.pageBreakBefore) return xml;

  const titleWord = ubdWordTexts.page2Title.split(" ")[0] || "ВІДКРИТА";
  const marker = `>${titleWord}<`;
  const textIndex = xml.indexOf(marker);
  if (textIndex < 0) return xml;

  const paragraphStart = xml.lastIndexOf("<w:p ", textIndex);
  if (paragraphStart < 0) return xml;

  const before = xml.slice(0, paragraphStart);
  if (before.includes('w:type="page"')) return xml;

  const paragraphXml = xml.slice(paragraphStart);
  return `${before}${PAGE_BREAK_PARAGRAPH}${paragraphXml}`;
};

export const applyUbdWordLayout = (documentXml: string) =>
  ensurePage2Break(
    applyKeepTogether(
      applySignatureNameAlignment(
        applyTable(applyParagraphs(applyPage(applyTexts(documentXml)))),
      ),
    ),
  );

export const applyUbdWordSettings = (settingsXml: string) =>
  settingsXml.replace(
    /<w:defaultTabStop w:val="\d+"\s*\/>/,
    `<w:defaultTabStop w:val="${ubdWordFormat.page.defaultTabStop}"/>`,
  );
