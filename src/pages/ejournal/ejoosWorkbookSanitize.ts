import JSZip from "jszip";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const EXTERNAL_LINK_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";

const EXTERNAL_LINK_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";

const CALC_CHAIN_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

const CALC_CHAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";

const escapeRegExp = (value: string) =>
  value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const parseAttr = (attrs: string, name: string) =>
  attrs.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];

const columnLetterToNumber = (letter: string) => {
  let index = 0;
  for (const char of letter.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index;
};

const columnNumberToLetter = (columnNumber: number) => {
  let n = columnNumber;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
};

const A1_REF_IN_FORMULA = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g;

const shiftFormulaByCellDelta = (text: string, dCol: number, dRow: number) => {
  if (!dCol && !dRow) return text;
  return text.replace(
    A1_REF_IN_FORMULA,
    (match, colAbs, col, rowAbs, rowStr) => {
      let colNum = columnLetterToNumber(col);
      const row = Number(rowStr);
      if (!Number.isFinite(row) || colNum < 1) return match;
      if (!colAbs) colNum += dCol;
      const nextRow = rowAbs ? row : row + dRow;
      if (colNum < 1 || nextRow < 1) return match;
      return `${colAbs}${columnNumberToLetter(colNum)}${rowAbs}${nextRow}`;
    },
  );
};

const cellReferencePartsFromXml = (cellXml: string) => {
  const ref = cellXml.match(/<c\b[^<>]*?\br="([A-Z]+)([0-9]+)"/i);
  if (!ref) return null;
  return {
    row: Number(ref[2]),
    columnNumber: columnLetterToNumber(ref[1]),
  };
};

/**
 * Shared formula після вставки рядків на Табелі часто ламається — Excel відкриває
 * файл через Repair і видаляє `<f t="shared">`. Розгортаємо в звичайні формули.
 */
export const expandSharedFormulas = (sheetXml: string) => {
  if (!/\bt="shared"/i.test(sheetXml)) return sheetXml;

  const masters = new Map<
    string,
    { formula: string; col: number; row: number }
  >();
  for (const cell of sheetXml.matchAll(/<c\b[^<>]*?(?:\/>|>[\s\S]*?<\/c>)/gi)) {
    const cellXml = cell[0];
    const ref = cellReferencePartsFromXml(cellXml);
    if (!ref) continue;
    const formulaOpen = cellXml.match(
      /<f\b((?:(?!\/>)[^<>])*)>([\s\S]*?)<\/f>/i,
    );
    if (!formulaOpen) continue;
    const attrs = formulaOpen[1];
    if (!/\bt="shared"/i.test(attrs)) continue;
    const body = formulaOpen[2].trim();
    if (!body) continue;
    const si = parseAttr(attrs, "si");
    if (si == null) continue;
    masters.set(si, { formula: body, col: ref.columnNumber, row: ref.row });
  }
  if (!masters.size) return sheetXml;

  return sheetXml.replace(/<c\b[^<>]*?(?:\/>|>[\s\S]*?<\/c>)/gi, (cellXml) => {
    const ref = cellReferencePartsFromXml(cellXml);
    if (!ref) return cellXml;
    return cellXml.replace(
      /<f\b([^<>]*?)(?:\/>|>([\s\S]*?)<\/f>)/i,
      (formulaXml, attrs, body = "") => {
        if (!/\bt="shared"/i.test(String(attrs))) return formulaXml;
        const si = parseAttr(String(attrs), "si");
        if (si == null || !masters.has(si)) return "";
        const master = masters.get(si)!;
        const masterBody = body.trim();
        const dCol = ref.columnNumber - master.col;
        const dRow = ref.row - master.row;
        const expanded = masterBody
          ? masterBody
          : shiftFormulaByCellDelta(master.formula, dCol, dRow);
        return `<f>${expanded}</f>`;
      },
    );
  });
};

const hasStaleExternalWorkbookRef = (xml: string) =>
  /\[[^\]]+\.xlsx\]/i.test(xml) ||
  /externalLink/i.test(xml) ||
  /externalReference/i.test(xml);

const stripWorkbookExternalReferences = (xml: string) => {
  const withoutExternalRefs = xml.replace(
    /<externalReferences\b[^>]*>[\s\S]*?<\/externalReferences>/gi,
    "",
  );
  if (/<calcPr\b/i.test(withoutExternalRefs)) {
    return withoutExternalRefs.replace(/<calcPr\b([^>]*)\/>/i, (_match, attrs) => {
      let nextAttrs = String(attrs);
      if (!/\bcalcMode=/i.test(nextAttrs)) nextAttrs += ' calcMode="auto"';
      if (!/\bfullCalcOnLoad=/i.test(nextAttrs)) nextAttrs += ' fullCalcOnLoad="1"';
      if (!/\bforceFullCalc=/i.test(nextAttrs)) nextAttrs += ' forceFullCalc="1"';
      return `<calcPr${nextAttrs}/>`;
    });
  }
  return withoutExternalRefs.replace(
    /<\/workbook>\s*$/i,
    '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>',
  );
};

const stripWorkbookExternalRelationships = (xml: string) =>
  xml.replace(
    new RegExp(
      `<Relationship\\b(?=[^>]*\\bType="${escapeRegExp(EXTERNAL_LINK_REL)}")` +
        `[^>]*(?:/>|>\\s*</Relationship>)`,
      "gi",
    ),
    "",
  );

const stripWorkbookCalcChainRelationship = (xml: string) =>
  xml.replace(
    new RegExp(
      `<Relationship\\b(?=[^>]*\\bType="${escapeRegExp(CALC_CHAIN_REL)}")` +
        `[^>]*(?:/>|>\\s*</Relationship>)`,
      "gi",
    ),
    "",
  );

const stripExternalLinkContentTypes = (xml: string) =>
  xml.replace(
    new RegExp(
      `<Override\\b(?=[^>]*\\bContentType="${escapeRegExp(EXTERNAL_LINK_CONTENT_TYPE)}")` +
        `[^>]*/>`,
      "gi",
    ),
    "",
  );

const stripCalcChainContentTypes = (xml: string) =>
  xml.replace(
    new RegExp(
      `<Override\\b(?=[^>]*\\bContentType="${escapeRegExp(CALC_CHAIN_CONTENT_TYPE)}")` +
        `[^>]*/>`,
      "gi",
    ),
    "",
  );

const ensureFullCalcOnLoad = (xml: string) => {
  if (/<calcPr\b/i.test(xml)) {
    return xml.replace(/<calcPr\b([^>]*)\/>/i, (_match, attrs) => {
      let nextAttrs = String(attrs);
      if (!/\bcalcMode=/i.test(nextAttrs)) nextAttrs += ' calcMode="auto"';
      if (!/\bfullCalcOnLoad=/i.test(nextAttrs)) nextAttrs += ' fullCalcOnLoad="1"';
      if (!/\bforceFullCalc=/i.test(nextAttrs)) nextAttrs += ' forceFullCalc="1"';
      return `<calcPr${nextAttrs}/>`;
    });
  }
  return xml.replace(
    /<\/workbook>\s*$/i,
    '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>',
  );
};

/** Після точкового запису клітинок старий calcChain дає Recovery в Excel. */
export async function stripStaleCalcChainFromZip(zip: JSZip) {
  if (zip.file("xl/calcChain.xml")) zip.remove("xl/calcChain.xml");
  // Незадекларована частина в xl/ — Excel відкриває файл через Repair.
  if (zip.file("xl/ejoosAppliedChanges.json")) {
    zip.remove("xl/ejoosAppliedChanges.json");
  }
  const rels = zip.file("xl/_rels/workbook.xml.rels");
  if (rels) {
    const xml = await rels.async("string");
    const next = stripWorkbookCalcChainRelationship(xml);
    if (next !== xml) zip.file("xl/_rels/workbook.xml.rels", next);
  }
  const types = zip.file("[Content_Types].xml");
  if (types) {
    const xml = await types.async("string");
    const next = stripCalcChainContentTypes(xml);
    if (next !== xml) zip.file("[Content_Types].xml", next);
  }
  const workbook = zip.file("xl/workbook.xml");
  if (workbook) {
    const xml = await workbook.async("string");
    const next = ensureFullCalcOnLoad(xml);
    if (next !== xml) zip.file("xl/workbook.xml", next);
  }
}

const stripStaleExternalValidationExts = (xml: string) =>
  xml.replace(/<ext\b[^>]*>[\s\S]*?<\/ext>/gi, (block) => {
    if (!/dataValidation/i.test(block)) return block;
    if (!/\[[^\]]+\.xlsx\]|#REF!/i.test(block)) return block;
    return "";
  });

const readDxfCount = (stylesXml: string) => {
  const attrCount = Number(stylesXml.match(/<dxfs\b[^>]*\bcount="(\d+)"/i)?.[1]);
  if (Number.isFinite(attrCount) && attrCount >= 0) return attrCount;
  return (stylesXml.match(/<dxf\b/gi) ?? []).length;
};

const stripDxfs = (stylesXml: string) =>
  stylesXml.replace(/<dxfs\b[^>]*>[\s\S]*?<\/dxfs>/i, '<dxfs count="0"/>');

/**
 * xlsx-populate повертає системний колір меж як текст і потім записує його в
 * rgb, хоча OOXML дозволяє там лише 8-значний hex. Excel відновлює styles.xml.
 */
const normalizeSystemStyleColors = (stylesXml: string) =>
  stylesXml.replace(
    /\brgb="SYSTEM FOREGROUND"/gi,
    'indexed="64"',
  );

/**
 * xlsx-populate інколи серіалізує колекції styles.xml без обов'язкового
 * атрибута count. XML лишається валідним, але Excel запускає Recovery.
 */
const normalizeStyleCollectionCounts = (stylesXml: string) => {
  const collections: Array<[string, string]> = [
    ["numFmts", "numFmt"],
    ["fonts", "font"],
    ["fills", "fill"],
    ["borders", "border"],
    ["cellStyleXfs", "xf"],
    ["cellXfs", "xf"],
    ["cellStyles", "cellStyle"],
    ["dxfs", "dxf"],
  ];
  let next = stylesXml;
  for (const [collection, child] of collections) {
    const blockRe = new RegExp(
      `<${collection}\\b([^>]*)>([\\s\\S]*?)<\\/${collection}>`,
      "i",
    );
    next = next.replace(blockRe, (_block, rawAttrs, body) => {
      const count = (
        String(body).match(new RegExp(`<${child}\\b`, "gi")) ?? []
      ).length;
      const attrs = String(rawAttrs).replace(/\s+count="[^"]*"/i, "");
      return `<${collection}${attrs} count="${count}">${body}</${collection}>`;
    });
  }
  return next;
};

const stripConditionalFormatting = (sheetXml: string) =>
  sheetXml.replace(
    /<conditionalFormatting\b[^>]*>[\s\S]*?<\/conditionalFormatting>/gi,
    "",
  );

const normalizeInternalFormulaWorkbookRefs = (sheetXml: string) =>
  sheetXml.replace(
    /(<f\b[^>]*>[\s\S]*?<\/f>)/gi,
    (formulaBlock) =>
      formulaBlock.replace(/'\[\d+\]([^']+)'!/g, (_match, sheetName) => {
        return `'${sheetName}'!`;
      }),
  );

/** xlsx-populate пише `<v>NaN</v>` у порожні числові клітинки — Excel це лагодить. */
export const stripInvalidNumericCellValues = (sheetXml: string) =>
  sheetXml.replace(/<v>\s*(?:NaN|[+-]?Infinity)\s*<\/v>/gi, "<v></v>");

/** `<c r="F53"/ s="202">` — слеш самозакритої клітинки всередині тега. */
export const repairBrokenCellOpenTags = (sheetXml: string) =>
  sheetXml.replace(
    /<c\b([^>]*?)\/(\s+[^>]+)>/gi,
    (_all, before, after) => `<c${before}${after}>`,
  );

const VALID_CELL_REF = /\br="[A-Z]{1,3}\d+"(?![0-9A-Za-z])/i;

/** xlsx-populate падає на `<c>` без r="A1" — Cannot read properties of undefined (reading 'columnNumber'). */
export const stripCellsWithoutValidReference = (sheetXml: string) =>
  sheetXml.replace(/<c\b[^<>]*?(?:\/>|>[\s\S]*?<\/c>)/gi, (cellXml) =>
    VALID_CELL_REF.test(cellXml) ? cellXml : "",
  );

const stripOrphanSharedFormulas = (sheetXml: string) => {
  const masterSharedFormulaIds = new Set<string>();
  for (const formula of sheetXml.matchAll(
    /<f\b((?:(?!\/>)[^<>])*)>([\s\S]*?)<\/f>/gi,
  )) {
    const attrs = formula[1];
    if (!/\bt="shared"/i.test(attrs)) continue;
    const si = parseAttr(attrs, "si");
    if (si && formula[2].trim()) masterSharedFormulaIds.add(si);
  }

  return sheetXml.replace(
    /<f\b([^<>]*?\bt="shared"[^<>]*?)(?:\/>|>[\s\S]*?<\/f>)/gi,
    (formulaXml, attrs) => {
      const si = parseAttr(String(attrs), "si");
      if (si && !masterSharedFormulaIds.has(si)) return "";
      return formulaXml;
    },
  );
};

const countSharedStringRefs = async (zip: JSZip) => {
  let refs = 0;
  for (const path of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/i.test(path)) continue;
    const xml = await zip.file(path)?.async("string");
    if (!xml) continue;
    refs += (xml.match(/<c\b[^<>]*?\bt="s"/gi) ?? []).length;
  }
  return refs;
};

const normalizeSharedStringsCounts = (
  sharedStringsXml: string,
  count: number,
) => {
  const uniqueCount = (sharedStringsXml.match(/<si>/g) ?? []).length;
  return sharedStringsXml.replace(/<sst\b([^>]*)>/i, (_match, attrs) => {
    const nextAttrs = String(attrs)
      .replace(/\s+count="[^"]*"/i, "")
      .replace(/\s+uniqueCount="[^"]*"/i, "");
    return `<sst${nextAttrs} count="${count}" uniqueCount="${uniqueCount}">`;
  });
};

const resolveSheetPath = async (zip: JSZip, sheetName: string) => {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!workbookXml || !relsXml) return null;

  const ridToTarget = new Map<string, string>();
  for (const rel of relsXml.matchAll(
    /\bId="([^"]+)"[^>]*\bTarget="([^"]+)"/g,
  )) {
    ridToTarget.set(rel[1], rel[2]);
  }
  for (const rel of relsXml.matchAll(
    /\bTarget="([^"]+)"[^>]*\bId="([^"]+)"/g,
  )) {
    if (!ridToTarget.has(rel[2])) ridToTarget.set(rel[2], rel[1]);
  }

  for (const sheet of workbookXml.matchAll(/<sheet\b([^>]*)>/gi)) {
    const attrs = sheet[1];
    if (parseAttr(attrs, "name") !== sheetName) continue;
    const rid = attrs.match(/\br:id="([^"]+)"/)?.[1] ?? parseAttr(attrs, "id");
    if (!rid) continue;
    const target = ridToTarget.get(rid);
    if (!target) continue;
    return target.startsWith("/")
      ? target.slice(1)
      : target.startsWith("xl/")
        ? target
        : `xl/${target.replace(/^\.\//, "")}`;
  }
  return null;
};

const cellColumnNumber = (cellXml: string) => {
  const ref = cellXml.match(/<c\b[^<>]*?\br="([A-Z]+)[0-9]+"/i)?.[1];
  return ref ? columnLetterToNumber(ref) : null;
};

const cellReferenceParts = (cellXml: string) => {
  const ref = cellXml.match(/<c\b[^<>]*?\br="([A-Z]+)([0-9]+)"/i);
  if (!ref) return null;
  return {
    row: Number(ref[2]),
    columnNumber: columnLetterToNumber(ref[1]),
  };
};

const cellHasValue = (cellXml: string) => /<(?:v|f|is)\b/i.test(cellXml);
const cellHasContent = cellHasValue;

const updateRowSpans = (openTag: string, cells: string[]) => {
  const columns = cells
    .map((cellXml) => cellReferenceParts(cellXml)?.columnNumber ?? 0)
    .filter(Boolean);
  if (!columns.length) return openTag.replace(/\s+spans="[^"]*"/i, "");
  const spans = `spans="${Math.min(...columns)}:${Math.max(...columns)}"`;
  if (/\bspans="[^"]*"/i.test(openTag)) {
    return openTag.replace(/\bspans="[^"]*"/i, spans);
  }
  return openTag.replace(/>$/, ` ${spans}>`);
};

const insertRowXml = (sheetXml: string, rowNumber: number, rowXml: string) => {
  for (const row of sheetXml.matchAll(
    /<row\b[^>]*\br="([0-9]+)"[^>]*(?:\/>|>[\s\S]*?<\/row>)/gi,
  )) {
    const existingRowNumber = Number(row[1]);
    if (existingRowNumber > rowNumber) {
      return `${sheetXml.slice(0, row.index)}${rowXml}${sheetXml.slice(row.index)}`;
    }
  }

  const sheetDataClose = sheetXml.search(/<\/sheetData>/i);
  if (sheetDataClose >= 0) {
    return `${sheetXml.slice(0, sheetDataClose)}${rowXml}${sheetXml.slice(sheetDataClose)}`;
  }
  return sheetXml;
};

const normalizeWorksheetRows = (sheetXml: string) => {
  const rowCells = new Map<number, Map<number, string>>();
  for (const cell of sheetXml.matchAll(/<c\b[^<>]*?(?:\/>|>[\s\S]*?<\/c>)/gi)) {
    const cellXml = cell[0];
    const ref = cellReferenceParts(cellXml);
    if (!ref) continue;
    if (ref.columnNumber > 256 && !cellHasContent(cellXml)) continue;
    const cells = rowCells.get(ref.row) ?? new Map<number, string>();
    cells.set(ref.columnNumber, cellXml);
    rowCells.set(ref.row, cells);
  }

  const existingRows = new Set<number>();
  let normalized = sheetXml.replace(
    /<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/gi,
    (rowXml) => {
      const openMatch = rowXml.match(/^<row\b[^>]*>/i);
      if (!openMatch) return rowXml;
      const rowNumber = Number(parseAttr(openMatch[0], "r") || 0);
      if (!rowNumber) return rowXml;
      existingRows.add(rowNumber);

      const isSelfClosing = /\/\s*>$/.test(openMatch[0]);
      if (isSelfClosing) {
        const cells = [...(rowCells.get(rowNumber)?.entries() ?? [])]
          .sort(([a], [b]) => a - b)
          .map(([, cellXml]) => cellXml);
        return `${updateRowSpans(openMatch[0].replace(/\s*\/\s*>$/, ">"), cells)}${cells.join("")}</row>`;
      }
      const cells = [...(rowCells.get(rowNumber)?.entries() ?? [])]
        .sort(([a], [b]) => a - b)
        .map(([, cellXml]) => cellXml);
      return `${updateRowSpans(openMatch[0], cells)}${cells.join("")}</row>`;
    },
  );

  [...rowCells.keys()]
    .filter((rowNumber) => !existingRows.has(rowNumber))
    .sort((a, b) => a - b)
    .forEach((rowNumber) => {
      const cells = [...(rowCells.get(rowNumber)?.entries() ?? [])]
        .sort(([a], [b]) => a - b)
        .map(([, cellXml]) => cellXml);
      normalized = insertRowXml(
        normalized,
        rowNumber,
        `<row r="${rowNumber}" ${updateRowSpans(">", cells).slice(1)}${cells.join("")}</row>`,
      );
    });

  return normalized;
};

const stripShpoNameOnlyRows = (sheetXml: string) =>
  sheetXml.replace(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/gi, (rowXml) => {
    const cells = [...rowXml.matchAll(/<c\b[^<>]*?(?:\/>|>[\s\S]*?<\/c>)/gi)].map(
      (match) => match[0],
    );
    if (cells.length === 0) return rowXml;

    const valueCellsInPersonArea = cells
      .map((cellXml) => ({ cellXml, column: cellColumnNumber(cellXml) }))
      .filter(({ cellXml, column }) => {
        return column != null && column >= 1 && column <= 18 && cellHasValue(cellXml);
      });

    const hasNameOnly =
      valueCellsInPersonArea.length > 0 &&
      valueCellsInPersonArea.every(({ column }) => column === 7);
    if (!hasNameOnly) return rowXml;

    const cellsOutsideName = cells.filter((cellXml) => {
      const column = cellColumnNumber(cellXml);
      return column !== 7;
    });
    if (cellsOutsideName.length === 0) return "";

    let nextRow = rowXml;
    for (const { cellXml } of valueCellsInPersonArea) {
      nextRow = nextRow.replace(cellXml, "");
    }
    return nextRow;
  });

/**
 * Старі шаблони ЕЖООС часто тягнуть зовнішні книги з локальних шляхів
 * попередніх операторів. Excel for Mac відкриває такі файли через recovery.
 * Для канонічного журналу нам потрібні значення в самій книзі, не stale links.
 */
export async function sanitizeEjoosWorkbookBlob(input: Blob | File): Promise<Blob> {
  const zip = await JSZip.loadAsync(await input.arrayBuffer());
  let changed = false;
  let stripBloatedConditionalFormatting = false;

  const removePath = (path: string) => {
    if (zip.file(path)) {
      zip.remove(path);
      changed = true;
    }
  };

  for (const path of Object.keys(zip.files)) {
    if (path.startsWith("xl/externalLinks/")) removePath(path);
  }
  removePath("xl/calcChain.xml");
  removePath("xl/ejoosAppliedChanges.json");

  const patchXml = async (path: string, patch: (xml: string) => string) => {
    const file = zip.file(path);
    if (!file) return;
    const before = await file.async("string");
    const after = patch(before);
    if (after !== before) {
      zip.file(path, after);
      changed = true;
    }
  };

  await patchXml("xl/workbook.xml", stripWorkbookExternalReferences);
  await patchXml("xl/_rels/workbook.xml.rels", (xml) =>
    stripWorkbookCalcChainRelationship(stripWorkbookExternalRelationships(xml)),
  );
  await patchXml("[Content_Types].xml", (xml) =>
    stripCalcChainContentTypes(stripExternalLinkContentTypes(xml)),
  );
  const shpoSheetPath = await resolveSheetPath(zip, "1. ШПО");

  const stylesFile = zip.file("xl/styles.xml");
  if (stylesFile) {
    const stylesXml = await stylesFile.async("string");
    stripBloatedConditionalFormatting = readDxfCount(stylesXml) > 60_000;
    let nextStyles = normalizeStyleCollectionCounts(
      normalizeSystemStyleColors(stylesXml),
    );
    if (stripBloatedConditionalFormatting) {
      nextStyles = stripDxfs(nextStyles);
    }
    if (nextStyles !== stylesXml) {
      zip.file("xl/styles.xml", nextStyles);
      changed = true;
    }
  }

  for (const path of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/i.test(path)) continue;
    const file = zip.file(path);
    if (!file) continue;
    const xml = await file.async("string");
    let next = xml;
    next = repairBrokenCellOpenTags(next);
    next = stripInvalidNumericCellValues(next);
    next = normalizeWorksheetRows(next);
    next = normalizeInternalFormulaWorkbookRefs(next);
    next = expandSharedFormulas(next);
    next = stripOrphanSharedFormulas(next);
    if (hasStaleExternalWorkbookRef(next)) {
      next = stripStaleExternalValidationExts(next);
    }
    if (stripBloatedConditionalFormatting) {
      next = stripConditionalFormatting(next);
    }
    if (path === shpoSheetPath) {
      next = stripShpoNameOnlyRows(next);
    }
    if (next !== xml) {
      zip.file(path, next);
      changed = true;
    }
  }

  const sharedStringsFile = zip.file("xl/sharedStrings.xml");
  if (sharedStringsFile) {
    const sharedStringsXml = await sharedStringsFile.async("string");
    const nextSharedStringsXml = normalizeSharedStringsCounts(
      sharedStringsXml,
      await countSharedStringRefs(zip),
    );
    if (nextSharedStringsXml !== sharedStringsXml) {
      zip.file("xl/sharedStrings.xml", nextSharedStringsXml);
      changed = true;
    }
  }

  if (!changed) return input;
  return zip.generateAsync({
    type: "blob",
    mimeType: XLSX_MIME,
    compression: "DEFLATE",
  });
}

/**
 * Повертає styles.xml (+ theme) з робочої копії в поточний файл.
 * Дані аркушів не чіпаємо — лише оформлення, яке Excel зняв після битого styles.xml.
 */
export async function graftWorkbookStyles(
  target: Blob | File,
  styleSource: Blob | File,
): Promise<Blob> {
  const dest = await JSZip.loadAsync(await target.arrayBuffer());
  const src = await JSZip.loadAsync(await styleSource.arrayBuffer());
  const styles = src.file("xl/styles.xml");
  if (!styles) {
    throw new Error("У резервній версії немає xl/styles.xml");
  }
  dest.file("xl/styles.xml", await styles.async("uint8array"));
  for (const path of Object.keys(src.files)) {
    if (!/^xl\/theme\//i.test(path) || src.files[path]?.dir) continue;
    const part = src.file(path);
    if (part) dest.file(path, await part.async("uint8array"));
  }
  return dest.generateAsync({
    type: "blob",
    mimeType: XLSX_MIME,
    compression: "DEFLATE",
  });
}
