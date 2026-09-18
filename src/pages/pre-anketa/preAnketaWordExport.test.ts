import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import JSZip from "jszip";
import { createEmptyPreAnketaForm } from "./preAnketaFields";
import { fillPreAnketaDocumentXml } from "./preAnketaWordExport";

const templatePath = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../public/templates/pre-anketa-form.docx",
);

describe("preAnketaWordExport", () => {
  it("fills template paragraphs for key fields", async () => {
    const zip = await JSZip.loadAsync(readFileSync(templatePath));
    const documentXml = await zip.file("word/document.xml")?.async("string");
    expect(documentXml).toBeTruthy();

    const fields = {
      ...createEmptyPreAnketaForm(),
      callsign: "СІМБА",
      fullName: "СІБАРЦЕВ ДМИТРО МИХАЙЛОВИЧ",
      rank: "солдат",
      rnokpp: "1234567890",
      passport: "Паспорт громадянина України КВ 001828",
      phone1: "+380501234567",
      serviceMobilized: "1",
    };

    const filled = fillPreAnketaDocumentXml(documentXml!, fields);
    expect(filled).toContain("СІМБА");
    expect(filled).toContain("СІБАРЦЕВ ДМИТРО МИХАЙЛОВИЧ");
    expect(filled).toContain("1234567890");
    expect(filled).toContain("Мобілізований- ☑");
  });

  it("writes relatives into the right-hand value column, not the hint column", async () => {
    const zip = await JSZip.loadAsync(readFileSync(templatePath));
    const documentXml = await zip.file("word/document.xml")?.async("string");
    expect(documentXml).toBeTruthy();

    const fields = {
      ...createEmptyPreAnketaForm(),
      mother: "померла",
      father: "помер",
      trustedPerson:
        "син - Барціковський Даніїл В'ячеславович, 17.04.2004, тел: 0660284874",
      children: "син - Барціковський Даніїл В'ячеславович, 17.04.2004",
    };

    const filled = fillPreAnketaDocumentXml(documentXml!, fields);
    expect(filled).toContain(
      "Мати:</w:t></w:r></w:p><w:p",
    );
    expect(filled).toContain(
      "(П.І Б., рік народження, номер телефону, адреса проживання)",
    );
    expect(filled).toContain("померла");
    expect(filled).toContain("помер");
    expect(filled).toContain("0660284874");
    expect(filled).not.toContain(
      "(П.І Б., рік народження, номер телефону, адреса проживання)</w:t></w:r></w:p><w:p",
    );
  });
});
