import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { styleEjoosExport } from "./ejoosExportStyle";

describe("EJOOS export formatting", () => {
  it("formats only the four requested sheets and preserves all workbook content and structure", async () => {
    const module = await import("xlsx-populate/browser/xlsx-populate-no-encryption");
    const wb: any = await module.default.fromBlankAsync();
    const names = ["1. ШПО", "2. ООС", "3. Виключені", "6. Табель", "10. Історія змін", "Примітки"];
    for (const [index, name] of names.entries()) {
      const sheet = index ? wb.addSheet(name) : wb.sheet(0).name(name);
      sheet.range("A1:C4").style({ fontFamily: "Arial", fontSize: 18, bold: true, horizontalAlignment: "left", verticalAlignment: "top", wrapText: true, fill: "FFF2CC", fontColor: "FF0000" });
      sheet.cell("A1").value("Заголовок");
      sheet.range("A1:C1").merged(true);
      sheet.cell("A2").value("Тестовий запис");
      sheet.cell("B2").value(12.5).style("numberFormat", "0.00");
      sheet.cell("C2").formula("B2*2");
      sheet.column("A").width(24);
      sheet.row(2).height(35);
      sheet.definedName("_xlnm.Print_Area", sheet.range("A1:C4"));
    }
    const original = await wb.outputAsync("blob") as Blob;
    const styled = await styleEjoosExport(original);
    const before = await JSZip.loadAsync(await original.arrayBuffer());
    const after = await JSZip.loadAsync(await styled.arrayBuffer());
    expect(Object.keys(after.files).sort()).toEqual(Object.keys(before.files).sort());
    for (const path of Object.keys(before.files).filter(path => !before.files[path].dir)) {
      if (path === "xl/styles.xml") continue;
      const oldXml = await before.file(path)!.async("string");
      const newXml = await after.file(path)!.async("string");
      if (/xl\/worksheets\/sheet[1-4]\.xml$/.test(path)) {
        const withoutStyleRefs = (xml: string) => xml.replace(/<(c|row|col)\b[^>]*>/g, tag => tag.replace(/\s(?:s|style)="[^"]*"/g, ""));
        expect(withoutStyleRefs(newXml)).toBe(withoutStyleRefs(oldXml));
      } else expect(newXml).toBe(oldXml);
    }
    const loaded: any = await module.default.fromDataAsync(await styled.arrayBuffer());
    for (const [index, name] of names.entries()) {
      const sheet = loaded.sheet(name);
      const wide = index === 1 || index === 2;
      for (const ref of ["A1", "A2", "B2", "C2", "C4"]) {
        const cell = sheet.cell(ref);
        if (index >= 4) {
          expect(cell.style("fontFamily")).toBe("Arial");
          expect(cell.style("bold")).toBe(true);
          expect(cell.style("horizontalAlignment")).toBe("left");
          expect(cell.style("fontSize")).toBe(18);
          continue;
        }
        expect(cell.style("fontFamily")).toBe("Times New Roman");
        expect(cell.style("bold")).toBe(false);
        expect(cell.style("horizontalAlignment")).toBe("center");
        expect(cell.style("verticalAlignment")).toBe("center");
        expect(cell.style("fontSize")).toBe(wide ? 14 : index === 4 ? 18 : 12);
        expect(cell.style("wrapText")).toBe(wide || index === 4);
        if (wide) for (const edge of ["leftBorder", "rightBorder", "topBorder", "bottomBorder"]) expect(cell.style(edge).style).toBe("thin");
        expect(cell.style("fontColor")).toMatchObject({ rgb: "FF0000" });
      }
      expect(sheet.cell("B2").style("numberFormat")).toBe("0.00");
      expect(sheet.cell("C2").formula()).toBe("B2*2");
      expect(sheet.range("A1:C1").merged()).toBe(true);
    }
    const repeated = await JSZip.loadAsync(await (await styleEjoosExport(styled)).arrayBuffer());
    expect(await repeated.file("xl/styles.xml")!.async("string")).toBe(await after.file("xl/styles.xml")!.async("string"));
  });

  it("forces ООС РНОКПП to General instead of Custom @", async () => {
    const module = await import("xlsx-populate/browser/xlsx-populate-no-encryption");
    const wb: any = await module.default.fromBlankAsync();
    const sheet = wb.sheet(0).name("2. ООС");
    sheet.cell("V7").value(3142223156).style({
      numberFormat: "@",
      horizontalAlignment: "center",
      verticalAlignment: "center",
      wrapText: true,
    });
    sheet.cell("C7").value("21643").style({
      numberFormat: "@",
      horizontalAlignment: "center",
      verticalAlignment: "center",
    });
    const original = await wb.outputAsync("blob") as Blob;
    const styled = await styleEjoosExport(original);
    const loaded: any = await module.default.fromDataAsync(await styled.arrayBuffer());
    const oos = loaded.sheet("2. ООС");
    const rnokppFormat = String(oos.cell("V7").style("numberFormat") ?? "");
    expect(rnokppFormat === "General" || rnokppFormat === "0" || rnokppFormat === "").toBe(true);
    expect(oos.cell("C7").style("numberFormat")).toBe("@");
  });
});
