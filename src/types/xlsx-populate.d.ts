declare module 'xlsx-populate/browser/xlsx-populate-no-encryption' {
  type CellValue = string | number | boolean | Date | null | undefined | unknown

  interface Cell {
    value(): CellValue
    value(value: CellValue): Cell
    style(name: string): CellValue
    style(name: string, value: CellValue): Cell
  }

  interface Range {
    value(): CellValue[][]
    style(name: string, value: CellValue): Range
  }

  interface Sheet {
    name(): string
    name(value: string): Sheet
    cell(rowNumber: number, columnNumber: number): Cell
    range(startCell: string, endCell: string): Range
    usedRange(): Range | undefined
  }

  interface Workbook {
    sheet(index: number): Sheet
    sheets(): Sheet[]
    addSheet(name: string): Sheet
    outputAsync(type?: 'blob' | 'arraybuffer' | 'base64' | 'uint8array'): Promise<Blob>
  }

  interface XlsxPopulateStatic {
    MIME_TYPE: string
    fromDataAsync(data: Blob | ArrayBuffer): Promise<Workbook>
    fromBlankAsync(): Promise<Workbook>
  }

  const XlsxPopulate: XlsxPopulateStatic
  export default XlsxPopulate
}
