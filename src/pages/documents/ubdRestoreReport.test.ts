import { describe, expect, it } from "vitest";
import {
  buildUbdRestoreFolderName,
  resolveUbdRestoreFolderName,
  capitalizeSignerTitleBlock,
  formatPositionTitleBlock,
  resolveCoveringSignerParts,
  resolveApproverParts,
  resolveUbdRestoreSignerTitle,
  type UbdRestoreReportFields,
} from "./ubdRestoreReport";

describe("ubdRestoreReport position title", () => {
  it("prefixes exported folder name with 1ПБ", () => {
    expect(buildUbdRestoreFolderName("НОВІКОВ Олег")).toBe(
      "1ПБ Рапорт на відновлення УБД · НОВІКОВ Олег",
    );
  });

  it("upgrades legacy folder name without 1ПБ prefix", () => {
    expect(
      resolveUbdRestoreFolderName(
        "БАРАНОВ Ігор",
        "Рапорт на відновлення УБД · БАРАНОВ Ігор",
      ),
    ).toBe("1ПБ Рапорт на відновлення УБД · БАРАНОВ Ігор");
  });

  it("capitalizes the first line of the signer position block", () => {
    expect(
      formatPositionTitleBlock(
        "водій мінометного взводу 1 піхотного батальйону військової частини А4862",
      ),
    ).toBe(
      "Водій мінометного взводу 1 піхотного батальйону\nвійськової частини А4862",
    );
  });

  it("capitalizes a manually entered multiline signer title", () => {
    expect(
      capitalizeSignerTitleBlock(
        "водій мінометного взводу 1 піхотного батальйону\nвійськової частини А4862",
      ),
    ).toBe(
      "Водій мінометного взводу 1 піхотного батальйону\nвійськової частини А4862",
    );
  });

  it("prefers staffPosition when building export signer title", () => {
    expect(
      resolveUbdRestoreSignerTitle({
        staffPosition:
          "водій мінометного взводу 1 піхотного батальйону військової частини А4862",
        signerTitle: "застарілий рядок",
      }),
    ).toBe(
      "Водій мінометного взводу 1 піхотного батальйону\nвійськової частини А4862",
    );
  });

  it("builds covering commander block from configured signatory", () => {
    const parts = resolveCoveringSignerParts({
      date: "29.08.2026",
      staffPosition: "",
      signerTitle: "",
      signatories: [
        {
          blockType: "SIGNER",
          title: "Командир 1 піхотного батальйону\nвійськової частини А4862",
          rank: "старший лейтенант",
          fullName: "Єгор СИДОРЕНКО",
          signatureData: "data:image/png;base64,abc",
        },
      ],
    } as UbdRestoreReportFields);

    expect(parts.titleLines).toEqual([
      "Командир 1 піхотного батальйону",
      "військової частини А4862",
    ]);
    expect(parts.rank).toBe("старший лейтенант");
    expect(parts.fullName).toBe("Єгор СИДОРЕНКО");
    expect(parts.date).toBe("29.08.2026");
    expect(parts.signatureData).toBe("data:image/png;base64,abc");
  });

  it("does not duplicate approver rank in title lines", () => {
    const parts = resolveApproverParts({
      date: "29.08.2026",
      staffPosition: "",
      signerTitle: "",
      signatories: [
        {
          blockType: "APPROVAL",
          title: "Командир військової частини А4862\nстарший лейтенант",
          rank: "старший лейтенант",
          fullName: "СЕМЕНЮК",
          signatureData: null,
        },
      ],
    } as UbdRestoreReportFields);

    expect(parts.titleLines).toEqual([
      "Командир військової частини А4862",
      "",
    ]);
    expect(parts.rank).toBe("старший лейтенант");
    expect(parts.fullName).toBe("СЕМЕНЮК");
  });

  it("drops duplicate commander unit line in approver title", () => {
    const parts = resolveApproverParts({
      date: "29.08.2026",
      staffPosition: "",
      signerTitle: "",
      signatories: [
        {
          blockType: "APPROVAL",
          title:
            "Командир військової частини А4862\nкомандира військової частини А4862",
          rank: "старший лейтенант",
          fullName: "СЕМЕНЮК",
          signatureData: null,
        },
      ],
    } as UbdRestoreReportFields);

    expect(parts.titleLines).toEqual([
      "Командир військової частини А4862",
      "",
    ]);
  });
});
