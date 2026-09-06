import { describe, expect, it } from "vitest";
import type { BackendPersonDocument } from "../../api";
import {
  extractPersonSignature,
  PERSON_SIGNATURE_DOCUMENT_TYPE,
} from "./personSignatureStore";

const image = "data:image/png;base64,AA==";

const document = (
  id: string,
  personExternalId: string,
  type: string,
  fields: Record<string, unknown>,
) =>
  ({
    id,
    personExternalId,
    type,
    fields,
  }) as BackendPersonDocument;

describe("extractPersonSignature", () => {
  it("never takes a signature from another person's documents", () => {
    const documents = [
      document("galamaga-sign", "person-a", PERSON_SIGNATURE_DOCUMENT_TYPE, {
        signatureData: image,
        signatureFileName: "ГАЛАМАГА_підпис.png",
      }),
      document("yakushenko-form", "person-b", "form12Report", {
        fullName: "ЯКУШЕНОК Олександр Володимирович",
      }),
    ];

    expect(extractPersonSignature(documents, "person-b").signature).toBeNull();
  });

  it("rejects an obviously foreign named signature already attached to the person", () => {
    const documents = [
      document("wrong-sign", "person-b", PERSON_SIGNATURE_DOCUMENT_TYPE, {
        signatureData: image,
        signatureFileName: "ГАЛАМАГА_підпис.png",
      }),
      document("yakushenko-form", "person-b", "form12Report", {
        fullName: "ЯКУШЕНОК Олександр Володимирович",
      }),
    ];

    expect(extractPersonSignature(documents, "person-b").signature).toBeNull();
  });

  it("uses the signature named for the selected person", () => {
    const documents = [
      document("yakushenko-sign", "person-b", PERSON_SIGNATURE_DOCUMENT_TYPE, {
        signatureData: image,
        signatureFileName: "ЯКУШЕНОК_підпис.png",
      }),
      document("yakushenko-form", "person-b", "form12Report", {
        fullName: "ЯКУШЕНОК Олександр Володимирович",
      }),
    ];

    expect(extractPersonSignature(documents, "person-b").signature).toEqual({
      signatureData: image,
      signatureFileName: "ЯКУШЕНОК_підпис.png",
    });
  });
});
