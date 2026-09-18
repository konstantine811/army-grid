import { describe, expect, it } from "vitest";
import {
  isFatalOcrError,
  isSkippableOcrPageError,
} from "./useAnketaQuestionnaireOcr";

const httpError = (status: number, message: string) =>
  Object.assign(new Error(message), { status });

describe("OCR page error handling", () => {
  it("treats 400/413/500 as skippable so other pages can still return fields", () => {
    expect(isSkippableOcrPageError(httpError(400, "Bad Request"))).toBe(true);
    expect(isSkippableOcrPageError(httpError(413, "Payload Too Large"))).toBe(
      true,
    );
    expect(isSkippableOcrPageError(httpError(500, "Gemini timeout"))).toBe(true);
    expect(isFatalOcrError(httpError(400, "Bad Request"))).toBe(false);
  });

  it("stops OCR on auth errors", () => {
    expect(isFatalOcrError(httpError(401, "Unauthorized"))).toBe(true);
    expect(isSkippableOcrPageError(httpError(401, "Unauthorized"))).toBe(false);
    expect(isFatalOcrError(httpError(403, "Forbidden"))).toBe(true);
  });
});
