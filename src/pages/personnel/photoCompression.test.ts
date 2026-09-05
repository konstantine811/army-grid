import { describe, expect, it } from "vitest";
import {
  estimateDataUrlBytes,
  fitPhotoDimensions,
} from "./photoCompression";

describe("photoCompression", () => {
  it("fits a large portrait into the preview limit", () => {
    expect(fitPhotoDimensions(2400, 3200)).toEqual({
      width: 480,
      height: 640,
    });
  });

  it("preserves aspect ratio and does not upscale small photos", () => {
    expect(fitPhotoDimensions(300, 400)).toEqual({
      width: 300,
      height: 400,
    });
    expect(fitPhotoDimensions(2000, 1000)).toEqual({
      width: 480,
      height: 240,
    });
  });

  it("estimates base64 payload size", () => {
    expect(estimateDataUrlBytes("data:image/jpeg;base64,AAAA")).toBe(3);
  });
});
