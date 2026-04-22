import { describe, expect, it } from "vitest";
import { cmToPt, parseLengthToPt } from "../lib/units";

describe("unit helpers", () => {
  it("converts centimeters to points", () => {
    expect(cmToPt(1)).toBeCloseTo(28.34646);
  });

  it("parses point and centimeter lengths", () => {
    expect(parseLengthToPt("20")).toBe(20);
    expect(parseLengthToPt("2cm")).toBeCloseTo(56.69292);
    expect(parseLengthToPt("0cm")).toBeUndefined();
    expect(parseLengthToPt("")).toBeUndefined();
  });
});
