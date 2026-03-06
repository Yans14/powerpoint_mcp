import { describe, expect, it } from "vitest";

import { measurementToEmu } from "../../src/bridge/units.js";

describe("measurementToEmu", () => {
  it("converts inches", () => {
    expect(measurementToEmu("1in")).toBe(914400);
  });

  it("converts points", () => {
    expect(measurementToEmu("24pt")).toBe(304800);
  });

  it("converts centimeters", () => {
    expect(measurementToEmu("2cm")).toBe(720000);
  });

  it("passes through numeric EMU", () => {
    expect(measurementToEmu(12345)).toBe(12345);
  });

  it("rejects malformed strings", () => {
    expect(() => measurementToEmu("foo")).toThrow(/Invalid measurement/);
  });
});
