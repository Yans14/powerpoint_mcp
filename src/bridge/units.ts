const EMU_PER_INCH = 914400;
const EMU_PER_POINT = 12700;
const EMU_PER_CM = 360000;
const EMU_PER_PIXEL = 9525;

export type MeasurementInput = number | string;

const MEASUREMENT_RE = /^(\d+(?:\.\d+)?)(in|pt|cm|px)$/i;

export function measurementToEmu(value: MeasurementInput): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.round(value);
  }

  if (typeof value !== "string") {
    throw new Error("Measurement must be a number (EMU) or string with unit");
  }

  const input = value.trim().toLowerCase();
  const match = input.match(MEASUREMENT_RE);
  if (!match) {
    throw new Error(`Invalid measurement '${value}'. Expected format like 2in, 24pt, 5cm, or 96px.`);
  }

  const numeric = Number(match[1]);
  const unit = match[2];

  switch (unit) {
    case "in":
      return Math.round(numeric * EMU_PER_INCH);
    case "pt":
      return Math.round(numeric * EMU_PER_POINT);
    case "cm":
      return Math.round(numeric * EMU_PER_CM);
    case "px":
      return Math.round(numeric * EMU_PER_PIXEL);
    default:
      throw new Error(`Unsupported measurement unit '${unit}'.`);
  }
}

export function emuToInches(emu: number): number {
  return emu / EMU_PER_INCH;
}
