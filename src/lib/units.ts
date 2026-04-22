const POINTS_PER_CM = 28.34646;

export function cmToPt(value: number): number {
  return value * POINTS_PER_CM;
}

export function parseLengthToPt(value: string): number | undefined {
  const text = value.trim();
  if (!text) {
    return undefined;
  }

  const number = Number.parseFloat(text.replace(",", "."));
  if (!Number.isFinite(number) || number <= 0) {
    return undefined;
  }

  if (/cm/i.test(text)) {
    return cmToPt(number);
  }

  return number;
}

export function clampPositiveInteger(value: string, fallback: number): number {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

export function clampNonNegativeNumber(value: string, fallback: number): number {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : fallback;
}
