export function normalizeLatexInput(input: string): string {
  let trimmed = input.trim();
  if (trimmed.startsWith("$$") && trimmed.endsWith("$$")) {
    trimmed = trimmed.slice(2, -2);
  } else if (trimmed.startsWith("$") && trimmed.endsWith("$")) {
    trimmed = trimmed.slice(1, -1);
  } else if (trimmed.startsWith("\\(") && trimmed.endsWith("\\)")) {
    trimmed = trimmed.slice(2, -2);
  } else if (trimmed.startsWith("\\[") && trimmed.endsWith("\\]")) {
    trimmed = trimmed.slice(2, -2);
  }
  return trimmed.replace(/\r/g, "").trim();
}

export function shouldUseDisplayMode(input: string): boolean {
  const trimmed = input.trim();
  return (
    trimmed.startsWith("\\[") ||
    trimmed.startsWith("$$") ||
    /\\begin\s*\{(align|equation|gather|multline|cases|matrix)/.test(trimmed) ||
    /\\\\/.test(trimmed)
  );
}
