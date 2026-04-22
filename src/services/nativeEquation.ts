import type { NativeEquationRun } from "../lib/types";
import { selectTextRange } from "./powerpoint";

const HELPER_BASE_URL = "http://127.0.0.1:17926";

export interface NativeEquationHelperHealth {
  ok: boolean;
  powerpointRunning?: boolean;
  nativeEquationAvailable?: boolean;
  message: string;
}

export interface NativeEquationConversionResponse {
  ok: boolean;
  mode: "native" | "unsupported";
  message: string;
}

export interface NativeEquationConversionSummary {
  nativeCount: number;
  fallbackCount: number;
  messages: string[];
}

async function readJsonResponse<T>(response: Response): Promise<T> {
  const text = await response.text();
  const data = text ? JSON.parse(text) : {};
  if (!response.ok) {
    throw new Error(data.message || response.statusText);
  }
  return data as T;
}

function messageFromError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export async function checkNativeEquationHelper(): Promise<NativeEquationHelperHealth> {
  try {
    const response = await fetch(`${HELPER_BASE_URL}/health`);
    return readJsonResponse<NativeEquationHelperHealth>(response);
  } catch (error) {
    return {
      ok: false,
      nativeEquationAvailable: false,
      message: `本地公式 helper 未运行或无法访问：${messageFromError(error)}`,
    };
  }
}

export async function convertSelectedTextToNativeEquation(latex: string, display = false): Promise<NativeEquationConversionResponse> {
  const response = await fetch(`${HELPER_BASE_URL}/equation/convert-selection`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ latex, display }),
  });
  return readJsonResponse<NativeEquationConversionResponse>(response);
}

export async function convertEquationRuns(
  shapeId: string,
  equations: NativeEquationRun[],
  deps: {
    selectRange?: typeof selectTextRange;
    convertSelection?: typeof convertSelectedTextToNativeEquation;
  } = {},
): Promise<NativeEquationConversionSummary> {
  const selectRange = deps.selectRange ?? selectTextRange;
  const convertSelection = deps.convertSelection ?? convertSelectedTextToNativeEquation;
  const summary: NativeEquationConversionSummary = { nativeCount: 0, fallbackCount: 0, messages: [] };

  for (const equation of [...equations].sort((a, b) => b.start - a.start)) {
    try {
      await selectRange(shapeId, equation.start, equation.length);
      const result = await convertSelection(equation.latex, equation.display);
      if (result.ok && result.mode === "native") {
        summary.nativeCount += 1;
      } else {
        summary.fallbackCount += 1;
        summary.messages.push(result.message || `公式 ${equation.latex} 未转换为原生公式。`);
      }
    } catch (error) {
      summary.fallbackCount += 1;
      summary.messages.push(`公式 ${equation.latex} 保留为文本：${messageFromError(error)}`);
    }
  }

  return summary;
}

export function formatEquationConversionSummary(summary: NativeEquationConversionSummary): string | undefined {
  const total = summary.nativeCount + summary.fallbackCount;
  if (total === 0) {
    return undefined;
  }
  const base = `原生公式成功 ${summary.nativeCount} 个，公式降级 ${summary.fallbackCount} 个`;
  if (summary.messages.length === 0) {
    return base;
  }
  return `${base}：${summary.messages.slice(0, 2).join("；")}`;
}
