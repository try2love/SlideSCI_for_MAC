import type { Box, NativeEquationRun, TextRun, TextStyle } from "../lib/types";
import { getSelectedShapeIds, selectTextRange } from "./powerpoint";

const HELPER_BASE_URLS = ["/native-helper", "http://127.0.0.1:17926", "http://localhost:17926"];
let preferredHelperBaseUrl = HELPER_BASE_URLS[0];

export interface NativeEquationHelperHealth {
  ok: boolean;
  powerpointRunning?: boolean;
  nativeEquationAvailable?: boolean;
  guiAutomationAvailable?: boolean;
  accessibilityGranted?: boolean;
  message: string;
}

export interface NativeEquationConversionResponse {
  ok: boolean;
  mode: "native" | "unsupported";
  id?: string;
  nativeCount?: number;
  message: string;
}

export interface NativeEquationConversionSummary {
  shapeId: string;
  nativeCount: number;
  fallbackCount: number;
  messages: string[];
  remainingEquations: NativeEquationRun[];
}

async function readJsonResponse<T>(response: Response): Promise<T> {
  const text = await response.text();
  let data: any = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(`helper 返回了非 JSON 响应（HTTP ${response.status}）。请确认 npm run helper 正在运行，且 Vite /native-helper 代理已生效。`);
  }
  if (!response.ok) {
    throw new Error(data.message || response.statusText);
  }
  return data as T;
}

function messageFromError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  if (/load failed|failed to fetch|networkerror|network request failed/i.test(message)) {
    return "任务窗格无法访问本地公式 helper。请确认已运行 npm run helper，Vite 开发服务已重启以启用 /native-helper 代理，并允许本机 17926 端口访问。";
  }
  return message;
}

function shouldTryNextHelperUrl(baseUrl: string, response: Response): boolean {
  if (!baseUrl.startsWith("/")) {
    return false;
  }
  const contentType = response.headers.get("Content-Type") ?? "";
  return response.status === 404 || !contentType.toLowerCase().includes("application/json");
}

async function fetchHelper(path: string, init?: RequestInit): Promise<Response> {
  const urls = [
    preferredHelperBaseUrl,
    ...HELPER_BASE_URLS.filter((url) => url !== preferredHelperBaseUrl),
  ];
  let lastError: unknown;

  for (const baseUrl of urls) {
    try {
      const response = await fetch(`${baseUrl}${path}`, init);
      if (shouldTryNextHelperUrl(baseUrl, response)) {
        lastError = new Error(`同源 helper 代理不可用：HTTP ${response.status}`);
        continue;
      }
      preferredHelperBaseUrl = baseUrl;
      return response;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError;
}

export async function checkNativeEquationHelper(): Promise<NativeEquationHelperHealth> {
  try {
    const response = await fetchHelper("/health");
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
  const health = await checkNativeEquationHelper();
  ensureNativeEquationAvailable(health);
  try {
    const response = await fetchHelper("/equation/convert-selection", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ latex, display }),
    });
    return readJsonResponse<NativeEquationConversionResponse>(response);
  } catch (error) {
    throw new Error(messageFromError(error));
  }
}

export interface NativeEquationTextBoxRequest {
  text: string;
  equations: NativeEquationRun[];
  box: Box;
  baseStyle?: TextStyle;
  runs?: TextRun[];
}

export interface NativeEquationBlockRequest {
  latex: string;
  box: Box;
  style?: TextStyle;
}

async function postNativeEquation(path: string, payload: unknown): Promise<NativeEquationConversionResponse> {
  const response = await fetchHelper(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  return readJsonResponse<NativeEquationConversionResponse>(response);
}

export function ensureNativeEquationAvailable(health: NativeEquationHelperHealth): void {
  if (!health.ok || !health.nativeEquationAvailable) {
    throw new Error(health.message || "本地公式 helper 不可用。请先运行 npm run helper 并授权 PowerPoint 自动化。");
  }
}

export async function insertNativeEquationTextBox(request: NativeEquationTextBoxRequest): Promise<NativeEquationConversionResponse> {
  void request;
  throw new Error("insertNativeEquationTextBox 已弃用。请改为先用 Office.js 插入文本框并选中文本，再调用 convertSelectedTextToNativeEquation。");
}

export async function insertNativeEquationBlock(request: NativeEquationBlockRequest): Promise<NativeEquationConversionResponse> {
  void request;
  throw new Error("insertNativeEquationBlock 已弃用。请改为先用 Office.js 插入文本框并选中文本，再调用 convertSelectedTextToNativeEquation。");
}

export interface EquationRunConversionDeps {
  selectRange: (shapeId: string, start: number, length: number) => Promise<void>;
  getSelectedIds: () => Promise<string[]>;
  convertSelection: (latex: string, display?: boolean) => Promise<NativeEquationConversionResponse>;
}

const defaultEquationRunDeps: EquationRunConversionDeps = {
  selectRange: selectTextRange,
  getSelectedIds: getSelectedShapeIds,
  convertSelection: convertSelectedTextToNativeEquation,
};

function sortEquationsDescending(equations: NativeEquationRun[]): NativeEquationRun[] {
  return [...equations].sort((a, b) => Number(b.start) - Number(a.start));
}

export async function convertEquationRuns(
  shapeId: string,
  equations: NativeEquationRun[],
  deps: EquationRunConversionDeps = defaultEquationRunDeps,
): Promise<NativeEquationConversionSummary> {
  const sortedEquations = sortEquationsDescending(equations);
  let currentShapeId = shapeId;
  const messages: string[] = [];
  let nativeCount = 0;

  for (const [index, equation] of sortedEquations.entries()) {
    try {
      await deps.selectRange(currentShapeId, equation.start, equation.length);
      const response = await deps.convertSelection(equation.latex, equation.display);
      nativeCount += 1;
      if (response.message) {
        messages.push(response.message);
      }
      const selectedIds = await deps.getSelectedIds();
      if (selectedIds[0]) {
        currentShapeId = selectedIds[0];
      }
    } catch (error) {
      messages.push(messageFromError(error));
      return {
        shapeId: currentShapeId,
        nativeCount,
        fallbackCount: sortedEquations.length - nativeCount,
        messages,
        remainingEquations: sortedEquations.slice(index),
      };
    }
  }

  return {
    shapeId: currentShapeId,
    nativeCount,
    fallbackCount: 0,
    messages,
    remainingEquations: [],
  };
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
