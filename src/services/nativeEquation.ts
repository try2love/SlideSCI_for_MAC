import type { Box, NativeEquationRun, TextRun, TextStyle } from "../lib/types";
import { convertLatexToUnicodeMath } from "./unicodeMath";
import {
  addRichTextBox,
  addTextBox,
  deleteShapes,
  getPowerPointHostCapabilities,
  getSelectedShapeIds,
  selectShapes,
  selectTextRange,
} from "./powerpoint";

const HELPER_BASE_URLS = ["/native-helper", "http://127.0.0.1:17926", "http://localhost:17926"];
let preferredHelperBaseUrl = HELPER_BASE_URLS[0];

export type NativeEquationHelperStrategy = "latex-ribbon" | "unicode-math";
export const DEFAULT_EQUATION_STRATEGY_ORDER: NativeEquationHelperStrategy[] = ["latex-ribbon"];

export interface NativeEquationHelperHealth {
  ok: boolean;
  helperBuildId?: string;
  scriptExecutionMode?: "temp-file";
  powerpointRunning?: boolean;
  nativeEquationAvailable?: boolean;
  guiAutomationAvailable?: boolean;
  accessibilityGranted?: boolean;
  hostSelectionApiRequired?: boolean;
  latexRibbonAvailable?: boolean;
  unicodeMathFallbackAvailable?: boolean;
  equationScriptSyntaxOk?: boolean;
  equationScriptSyntaxMessage?: string;
  message: string;
}

export interface NativeEquationConversionResponse {
  ok: boolean;
  helperBuildId?: string;
  mode: "native" | "unsupported";
  id?: string;
  nativeCount?: number;
  strategyUsed?: NativeEquationHelperStrategy;
  message: string;
}

export interface NativeEquationConversionSummary {
  shapeId: string;
  strategy: "officejs-selection" | "helper-gui";
  strategiesUsed: NativeEquationHelperStrategy[];
  nativeCount: number;
  fallbackCount: number;
  messages: string[];
  remainingEquations: NativeEquationRun[];
}

export interface NativeEquationPlaceholder extends NativeEquationRun {
  token: string;
  unicodeMath?: string;
}

export interface NativeEquationShapeRangeRequest {
  shapeId: string;
  originalText: string;
  workingText: string;
  placeholders: NativeEquationPlaceholder[];
  mode: "inline" | "block";
  strategyOrder: NativeEquationHelperStrategy[];
}

export interface NativeEquationSelectionRequest {
  latex: string;
  display?: boolean;
  unicodeMath?: string;
  strategyOrder: NativeEquationHelperStrategy[];
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
    const buildSuffix = data.helperBuildId ? ` [helper build ${data.helperBuildId}]` : "";
    throw new Error((data.message || response.statusText) + buildSuffix);
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
    const unicodeMath = tryConvertLatexToUnicodeMath(latex);
    const response = await fetchHelper("/equation/convert-selection", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        latex,
        display,
        unicodeMath,
        strategyOrder: DEFAULT_EQUATION_STRATEGY_ORDER,
      } satisfies NativeEquationSelectionRequest),
    });
    return readJsonResponse<NativeEquationConversionResponse>(response);
  } catch (error) {
    throw new Error(messageFromError(error));
  }
}

export async function convertShapeRangesToNativeEquations(
  request: NativeEquationShapeRangeRequest,
): Promise<NativeEquationConversionResponse> {
  const health = await checkNativeEquationHelper();
  ensureNativeEquationAvailable(health);
  try {
    const response = await fetchHelper("/equation/convert-shape-ranges", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(request),
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
    const buildSuffix = health.helperBuildId ? ` [helper build ${health.helperBuildId}]` : "";
    throw new Error((health.message || "本地公式 helper 不可用。请先运行 npm run helper 并授权 PowerPoint 自动化。") + buildSuffix);
  }
}

function tryConvertLatexToUnicodeMath(latex: string): string | undefined {
  try {
    return convertLatexToUnicodeMath(latex);
  } catch {
    return undefined;
  }
}

function placeholderToken(index: number, length: number): string {
  const alphabet = "QWERTYUIOPASDFGHJKLZXCVBNM0123456789";
  if (length <= 1) {
    return alphabet[index % alphabet.length];
  }
  const seed = `${alphabet[index % alphabet.length]}${index.toString(36).toUpperCase()}`;
  let token = "";
  while (token.length < length) {
    token += seed;
  }
  return token.slice(0, length);
}

export function shouldRetryWithGuiShapeRange(error: unknown): boolean {
  const message = messageFromError(error);
  return /无法进入文本编辑状态|请先选中文本框|AXSelectedTextRange|-2700/i.test(message);
}

export function buildShapeRangeEquationRequest(
  shapeId: string,
  originalText: string,
  equations: NativeEquationRun[],
  mode: "inline" | "block",
): NativeEquationShapeRangeRequest {
  const placeholders = [...equations]
    .sort((a, b) => Number(b.start) - Number(a.start))
    .map((equation, index) => ({
      ...equation,
      token: placeholderToken(index, Math.max(1, equation.length)),
      unicodeMath: tryConvertLatexToUnicodeMath(equation.latex),
    }));

  let workingText = originalText;
  for (const placeholder of placeholders) {
    workingText =
      workingText.slice(0, placeholder.start) +
      placeholder.token +
      workingText.slice(placeholder.start + placeholder.length);
  }

  return {
    shapeId,
    originalText,
    workingText,
    placeholders,
    mode,
    strategyOrder: DEFAULT_EQUATION_STRATEGY_ORDER,
  };
}

export async function insertNativeEquationTextBox(request: NativeEquationTextBoxRequest): Promise<NativeEquationConversionResponse> {
  const capabilities = await getPowerPointHostCapabilities();
  const guiRequest = buildShapeRangeEquationRequest("", request.text, request.equations, "inline");

  if (capabilities.textRangeSelection) {
    const shapeId = await addRichTextBox(request.text, request.box, request.baseStyle ?? {}, request.runs ?? []);
    const summary = await convertEquationRuns(shapeId, request.equations);
    if (summary.fallbackCount === 0) {
      return {
        ok: true,
        mode: "native",
        id: summary.shapeId,
        nativeCount: summary.nativeCount,
        strategyUsed: summary.strategiesUsed.at(0),
        message: formatEquationConversionSummary(summary) ?? "原生公式转换完成。",
      };
    }
    const failureMessage = summary.messages.at(-1);
    if (!shouldRetryWithGuiShapeRange(failureMessage)) {
      return {
        ok: false,
        mode: "unsupported",
        id: summary.shapeId,
        nativeCount: summary.nativeCount,
        strategyUsed: summary.strategiesUsed.at(0),
        message: formatEquationConversionSummary(summary) ?? "原生公式转换失败。",
      };
    }
    await deleteShapes([...new Set([shapeId, summary.shapeId].filter(Boolean))]);
  }

  const shapeId = await addRichTextBox(guiRequest.workingText, request.box, request.baseStyle ?? {}, request.runs ?? []);
  await selectShapes([shapeId]);
  const response = await convertShapeRangesToNativeEquations({ ...guiRequest, shapeId });
  const selectedIds = await getSelectedShapeIds();
  return {
    ...response,
    id: selectedIds[0] || shapeId,
  };
}

export async function insertNativeEquationBlock(request: NativeEquationBlockRequest): Promise<NativeEquationConversionResponse> {
  const capabilities = await getPowerPointHostCapabilities();
  const guiRequest = buildShapeRangeEquationRequest(
    "",
    request.latex,
    [{ start: 0, length: request.latex.length, latex: request.latex, display: true }],
    "block",
  );
  if (capabilities.textRangeSelection) {
    const shapeId = await addTextBox(request.latex, request.box, request.style ?? {});
    try {
      await selectTextRange(shapeId, 0, request.latex.length);
      const response = await convertSelectedTextToNativeEquation(request.latex, true);
      const selectedIds = await getSelectedShapeIds();
      return {
        ...response,
        id: selectedIds[0] || shapeId,
      };
    } catch (error) {
      if (!shouldRetryWithGuiShapeRange(error)) {
        throw error;
      }
      await deleteShapes([shapeId]);
    };
  }

  const shapeId = await addTextBox(guiRequest.workingText, request.box, request.style ?? {});
  await selectShapes([shapeId]);
  const response = await convertShapeRangesToNativeEquations({ ...guiRequest, shapeId });
  const selectedIds = await getSelectedShapeIds();
  return {
    ...response,
    id: selectedIds[0] || shapeId,
  };
}

export interface EquationRunConversionDeps {
  selectRange: (shapeId: string, start: number, length: number) => Promise<void>;
  getSelectedIds: () => Promise<string[]>;
  convertSelection: (latex: string, display?: boolean) => Promise<NativeEquationConversionResponse>;
}

function sortEquationsDescending(equations: NativeEquationRun[]): NativeEquationRun[] {
  return [...equations].sort((a, b) => Number(b.start) - Number(a.start));
}

export async function convertEquationRuns(
  shapeId: string,
  equations: NativeEquationRun[],
  deps?: EquationRunConversionDeps,
): Promise<NativeEquationConversionSummary> {
  const resolvedDeps = deps ?? {
    selectRange: selectTextRange,
    getSelectedIds: getSelectedShapeIds,
    convertSelection: convertSelectedTextToNativeEquation,
  };
  const sortedEquations = sortEquationsDescending(equations);
  let currentShapeId = shapeId;
  const messages: string[] = [];
  const strategiesUsed = new Set<NativeEquationHelperStrategy>();
  let nativeCount = 0;

  for (const [index, equation] of sortedEquations.entries()) {
    try {
      await resolvedDeps.selectRange(currentShapeId, equation.start, equation.length);
      const response = await resolvedDeps.convertSelection(equation.latex, equation.display);
      nativeCount += 1;
      if (response.strategyUsed) {
        strategiesUsed.add(response.strategyUsed);
      }
      if (response.message) {
        messages.push(response.message);
      }
      const selectedIds = await resolvedDeps.getSelectedIds();
      if (selectedIds[0]) {
        currentShapeId = selectedIds[0];
      }
    } catch (error) {
      messages.push(messageFromError(error));
      return {
        shapeId: currentShapeId,
        strategy: "officejs-selection",
        strategiesUsed: [...strategiesUsed],
        nativeCount,
        fallbackCount: sortedEquations.length - nativeCount,
        messages,
        remainingEquations: sortedEquations.slice(index),
      };
    }
  }

  return {
    shapeId: currentShapeId,
    strategy: "officejs-selection",
    strategiesUsed: [...strategiesUsed],
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
  const strategyLabel = summary.strategiesUsed.length === 1 ? `（${summary.strategiesUsed[0]}）` : "";
  const base = `原生公式成功 ${summary.nativeCount} 个${strategyLabel}，公式降级 ${summary.fallbackCount} 个`;
  if (summary.messages.length === 0) {
    return base;
  }
  return `${base}：${summary.messages.slice(0, 2).join("；")}`;
}
