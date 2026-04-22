import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";
import { AllPackages } from "mathjax-full/js/input/tex/AllPackages.js";
import { normalizeLatexInput, shouldUseDisplayMode } from "../lib/latex";

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);

const tex = new TeX({
  packages: AllPackages,
  inlineMath: [["\\(", "\\)"], ["$", "$"]],
  displayMath: [["\\[", "\\]"], ["$$", "$$"]],
});

const svg = new SVG({
  fontCache: "none",
});

const html = mathjax.document("", {
  InputJax: tex,
  OutputJax: svg,
});

function encodeBase64Utf8(text: string): string {
  return btoa(unescape(encodeURIComponent(text)));
}

export function convertLatexToSvg(input: string, options: { display?: boolean } = {}): string {
  const normalized = normalizeLatexInput(input);
  if (!normalized) {
    throw new Error("LaTeX 公式不能为空。");
  }

  const node = html.convert(normalized, {
    display: options.display ?? shouldUseDisplayMode(input),
    em: 20,
    ex: 10,
    containerWidth: 80 * 20,
  });

  return adaptor
    .innerHTML(node)
    .replace(/stroke="currentColor"/g, 'stroke="#000"')
    .replace(/fill="currentColor"/g, 'fill="#000"')
    .replace(/\n{2,}/g, "\n")
    .trim();
}

function parseSvgSize(svgText: string): { width: number; height: number } {
  const widthMatch = svgText.match(/\swidth="([\d.]+)(?:ex|em|px|pt)?"/);
  const heightMatch = svgText.match(/\sheight="([\d.]+)(?:ex|em|px|pt)?"/);
  const viewBoxMatch = svgText.match(/\sviewBox="([^"]+)"/);
  if (viewBoxMatch) {
    const [, , width, height] = viewBoxMatch[1].split(/\s+/).map(Number);
    if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
      return { width: Math.ceil(width / 10), height: Math.ceil(height / 10) };
    }
  }
  return {
    width: Math.max(120, Math.ceil(Number(widthMatch?.[1] ?? 18) * 10)),
    height: Math.max(48, Math.ceil(Number(heightMatch?.[1] ?? 6) * 10)),
  };
}

export interface LatexPngResult {
  latex: string;
  svg: string;
  base64: string;
  width: number;
  height: number;
}

export async function convertLatexToPngBase64(input: string, options: { display?: boolean } = {}): Promise<LatexPngResult> {
  const latex = normalizeLatexInput(input);
  const svgText = convertLatexToSvg(input, options);
  if (typeof document === "undefined" || typeof Image === "undefined") {
    throw new Error("当前环境不支持浏览器 Canvas，无法生成 LaTeX 图片。");
  }

  const svgDataUrl = `data:image/svg+xml;base64,${encodeBase64Utf8(svgText)}`;
  const image = await new Promise<HTMLImageElement>((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("LaTeX SVG 转 PNG 失败：WebView 无法加载 SVG data URL。"));
    img.src = svgDataUrl;
  });
  const fallback = parseSvgSize(svgText);
  const scale = 2;
  const width = Math.max(1, image.naturalWidth || fallback.width);
  const height = Math.max(1, image.naturalHeight || fallback.height);
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(width * scale);
  canvas.height = Math.ceil(height * scale);
  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("当前环境无法创建 Canvas，不能生成 LaTeX PNG。");
  }
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(image, 0, 0, canvas.width, canvas.height);
  const dataUrl = canvas.toDataURL("image/png");
  const base64 = dataUrl.split(",")[1];
  if (!base64) {
    throw new Error("LaTeX PNG 编码失败。");
  }
  return { latex, svg: svgText, base64, width, height };
}
