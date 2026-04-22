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

export function convertLatexToSvg(input: string): string {
  const normalized = normalizeLatexInput(input);
  if (!normalized) {
    throw new Error("LaTeX 公式不能为空。");
  }

  const node = html.convert(normalized, {
    display: shouldUseDisplayMode(input),
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
