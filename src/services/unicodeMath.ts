import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";
import { AllPackages } from "mathjax-full/js/input/tex/AllPackages.js";
import { JsonMmlVisitor } from "mathjax-full/js/core/MmlTree/JsonMmlVisitor.js";
import { normalizeLatexInput } from "../lib/latex";

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

const jsonVisitor = new JsonMmlVisitor();

interface JsonMmlNode {
  kind: string;
  text?: string;
  childNodes?: JsonMmlNode[];
  properties?: Record<string, unknown>;
  attributes?: Record<string, unknown>;
}

function parseLatexToJsonTree(input: string): JsonMmlNode {
  const item = new html.options.MathItem(input, tex, false);
  item.start = { node: null, delim: "", n: 0 };
  item.end = { node: null, delim: "", n: 0 };
  item.inputData = {};
  item.compile(html);
  return jsonVisitor.visitTree(item.root) as JsonMmlNode;
}

function textOf(node: JsonMmlNode | undefined): string {
  if (!node) {
    return "";
  }
  if (node.kind === "text") {
    return node.text ?? "";
  }
  return (node.childNodes ?? []).map((child) => textOf(child)).join("");
}

function isBareToken(text: string): boolean {
  return /^[\p{L}\p{N}∞∑∫∏∂∇α-ωΑ-Ω]+$/u.test(text);
}

function wrapForLinearMath(text: string): string {
  if (!text) {
    return "()";
  }
  return isBareToken(text) ? text : `(${text})`;
}

function joinChildren(node: JsonMmlNode | undefined): string {
  return (node?.childNodes ?? []).map((child) => serializeNode(child)).join("");
}

function serializeAccentedNode(kind: string, baseNode?: JsonMmlNode, accentNode?: JsonMmlNode): string {
  const base = serializeNode(baseNode);
  const accent = textOf(accentNode);
  if (kind === "mover") {
    if (accent === "→") {
      return `${wrapForLinearMath(base)}\\vec`;
    }
    if (accent === "¯" || accent === "‾") {
      return `\\overbar(${base})`;
    }
    if (accent === "^" || accent === "ˆ" || accent === "∧") {
      return `${wrapForLinearMath(base)}\\hat`;
    }
    if (accent === "˙") {
      return `\\dot(${base})`;
    }
  }
  if (kind === "munder") {
    return `${base}_(${serializeNode(accentNode)})`;
  }
  return `${base}^(${serializeNode(accentNode)})`;
}

function serializeTable(node: JsonMmlNode): string {
  const rows = (node.childNodes ?? [])
    .filter((child) => child.kind === "mtr" || child.kind === "mlabeledtr")
    .map((row) =>
      (row.childNodes ?? [])
        .filter((cell) => cell.kind === "mtd")
        .map((cell) => joinChildren(cell).trim())
        .join("&"),
    );

  return `\\matrix(${rows.join("@")})`;
}

function serializeNode(node: JsonMmlNode | undefined): string {
  if (!node) {
    return "";
  }

  switch (node.kind) {
    case "math":
      return joinChildren(node).trim();
    case "text":
      return node.text ?? "";
    case "mrow":
    case "TeXAtom":
    case "semantics":
    case "mstyle":
    case "mpadded":
    case "menclose":
    case "mphantom":
    case "mtd":
      return joinChildren(node);
    case "mi":
    case "mn":
    case "mo":
    case "mtext":
    case "ms":
      return textOf(node);
    case "mspace":
      return " ";
    case "mfrac": {
      const [numerator, denominator] = node.childNodes ?? [];
      return `${wrapForLinearMath(serializeNode(numerator))}/${wrapForLinearMath(serializeNode(denominator))}`;
    }
    case "msqrt":
      return `\\sqrt(${joinChildren(node)})`;
    case "mroot": {
      const [base, index] = node.childNodes ?? [];
      return `\\sqrt[${serializeNode(index)}](${serializeNode(base)})`;
    }
    case "msub": {
      const [base, sub] = node.childNodes ?? [];
      return `${serializeNode(base)}_(${serializeNode(sub)})`;
    }
    case "msup": {
      const [base, sup] = node.childNodes ?? [];
      return `${serializeNode(base)}^(${serializeNode(sup)})`;
    }
    case "msubsup": {
      const [base, sub, sup] = node.childNodes ?? [];
      return `${serializeNode(base)}_(${serializeNode(sub)})^(${serializeNode(sup)})`;
    }
    case "munder": {
      const [base, under] = node.childNodes ?? [];
      return serializeAccentedNode("munder", base, under);
    }
    case "mover": {
      const [base, over] = node.childNodes ?? [];
      return serializeAccentedNode("mover", base, over);
    }
    case "munderover": {
      const [base, under, over] = node.childNodes ?? [];
      return `${serializeNode(base)}_(${serializeNode(under)})^(${serializeNode(over)})`;
    }
    case "mtable":
      return serializeTable(node);
    case "mtr":
    case "mlabeledtr":
      return (node.childNodes ?? []).map((child) => serializeNode(child)).join("&");
    case "mmultiscripts": {
      const children = node.childNodes ?? [];
      const base = serializeNode(children[0]);
      const preIndex = children.findIndex((child) => child.kind === "mprescripts");
      const posts = preIndex === -1 ? children.slice(1) : children.slice(1, preIndex);
      const postSub = posts[0]?.kind === "none" ? "" : serializeNode(posts[0]);
      const postSup = posts[1]?.kind === "none" ? "" : serializeNode(posts[1]);
      let result = base;
      if (postSub) {
        result += `_(${postSub})`;
      }
      if (postSup) {
        result += `^(${postSup})`;
      }
      return result;
    }
    case "mfenced": {
      const open = String(node.attributes?.open ?? "(");
      const close = String(node.attributes?.close ?? ")");
      return `${open}${joinChildren(node)}${close}`;
    }
    case "none":
      return "";
    default:
      throw new Error(`暂不支持的 MathML 节点：${node.kind}`);
  }
}

export function convertLatexToUnicodeMath(input: string): string {
  const normalized = normalizeLatexInput(input);
  if (!normalized) {
    throw new Error("LaTeX 公式不能为空。");
  }

  const tree = parseLatexToJsonTree(normalized);
  const unicodeMath = serializeNode(tree)
    .replace(/\s{2,}/g, " ")
    .replace(/\(\s+/g, "(")
    .replace(/\s+\)/g, ")")
    .trim();

  if (!unicodeMath) {
    throw new Error("未能生成 UnicodeMath 线性表达式。");
  }

  return unicodeMath;
}
