#!/usr/bin/env node
"use strict";

const readline = require("readline");
const { mathjax } = require("mathjax-full/js/mathjax.js");
const { TeX } = require("mathjax-full/js/input/tex.js");
const { SVG } = require("mathjax-full/js/output/svg.js");
const { liteAdaptor } = require("mathjax-full/js/adaptors/liteAdaptor.js");
const { RegisterHTMLHandler } = require("mathjax-full/js/handlers/html.js");
const { AllPackages } = require("mathjax-full/js/input/tex/AllPackages.js");

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);

const { fitMathType, stackFittedLines } = require("./mathtype_fit.cjs");

/** Fit-mode padding (pt), tuned against GT MathType v:shape boxes. */
const FIT_PAD_TOP_PT = 3.23;
const FIT_PAD_BOTTOM_PT = 2.64;

function convertToSvg(latex) {
  const node = html.convert(latex, { display: false, em: 16, ex: 8, containerWidth: 100000 });
  let svg = adaptor.outerHTML(node);
  const start = svg.indexOf("<svg");
  const end = svg.lastIndexOf("</svg>");
  if (start < 0 || end < 0) {
    throw new Error("MathJax did not produce SVG");
  }
  const rendered = svg.slice(start, end + 6)
    .replace(/currentColor/g, "#000000")
    .replace(/\s+focusable="false"/g, "")
    .replace(/\s+role="img"/g, "");
  if (/data-mml-node=["']merror["']/.test(rendered)
      || /data-mml-node=["']mtext["'][^>]*(?:fill|stroke)=["']red["']/.test(rendered)) {
    throw new Error(`MathJax produced an error glyph for: ${latex}`);
  }
  return rendered;
}

function materializeArrayRuleStyles(svg) {
  return svg.replace(/<(line|rect)\b([^>]*(?:data-line|data-frame)=["'][^"']+["'][^>]*)>/g,
    (tag, element, attrs) => {
      const dashed = /class=["']mjx-dashed["']/.test(attrs);
      const cleaned = attrs.replace(/\s*\/$/, "");
      const dash = dashed ? ' stroke-dasharray="120 90"' : "";
      const close = /\/$/.test(attrs) ? "/>" : ">";
      return `<${element}${cleaned} fill="none" stroke="#000000" stroke-width="60"${dash}${close}`;
    });
}

/** Split on top-level \\ line breaks (MathType piles); respects brace depth
 * and \\begin/\\end environments (array rows are NOT pile lines).
 * Consumes the standard modifiers \\* and \\[len]; a [len] argument becomes
 * extra pile gap (em) before the following line instead of leaking into the
 * rendered output as literal text. */
function splitTopLevelLines(latex, fontPt) {
  const emPt = fontPt > 0 ? fontPt : 10;
  const UNIT_PT = { pt: 1, bp: 72 / 72.27, "in": 72, cm: 72 / 2.54,
                    mm: 72 / 25.4, pc: 12, em: emPt, ex: emPt * 0.431 };
  const lines = [];
  const gapsEm = [0];
  let depth = 0;
  let envDepth = 0;
  let current = "";
  for (let i = 0; i < latex.length; i++) {
    const ch = latex[i];
    if (ch === "\\" && latex[i + 1] === "\\" && depth === 0 && envDepth === 0) {
      lines.push(current);
      current = "";
      i += 1;
      // optional modifiers: \\* and/or \\[<len>]
      let j = i + 1;
      if (latex[j] === "*") {
        j += 1;
      }
      const mod = /^\s*\[\s*([+-]?[\d.]+)\s*(pt|bp|in|cm|mm|pc|em|ex)?\s*\]/
        .exec(latex.slice(j));
      let extra = 0;
      if (mod) {
        const factor = UNIT_PT[mod[2] || "pt"] || 1;
        extra = Math.max(0, parseFloat(mod[1]) * factor) / emPt;
        j += mod[0].length;
      }
      gapsEm.push(extra);
      i = j - 1;
      continue;
    }
    if (ch === "\\") {
      if (latex.startsWith("begin", i + 1)) envDepth += 1;
      else if (latex.startsWith("end", i + 1)) envDepth = Math.max(0, envDepth - 1);
    }
    if (ch === "{") depth += 1;
    if (ch === "}") depth = Math.max(0, depth - 1);
    current += ch;
  }
  lines.push(current);
  const trimmed = [];
  const trimmedGaps = [];
  lines.forEach((line, idx) => {
    const t = line.trim();
    if (t.length > 0) {
      trimmed.push(t);
      trimmedGaps.push(trimmed.length === 1 ? 0 : (gapsEm[idx] || 0));
    }
  });
  return trimmed.length > 0
    ? { lines: trimmed, gapsEm: trimmedGaps }
    : { lines: [latex], gapsEm: [0] };
}

const tex = new TeX({
  packages: AllPackages,
  macros: {
    overarc: ["\\overset{\\frown}{#1}", 1],
    arc: ["\\overset{\\frown}{#1}", 1],
    wideparen: ["\\overset{\\frown}{#1}", 1],
    whitestar: "\\star",
    blackstar: "\\star",
    whitediamond: "\\diamond",
    underbracechar: "\\underbrace{\\hphantom{0}}",
    upslopeellipsis: "\\begin{smallmatrix}&&\\cdot\\\\&\\cdot&\\\\\\cdot&&\\end{smallmatrix}",
    leftbarharpoon: "\\leftharpoonup\\!|",
    rightbarharpoon: "|\\!\\rightharpoonup",
    dlsh: "\\hookleftarrow",
    nwsearrow: "\\nwarrow\\!\\searrow",
    neswarrow: "\\nearrow\\!\\swarrow",
    nicefrac: ["{}^{#1}\\!/\\!{}_{#2}", 2],
    ltr: ["#1", 1],
    rtl: ["#1", 1],
    xlongrightarrow: "\\xrightarrow",
    xlongleftarrow: "\\xleftarrow",
    xlongleftrightarrow: "\\xleftrightarrow",
    xLongrightarrow: "\\xRightarrow",
    xLongleftarrow: "\\xLeftarrow",
    xLongleftrightarrow: "\\xLeftrightarrow"
  }
});
const svgOutput = new SVG({ fontCache: "none", internalSpeechTitles: false });
const html = mathjax.document("", { InputJax: tex, OutputJax: svgOutput });

function numberAttr(svg, name) {
  const match = svg.match(new RegExp("\\b" + name + "=['\"]([^'\"]+)['\"]", "i"));
  if (!match) {
    return 0;
  }
  const value = match[1].trim();
  const number = Number.parseFloat(value);
  return Number.isFinite(number) ? number : 0;
}

function viewBox(svg) {
  const match = svg.match(/\bviewBox=['"]([^'"]+)['"]/i);
  if (!match) {
    return [0, 0, 1000, 1000];
  }
  const parts = match[1].trim().split(/\s+/).map(Number);
  return parts.length === 4 && parts.every(Number.isFinite) ? parts : [0, 0, 1000, 1000];
}

function verticalAlignEx(svg) {
  const match = svg.match(/vertical-align:\s*([-+0-9.]+)ex/i);
  if (!match) {
    return 0;
  }
  const value = Number.parseFloat(match[1]);
  return Number.isFinite(value) ? value : 0;
}

function renderLatex(request) {
  const latex = Buffer.from(request.latexBase64 || "", "base64").toString("utf8");
  const fontPt = finiteOr(request.fontPt, 9.02);
  const exRatio = finiteOr(request.exRatio, 0.431);
  const paddingPt = finiteOr(request.paddingPt, 2.3);
  const maxWidthPt = finiteOr(request.maxWidthPt, 400);
  const split = request.mathTypeFit
    ? splitTopLevelLines(latex, fontPt)
    : { lines: [latex], gapsEm: [0] };
  const lines = split.lines;
  let svg = convertToSvg(lines[0]);

  const widthEx = Math.max(numberAttr(svg, "width"), 0.1);
  const heightEx = Math.max(numberAttr(svg, "height"), 0.1);
  const valignEx = verticalAlignEx(svg);
  const exPt = fontPt * exRatio;
  let contentWidthPt = widthEx * exPt;
  let contentHeightPt = heightEx * exPt;

  const vb = viewBox(svg);
  const unitsPerPt = contentWidthPt > 0 ? vb[2] / contentWidthPt : 1000 / fontPt;

  let contentViewBox = vb;
  let fitDepthPt = null;
  let padTopPt = paddingPt;
  let padBottomPt = paddingPt;
  if (request.mathTypeFit) {
    // Re-layout fractions/delimiters with measured MathType geometry.
    // emUnits is intrinsic to the SVG (independent of the requested fontPt).
    const emUnits = exRatio > 0 ? vb[2] / (widthEx * exRatio) : 1000;
    const parts = lines.map((line, idx) => {
      const lineSvg = idx === 0 ? svg : convertToSvg(line);
      return fitMathType(lineSvg, emUnits, request.fitParams);
    });
    const fit = parts.length > 1
      ? stackFittedLines(parts, emUnits, request.fitParams, split.gapsEm)
      : parts[0];
    svg = fit.svg;
    const [fx0, fy0, fx1, fy1] = fit.bbox;
    contentWidthPt = (fx1 - fx0) / unitsPerPt;
    contentHeightPt = (fy1 - fy0) / unitsPerPt;
    contentViewBox = [fx0, -fy1, fx1 - fx0, fy1 - fy0];
    fitDepthPt = Math.max(0, -fy0) / unitsPerPt;
    // GT v:shape boxes carry a slightly larger top margin than bottom.
    padTopPt = finiteOr(request.fitPadTopPt, FIT_PAD_TOP_PT);
    padBottomPt = finiteOr(request.fitPadBottomPt, FIT_PAD_BOTTOM_PT);
  }

  const unscaledWidthPt = contentWidthPt + paddingPt * 2;
  const unscaledHeightPt = contentHeightPt + padTopPt + padBottomPt;
  const scale = maxWidthPt > 0 ? Math.min(1, maxWidthPt / unscaledWidthPt) : 1;
  const widthPt = Math.max(unscaledWidthPt * scale, 1);
  const heightPt = Math.max(unscaledHeightPt * scale, 1);
  const depthPt = fitDepthPt !== null
    ? Math.max(fitDepthPt * scale + padBottomPt * scale, 0)
    : Math.max(Math.abs(valignEx) * exPt * scale + paddingPt * scale, 0);

  const padUnits = paddingPt * unitsPerPt;
  const padTopUnits = padTopPt * unitsPerPt;
  const padBottomUnits = padBottomPt * unitsPerPt;
  const paddedViewBox = [
    contentViewBox[0] - padUnits,
    contentViewBox[1] - padTopUnits,
    contentViewBox[2] + padUnits * 2,
    contentViewBox[3] + padTopUnits + padBottomUnits
  ];

  svg = svg
    .replace(/\bwidth=['"][^'"]+['"]/i, `width="${widthPt.toFixed(6)}pt"`)
    .replace(/\bheight=['"][^'"]+['"]/i, `height="${heightPt.toFixed(6)}pt"`)
    .replace(/\bviewBox=['"][^'"]+['"]/i, `viewBox="${paddedViewBox.map(v => v.toFixed(3)).join(" ")}"`)
    .replace(/\s+style=['"][^'"]*['"]/i, "");
  svg = materializeArrayRuleStyles(svg);

  return {
    ok: true,
    svgBase64: Buffer.from(svg, "utf8").toString("base64"),
    widthPt,
    heightPt,
    depthPt,
    widthEx,
    heightEx,
    verticalAlignEx: valignEx,
    scale
  };
}

function finiteOr(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function respond(response) {
  process.stdout.write(JSON.stringify(response) + "\n");
}

async function worker() {
  const rl = readline.createInterface({ input: process.stdin, crlfDelay: Infinity });
  for await (const line of rl) {
    if (!line.trim()) {
      continue;
    }
    let request;
    try {
      request = JSON.parse(line);
      respond({ id: request.id, ...renderLatex(request) });
    } catch (error) {
      respond({
        id: request && request.id !== undefined ? request.id : -1,
        ok: false,
        error: error && error.message ? error.message : String(error)
      });
    }
  }
}

if (process.argv.includes("--smoke")) {
  const latexBase64 = Buffer.from("\\sqrt{\\frac{a}{b}}+x_i^2", "utf8").toString("base64");
  const result = renderLatex({ id: 1, latexBase64, fontPt: 9.02, exRatio: 0.431, paddingPt: 2.3, maxWidthPt: 400 });
  console.log(JSON.stringify({ ok: result.ok, widthPt: result.widthPt, heightPt: result.heightPt }));
} else {
  worker().catch(error => {
    console.error(error && error.stack ? error.stack : error);
    process.exit(1);
  });
}
