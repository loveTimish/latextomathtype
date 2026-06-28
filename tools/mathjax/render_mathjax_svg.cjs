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

const tex = new TeX({
  packages: AllPackages,
  macros: {
    overarc: ["\\overset{\\frown}{#1}", 1],
    arc: ["\\overset{\\frown}{#1}", 1],
    wideparen: ["\\overset{\\frown}{#1}", 1],
    whitestar: "\\star",
    blackstar: "\\star",
    whitediamond: "\\diamond",
    underbracechar: "\\underbrace{\\hphantom{0}}"
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
  const node = html.convert(latex, { display: false, em: 16, ex: 8, containerWidth: 100000 });
  let svg = adaptor.outerHTML(node);
  const start = svg.indexOf("<svg");
  const end = svg.lastIndexOf("</svg>");
  if (start < 0 || end < 0) {
    throw new Error("MathJax did not produce SVG");
  }
  svg = svg.slice(start, end + 6)
    .replace(/currentColor/g, "#000000")
    .replace(/\s+focusable="false"/g, "")
    .replace(/\s+role="img"/g, "");

  const widthEx = Math.max(numberAttr(svg, "width"), 0.1);
  const heightEx = Math.max(numberAttr(svg, "height"), 0.1);
  const valignEx = verticalAlignEx(svg);
  const exPt = fontPt * exRatio;
  const contentWidthPt = widthEx * exPt;
  const contentHeightPt = heightEx * exPt;
  const unscaledWidthPt = contentWidthPt + paddingPt * 2;
  const unscaledHeightPt = contentHeightPt + paddingPt * 2;
  const scale = maxWidthPt > 0 ? Math.min(1, maxWidthPt / unscaledWidthPt) : 1;
  const widthPt = Math.max(unscaledWidthPt * scale, 1);
  const heightPt = Math.max(unscaledHeightPt * scale, 1);
  const depthPt = Math.max(Math.abs(valignEx) * exPt * scale + paddingPt * scale, 0);

  const vb = viewBox(svg);
  const unitsPerPt = contentWidthPt > 0 ? vb[2] / contentWidthPt : 1000 / fontPt;
  const padUnits = paddingPt * unitsPerPt;
  const paddedViewBox = [
    vb[0] - padUnits,
    vb[1] - padUnits,
    vb[2] + padUnits * 2,
    vb[3] + padUnits * 2
  ];

  svg = svg
    .replace(/\bwidth=['"][^'"]+['"]/i, `width="${widthPt.toFixed(6)}pt"`)
    .replace(/\bheight=['"][^'"]+['"]/i, `height="${heightPt.toFixed(6)}pt"`)
    .replace(/\bviewBox=['"][^'"]+['"]/i, `viewBox="${paddedViewBox.map(v => v.toFixed(3)).join(" ")}"`)
    .replace(/\s+style=['"][^'"]*['"]/i, "");

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
