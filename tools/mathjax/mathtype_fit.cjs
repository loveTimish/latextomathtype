#!/usr/bin/env node
"use strict";

/**
 * MathType-geometry re-layout for MathJax SVG output.
 *
 * MathJax renders \frac with script-size slots (scale 0.707) and tight
 * vertical gaps. MathType OLE previews use full-size slots with wider gaps.
 * This module rewrites the SVG produced by MathJax so fractions (and the
 * stretchy delimiters around them) follow measured MathType geometry.
 *
 * All measurements are per-em of the surrounding font size and come from
 * measurements of 421 genuine MathType WMF previews (structure_params.json).
 *
 * Coordinate notes: MathJax SVG root contains <g transform="scale(1,-1)">;
 * inside, y points UP and the math baseline is y=0. The math axis sits at
 * ~0.25em above the baseline. Glyph <path> data uses absolute commands.
 */

// Measured MathType fraction parameters (per-em). See docs/preview-display-calibration.md
const DEFAULT_PARAMS = {
  axisEm: 0.25,       // math axis height above baseline
  numGapEm: 0.384,    // numerator baseline -> fraction bar center
  denGapEm: 1.036,    // fraction bar center -> denominator baseline
  overhangEm: 0.09,   // bar overhang beyond the wider slot, per side
  delimMarginEm: 0.10, // delimiter extension beyond content, per side
                        // (GT varies 0.05-0.13em by equation; 0.10 minimizes
                        // worst-case error across the corpus)
  delimSpaceScale: 0.3, // gap scale adjacent to tall delimiters
  lineGapEm: 0.5,     // vertical clearance between stacked \\ lines (fresh
                        // MathType piles space lines generously)
  numDipAllowEm: 0.35,  // numerator content may dip this far below its baseline
                        // (descenders 0.21, text parens 0.25) before the slot is
                        // pushed up to clear the bar (nested fracs dip ~0.9)
  denRiseAllowEm: 0.85, // denominator content may rise this far above its
                        // baseline (cap 0.64, text parens 0.75) before the slot
                        // is pushed down to clear the bar (nested fracs ~1.1)
  slotScale: 1.0,     // MathType uses full-size numerator/denominator
  moScaleX: 1.0,      // horizontal compression for operators (disabled: with
  miScaleX: 1.0,      // spacing fixed, MJ glyph widths already match GT)
  spaceScale: 1.0     // keep MathJax's natural inter-atom spacing. Compressing
                      // this made linear formulas visibly crowded in Word.
};

// Ordinary operators retain MathJax's natural advance and spacing. Structure
// fitting (fractions, piles and tall delimiters) remains independent below.
const DEFAULT_SPACE_BY_C = {};

const DEFAULT_GLYPH_SCALE_BY_C = {};

// ---------------------------------------------------------------------------
// Minimal XML parse/serialize for MathJax SVG. Text nodes are preserved on
// their owning <text> element so Unicode glyphs survive fit-mode rewrites.
// ---------------------------------------------------------------------------

function parseAttrs(src) {
  const attrs = {};
  const re = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
  let m;
  while ((m = re.exec(src)) !== null) {
    attrs[m[1]] = m[2];
  }
  return attrs;
}

function parseXml(svg) {
  const root = { tag: null, attrs: {}, children: [], parent: null, text: "" };
  let current = root;
  const tagRe = /<(\/?)([\w:.-]+)((?:"[^"]*"|[^>"'])*?)(\/?)>/g;
  let m;
  let cursor = 0;
  while ((m = tagRe.exec(svg)) !== null) {
    const rawText = svg.slice(cursor, m.index);
    if (current.tag === "text" && rawText.length > 0) {
      current.text += rawText;
    }
    cursor = tagRe.lastIndex;
    const closing = m[1] === "/";
    const name = m[2];
    const attrSrc = m[3] || "";
    const selfClose = m[4] === "/";
    if (name.startsWith("?") || name.startsWith("!")) {
      continue;
    }
    if (closing) {
      if (current.parent && current.tag === name) {
        current = current.parent;
      }
      continue;
    }
    const el = { tag: name, attrs: parseAttrs(attrSrc), children: [], parent: current, text: "" };
    current.children.push(el);
    if (!selfClose) {
      current = el;
    }
  }
  return root.children[0] || null;
}

function serializeAttrs(attrs) {
  return Object.keys(attrs)
    .map(k => ` ${k}="${attrs[k]}"`)
    .join("");
}

const SELF_CLOSING = new Set(["path", "rect", "circle", "ellipse", "line", "use"]);

function serialize(el, out) {
  const attrs = serializeAttrs(el.attrs);
  if (SELF_CLOSING.has(el.tag) || (el.children.length === 0 && !el.text)) {
    out.push(`<${el.tag}${attrs}/>`);
    return;
  }
  out.push(`<${el.tag}${attrs}>`);
  if (el.text) {
    out.push(el.text);
  }
  for (const child of el.children) {
    serialize(child, out);
  }
  out.push(`</${el.tag}>`);
}

// ---------------------------------------------------------------------------
// Affine transforms: matrix [a, b, c, d, e, f] maps (x, y) ->
// (a*x + c*y + e, b*x + d*y + f).
// ---------------------------------------------------------------------------

const IDENT = [1, 0, 0, 1, 0, 0];

function multiply(m1, m2) {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5]
  ];
}

function applyTo(m, x, y) {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

function parseTransform(src) {
  if (!src) {
    return IDENT.slice();
  }
  let m = IDENT.slice();
  const re = /(matrix|translate|scale)\s*\(([^)]*)\)/g;
  let match;
  while ((match = re.exec(src)) !== null) {
    const nums = match[2].split(/[\s,]+/).filter(t => t.length > 0).map(Number);
    let t = IDENT.slice();
    if (match[1] === "translate") {
      t = [1, 0, 0, 1, nums[0] || 0, nums.length > 1 ? nums[1] : 0];
    } else if (match[1] === "scale") {
      t = [nums[0] || 1, 0, 0, nums.length > 1 ? nums[1] : (nums[0] || 1), 0, 0];
    } else if (match[1] === "matrix" && nums.length === 6) {
      t = nums.slice();
    }
    m = multiply(m, t);
  }
  return m;
}

function formatNumber(v) {
  const r = Math.round(v * 1000) / 1000;
  return Object.is(r, -0) ? "0" : String(r);
}

function translateTransform(tx, ty) {
  return `translate(${formatNumber(tx)},${formatNumber(ty)})`;
}

// ---------------------------------------------------------------------------
// Bounding boxes.
// ---------------------------------------------------------------------------

const PARAM_COUNT = { M: 2, L: 2, C: 6, Q: 4, S: 4, T: 2, A: 7, H: 1, V: 1, Z: 0 };

function pathBBox(d) {
  const tokens = d.match(/[a-zA-Z]|[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?/g) || [];
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  let i = 0, cmd = null, cx = 0, cy = 0;
  const push = (x, y) => {
    if (x < x0) x0 = x;
    if (y < y0) y0 = y;
    if (x > x1) x1 = x;
    if (y > y1) y1 = y;
  };
  while (i < tokens.length) {
    if (/[a-zA-Z]/.test(tokens[i])) {
      cmd = tokens[i];
      i += 1;
      if (cmd === "Z" || cmd === "z") {
        continue;
      }
    }
    if (cmd === null) {
      break;
    }
    const rel = cmd >= "a" && cmd <= "z";
    const up = cmd.toUpperCase();
    const n = PARAM_COUNT[up];
    if (n === undefined) {
      break;
    }
    const vals = tokens.slice(i, i + n).map(Number);
    i += n;
    if (up === "H") {
      cx = rel ? cx + vals[0] : vals[0];
      push(cx, cy);
    } else if (up === "V") {
      cy = rel ? cy + vals[0] : vals[0];
      push(cx, cy);
    } else if (up === "A") {
      cx = rel ? cx + vals[5] : vals[5];
      cy = rel ? cy + vals[6] : vals[6];
      push(cx, cy);
    } else {
      for (let k = 0; k + 1 < n; k += 2) {
        const x = rel ? cx + vals[k] : vals[k];
        const y = rel ? cy + vals[k + 1] : vals[k + 1];
        push(x, y);
        if (k + 2 >= n) {
          cx = x;
          cy = y;
        }
      }
    }
    // implicit repeat of moveto becomes lineto
    if (cmd === "M") cmd = "L";
    if (cmd === "m") cmd = "l";
  }
  return x0 === Infinity ? null : [x0, y0, x1, y1];
}

function unionBBox(a, b) {
  if (!a) return b;
  if (!b) return a;
  return [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[2], b[2]), Math.max(a[3], b[3])];
}

function transformBBox(m, bb) {
  const pts = [
    applyTo(m, bb[0], bb[1]), applyTo(m, bb[2], bb[1]),
    applyTo(m, bb[0], bb[3]), applyTo(m, bb[2], bb[3])
  ];
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const [x, y] of pts) {
    x0 = Math.min(x0, x); y0 = Math.min(y0, y);
    x1 = Math.max(x1, x); y1 = Math.max(y1, y);
  }
  return [x0, y0, x1, y1];
}

function elementBBox(el, m) {
  const own = parseTransform(el.attrs.transform);
  const mat = multiply(m, own);
  if (el.tag === "path") {
    const bb = pathBBox(el.attrs.d || "");
    return bb ? transformBBox(mat, bb) : null;
  }
  if (el.tag === "rect") {
    const x = Number(el.attrs.x || 0), y = Number(el.attrs.y || 0);
    const w = Number(el.attrs.width || 0), h = Number(el.attrs.height || 0);
    return w > 0 && h > 0 ? transformBBox(mat, [x, y, x + w, y + h]) : null;
  }
  let bb = null;
  for (const child of el.children) {
    bb = unionBBox(bb, elementBBox(child, mat));
  }
  return bb;
}

// bbox of el's children in el's own coordinate space (ignoring el.transform)
function contentBBox(el) {
  let bb = null;
  for (const child of el.children) {
    bb = unionBBox(bb, elementBBox(child, IDENT));
  }
  return bb;
}

// ---------------------------------------------------------------------------
// Fraction re-layout.
// ---------------------------------------------------------------------------

function isMfrac(el) {
  return el.tag === "g" && el.attrs["data-mml-node"] === "mfrac";
}

function isTallDelimiter(el, emUnits) {
  if (el.tag !== "g" || el.attrs["data-mml-node"] !== "mo") {
    return false;
  }
  const bb = contentBBox(el);
  if (!bb) {
    return false;
  }
  const h = bb[3] - bb[1];
  const w = bb[2] - bb[0];
  return h > 0.85 * emUnits && h > 2 * w;
}

function relayoutMfrac(mfrac, emUnits, params) {
  const slots = [];
  let rule = null;
  for (const child of mfrac.children) {
    if (child.tag === "rect") {
      rule = child;
    } else if (child.tag === "g") {
      slots.push(child);
    }
  }
  if (!rule || slots.length < 2) {
    return; // \atop-style fraction without a rule: leave untouched
  }
  const num = slots[0];
  const den = slots[1];
  const numBB = contentBBox(num);
  const denBB = contentBBox(den);
  if (!numBB || !denBB) {
    return;
  }

  // current mfrac footprint in parent coordinates
  const mfracMat = parseTransform(mfrac.attrs.transform);
  const oldBB = mfrac.children.reduce(
    (bb, c) => unionBBox(bb, elementBBox(c, IDENT)), null);
  const oldX0 = oldBB[0];
  const oldX1 = oldBB[2];

  const em = emUnits;
  const axis = params.axisEm * em;
  const ovh = params.overhangEm * em;
  const numW = numBB[2] - numBB[0];
  const denW = denBB[2] - denBB[0];
  const barW = Math.max(numW, denW) + 2 * ovh;
  const barX0 = oldX0;
  const slotArea = barW - 2 * ovh;
  const thickness = Number(rule.attrs.height || 0);

  // rule
  rule.attrs.x = formatNumber(barX0);
  rule.attrs.y = formatNumber(axis - thickness / 2);
  rule.attrs.width = formatNumber(barW);

  // slots at full size, re-centered, MathType gaps. Slots whose content
  // extends far beyond its own baseline (nested fractions) are pushed away
  // from the bar so the content bounding box never collides with it, while
  // ordinary descenders / cap-height content keeps the measured baseline gap.
  const numTx = barX0 + ovh + (slotArea - numW) / 2 - numBB[0];
  const numDip = Math.max(0, -numBB[1] - params.numDipAllowEm * em);
  const numTy = axis + params.numGapEm * em + numDip;
  num.attrs.transform = translateTransform(numTx, numTy);
  const denTx = barX0 + ovh + (slotArea - denW) / 2 - denBB[0];
  const denRise = Math.max(0, denBB[3] - params.denRiseAllowEm * em);
  const denTy = axis - params.denGapEm * em - denRise;
  den.attrs.transform = translateTransform(denTx, denTy);

  // shift following siblings by the width change (in parent coordinates)
  const dw = (barX0 + barW) - oldX1;
  if (Math.abs(dw) > 0.001 && mfrac.parent) {
    const scaleX = mfracMat[0] || 1;
    const shift = dw * scaleX;
    const siblings = mfrac.parent.children;
    const idx = siblings.indexOf(mfrac);
    for (let i = idx + 1; i < siblings.length; i++) {
      const sib = siblings[i];
      if (sib.tag === "rect") {
        sib.attrs.x = formatNumber(Number(sib.attrs.x || 0) + shift);
      } else {
        const m = parseTransform(sib.attrs.transform);
        m[4] += shift;
        sib.attrs.transform = m[0] === 1 && m[3] === 1 && m[1] === 0 && m[2] === 0
          ? translateTransform(m[4], m[5])
          : `matrix(${m.map(formatNumber).join(",")})`;
      }
    }
  }
}

// Stretch tall delimiters (parens/brackets/braces) to cover the row content,
// aligned to the CONTENT extremes (not the math axis) with a small measured
// per-side margin, MathType style.
function stretchDelimiters(container, emUnits, params) {
  const delims = [];
  let contentBB = null;
  for (const child of container.children) {
    if (isTallDelimiter(child, emUnits)) {
      delims.push(child);
    } else {
      contentBB = unionBBox(contentBB, elementBBox(child, IDENT));
    }
  }
  if (delims.length === 0 || !contentBB) {
    return;
  }
  const margin = params.delimMarginEm * emUnits;
  const reqTop = contentBB[3] + margin;
  const reqBottom = contentBB[1] - margin;
  const reqSpan = reqTop - reqBottom;
  const reqCenter = (reqTop + reqBottom) / 2;
  if (reqSpan <= 0) {
    return;
  }
  for (const mo of delims) {
    const bb = contentBBox(mo);
    const span = bb[3] - bb[1];
    if (span <= 0) {
      continue;
    }
    const k = reqSpan / span;
    if (Math.abs(k - 1) < 0.02) {
      continue;
    }
    // mo transform is effectively a translation; work in mo-local coordinates
    const m = parseTransform(mo.attrs.transform);
    const centerLocal = (reqCenter - m[5]) / (m[3] || 1);
    const wrapper = {
      tag: "g",
      attrs: {
        transform: `translate(0,${formatNumber(centerLocal)}) scale(1,${formatNumber(k)}) translate(0,${formatNumber(-centerLocal)})`
      },
      children: mo.children,
      parent: mo
    };
    for (const c of wrapper.children) {
      c.parent = wrapper;
    }
    mo.children = [wrapper];
  }
}

function firstPathDataC(el) {
  if (el.tag === "path") {
    return el.attrs["data-c"] || null;
  }
  for (const c of el.children) {
    const v = firstPathDataC(c);
    if (v) {
      return v;
    }
  }
  return null;
}

// Apply optional caller-provided glyph/spacing adjustments while shifting
// following siblings consistently. Defaults preserve natural MathJax layout.
function compressGlyphs(container, emUnits, params) {
  const moK = params.moScaleX;
  const miK = params.miScaleX;
  const delimSp = params.delimSpaceScale;
  const gapThr = 0.12 * emUnits;
  const spaceByC = Object.assign({}, DEFAULT_SPACE_BY_C, params.spaceScaleByC || {});
  const glyphByC = Object.assign({}, DEFAULT_GLYPH_SCALE_BY_C, params.glyphScaleByC || {});
  const spaceScaleFor = (el) => {
    const c = firstPathDataC(el);
    if (c && c.toUpperCase() in spaceByC) {
      return spaceByC[c.toUpperCase()];
    }
    return params.spaceScale;
  };
  const glyphScaleFor = (el) => {
    const c = firstPathDataC(el);
    if (c && c.toUpperCase() in glyphByC) {
      return glyphByC[c.toUpperCase()];
    }
    return moK;
  };
  const children = container.children;
  let acc = 0;
  let prevRight = null;
  const shiftEl = (el, dw) => {
    if (el.tag === "rect") {
      el.attrs.x = formatNumber(Number(el.attrs.x || 0) + dw);
    } else {
      const m = parseTransform(el.attrs.transform);
      m[4] += dw;
      el.attrs.transform = m[0] === 1 && m[3] === 1 && m[1] === 0 && m[2] === 0
        ? translateTransform(m[4], m[5])
        : `matrix(${m.map(formatNumber).join(",")})`;
    }
  };
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    const bb = elementBBox(child, IDENT);
    if (!bb) {
      continue;
    }
    if (acc !== 0) {
      shiftEl(child, acc);
    }
    const curLeft = bb[0] + acc;
    const curRight = bb[2] + acc;
    const kind = child.tag === "g" ? child.attrs["data-mml-node"] : null;
    const isDelim = kind === "mo" && isTallDelimiter(child, emUnits);
    const isMo = kind === "mo" && !isDelim;
    const k = isMo ? glyphScaleFor(child) : kind === "mi" ? miK : 1;
    const sp = isMo ? spaceScaleFor(child) : isDelim ? delimSp : 1;
    if (k >= 0.999 && sp >= 0.999) {
      prevRight = curRight;
      continue;
    }
    let selfShift = 0;
    let extra = 0;
    if (sp < 0.999 && prevRight !== null) {
      const gapL = curLeft - prevRight;
      if (gapL > gapThr) {
        selfShift += gapL * (sp - 1);
      }
    }
    if (selfShift !== 0) {
      shiftEl(child, selfShift);
    }
    const w = bb[2] - bb[0];
    if (k < 0.999 && w > 1) {
      const m = parseTransform(child.attrs.transform);
      const x0Local = (bb[0] + acc + selfShift - m[4]) / (m[0] || 1);
      const wrapper = {
        tag: "g",
        attrs: {
          transform: `translate(${formatNumber(x0Local)},0) scale(${formatNumber(k)},1) translate(${formatNumber(-x0Local)},0)`
        },
        children: child.children,
        parent: child
      };
      for (const c of wrapper.children) {
        c.parent = wrapper;
      }
      child.children = [wrapper];
      extra += w * (k - 1);
    }
    const newRight = curRight + selfShift + extra;
    if (sp < 0.999 && i + 1 < children.length) {
      const nextBB = elementBBox(children[i + 1], IDENT);
      if (nextBB) {
        // shift-invariant original gap; dR moves following siblings only
        const gapR = nextBB[0] - bb[2];
        if (gapR > gapThr) {
          extra += gapR * (sp - 1);
        }
      }
    }
    acc += selfShift + extra;
    prevRight = newRight;
  }
}

function isMtable(el) {
  return el.tag === "g" && el.attrs["data-mml-node"] === "mtable";
}

// Re-stack mtable rows (cases/matrix/aligned) after fraction re-layout.
// MathJax computed row positions for script-size fractions; once slots are
// full size the stale row baselines collide. Rows are re-placed bottom-to-
// top from their post-layout ink boxes with lineGapEm clearance, keeping
// the table's vertical center fixed so it stays aligned to the math axis.
function relayoutMtable(mtable, emUnits, params) {
  const rows = mtable.children.filter(
    c => c.tag === "g" && (c.attrs["data-mml-node"] === "mtr"
      || c.attrs["data-mml-node"] === "mlabeledtr"));
  if (rows.length < 2) {
    return;
  }
  const bbs = rows.map(r => contentBBox(r));
  if (bbs.some(b => !b)) {
    return;
  }
  const tys = rows.map(r => parseTransform(r.attrs.transform)[5]);
  const gap = params.lineGapEm * emUnits;

  const oldTop = Math.max(...rows.map((r, i) => tys[i] + bbs[i][3]));
  const oldBot = Math.min(...rows.map((r, i) => tys[i] + bbs[i][1]));

  // anchor on the first row, stack downward (y-up: below = smaller y)
  const newTys = [tys[0]];
  let cursor = tys[0] + bbs[0][1];
  for (let i = 1; i < rows.length; i++) {
    const ty = (cursor - gap) - bbs[i][3];
    newTys.push(ty);
    cursor = ty + bbs[i][1];
  }
  const newTop = newTys[0] + bbs[0][3];
  const newBot = cursor;
  const shift = ((oldTop + oldBot) - (newTop + newBot)) / 2;

  for (let i = 0; i < rows.length; i++) {
    const ty = newTys[i] + shift;
    if (Math.abs(ty - tys[i]) <= 0.001) {
      continue;
    }
    const m = parseTransform(rows[i].attrs.transform);
    m[5] = ty;
    rows[i].attrs.transform = m[0] === 1 && m[3] === 1 && m[1] === 0 && m[2] === 0
      ? translateTransform(m[4], m[5])
      : `matrix(${m.map(formatNumber).join(",")})`;
  }
}

function walk(el, emUnits, params) {
  for (const child of el.children) {
    walk(child, emUnits, params);
  }
  if (isMfrac(el)) {
    relayoutMfrac(el, emUnits, params);
  } else if (el.tag === "g") {
    if (isMtable(el)) {
      relayoutMtable(el, emUnits, params);
    }
    compressGlyphs(el, emUnits, params);
    stretchDelimiters(el, emUnits, params);
  }
}

// ---------------------------------------------------------------------------
// Entry point.
// ---------------------------------------------------------------------------

/**
 * Rewrite a MathJax SVG string with MathType fraction geometry.
 *
 * @param {string} svg     MathJax SVG (before width/height/viewBox rewrite)
 * @param {number} emUnits viewBox units per em of the surrounding font
 * @param {object} params  optional overrides of DEFAULT_PARAMS
 * @returns {{svg: string, bbox: number[]}} rewritten svg and the content
 *          bbox [x0, y0, x1, y1] in pre-flip (y-up) root coordinates
 */
function fitMathType(svg, emUnits, params) {
  const p = Object.assign({}, DEFAULT_PARAMS, params || {});
  const root = parseXml(svg);
  if (!root || root.tag !== "svg") {
    throw new Error("mathtype_fit: not an svg document");
  }
  const contentRoot = root.children.find(c => c.tag === "g");
  if (!contentRoot) {
    throw new Error("mathtype_fit: no content group");
  }
  walk(contentRoot, emUnits, p);
  const bbox = contentBBox(contentRoot) || [0, 0, 1, 1];
  const out = [];
  serialize(root, out);
  return { svg: out.join(""), bbox };
}

/**
 * Stack several already-fitted line SVGs vertically (MathType pile from a
 * top-level `\\` line break). Lines are left-aligned on their ink boxes with
 * a box-to-box clearance of params.lineGapEm. The returned bbox uses the
 * FIRST line's baseline as the pile baseline reference.
 *
 * @param {Array<{svg: string, bbox: number[]}>} parts fitted line results
 * @param {number} emUnits viewBox units per em
 * @param {object} params optional overrides of DEFAULT_PARAMS
 * @param {number[]} extraGapsEm optional per-line extra gap (em) from \\[len]
 * @returns {{svg: string, bbox: number[]}}
 */
function stackFittedLines(parts, emUnits, params, extraGapsEm) {
  const p = Object.assign({}, DEFAULT_PARAMS, params || {});
  if (parts.length === 1) {
    return parts[0];
  }
  const gapEmOf = (idx) => p.lineGapEm
    + (extraGapsEm && extraGapsEm[idx] ? extraGapsEm[idx] : 0);
  const roots = parts.map(part => {
    const root = parseXml(part.svg);
    if (!root || root.tag !== "svg") {
      throw new Error("mathtype_fit: line is not an svg document");
    }
    return root;
  });
  const host = roots[0];
  const hostContent = host.children.find(c => c.tag === "g");

  // save each line's content children BEFORE clearing the host (line 0's
  // content group IS hostContent, so clearing first would destroy it and
  // aliasing the same array would create a parent cycle)
  const savedChildren = roots.map(r => {
    const g = r.children.find(c => c.tag === "g");
    return g ? g.children.slice() : [];
  });
  hostContent.children = [];

  const minX0 = Math.min(...parts.map(part => part.bbox[0]));
  let cursor = null; // bottom (min y) of the stack so far
  let bbox = null;
  for (let i = 0; i < parts.length; i++) {
    const [x0, y0, x1, y1] = parts[i].bbox;
    const dx = minX0 - x0;
    const dy = cursor === null ? 0 : (cursor - gapEmOf(i) * emUnits) - y1;
    const wrapper = {
      tag: "g",
      attrs: { transform: translateTransform(dx, dy) },
      children: savedChildren[i],
      parent: hostContent
    };
    for (const c of wrapper.children) {
      c.parent = wrapper;
    }
    hostContent.children.push(wrapper);
    const placed = [x0 + dx, y0 + dy, x1 + dx, y1 + dy];
    bbox = unionBBox(bbox, placed);
    cursor = placed[1];
  }
  const out = [];
  serialize(host, out);
  return { svg: out.join(""), bbox };
}

module.exports = { fitMathType, stackFittedLines, DEFAULT_PARAMS };
