// Regression tests for the MathType-fit render path (worker protocol).
// Covers the issues that previously slipped through because the Node smoke
// test does not enable mathTypeFit:
//   1. cases/matrix/aligned mtable rows must not overlap after fraction
//      re-layout (row heights recomputed from post-layout ink boxes)
//   2. \\[len] / \\* modifiers are consumed, not rendered as literal text;
//      [len] becomes extra pile gap
//   3. plain single-line renders stay stable
//
// Usage: node tools/mathjax/test_mathtype_fit.cjs
const { spawn } = require("child_process");
const path = require("path");

const WORKER = path.join(__dirname, "render_mathjax_svg.cjs");
const FONT_PT = 12.0;

function b64(s) {
  return Buffer.from(s, "utf8").toString("base64");
}

function makeRequest(id, latex) {
  return {
    id,
    latexBase64: b64(latex),
    fontPt: FONT_PT,
    exRatio: 0.431,
    paddingPt: 2.3,
    maxWidthPt: 400,
    mathTypeFit: true
  };
}

const CASES = "f(x)=\\begin{cases}\\frac{1}{2} & x>0 \\\\ \\frac{1}{3} & x<0\\end{cases}";
const PILE = "\\frac{1}{2} \\\\ \\frac{1}{3}";
const MATRIX = "A=\\begin{matrix}\\frac{1}{2} & a \\\\ \\frac{1}{3} & b\\end{matrix}";
const BRK_PLAIN = "a+b \\\\ c+d";
const BRK_LEN = "a+b \\\\[24pt] c+d";
const BRK_STAR = "a+b \\\\* c+d";
const SINGLE = "\\frac{a+b}{a\\times b}";

const requests = [
  ["cases", CASES],
  ["pile", PILE],
  ["matrix", MATRIX],
  ["brkPlain", BRK_PLAIN],
  ["brkLen", BRK_LEN],
  ["brkStar", BRK_STAR],
  ["single", SINGLE]
];

const worker = spawn(process.execPath, [WORKER, "--worker"], {
  cwd: path.join(__dirname, "../.."),
  stdio: ["pipe", "pipe", "ignore"]
});

const results = {};
let nextId = 0;
let buffer = "";

worker.stdout.on("data", chunk => {
  buffer += chunk;
  let idx;
  while ((idx = buffer.indexOf("\n")) >= 0) {
    const line = buffer.slice(0, idx);
    buffer = buffer.slice(idx + 1);
    if (!line.trim()) continue;
    const resp = JSON.parse(line);
    results[resp.id] = resp;
  }
});

function sendAll() {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error("worker timed out")), 120000);
    const check = setInterval(() => {
      if (Object.keys(results).length === requests.length) {
        clearInterval(check);
        clearTimeout(timeout);
        resolve();
      }
    }, 50);
    requests.forEach(([key, latex], i) => {
      const req = makeRequest(i + 1, latex);
      req.key = undefined;
      worker.stdin.write(JSON.stringify(req) + "\n");
    });
  });
}

let failures = 0;
function check(name, cond, detail) {
  if (cond) {
    console.log(`PASS  ${name}`);
  } else {
    failures += 1;
    console.log(`FAIL  ${name}  (${detail})`);
  }
}

sendAll().then(() => {
  worker.stdin.end();
  const r = {};
  requests.forEach(([key], i) => {
    r[key] = results[i + 1];
  });
  for (const [key] of requests) {
    check(`${key} renders ok`, r[key] && r[key].ok, JSON.stringify(r[key]));
  }

  // #1: environment rows must reach pile-like clearance, not collide.
  check("cases rows do not overlap (height >= 0.9x pile)",
    r.cases.heightPt >= r.pile.heightPt * 0.9,
    `cases=${r.cases.heightPt.toFixed(1)} pile=${r.pile.heightPt.toFixed(1)}`);
  check("matrix rows do not overlap (height >= 0.9x pile)",
    r.matrix.heightPt >= r.pile.heightPt * 0.9,
    `matrix=${r.matrix.heightPt.toFixed(1)} pile=${r.pile.heightPt.toFixed(1)}`);

  // #2: \\[24pt] adds ~24pt of pile gap and leaks no literal text (width
  // stays the same as the plain pile), \\* behaves like plain.
  const gapGain = r.brkLen.heightPt - r.brkPlain.heightPt;
  check("\\\\[24pt] adds 24pt +/- 4pt of height",
    Math.abs(gapGain - 24) <= 4,
    `gain=${gapGain.toFixed(1)}pt`);
  check("\\\\[24pt] leaks no literal text (width ~ plain)",
    Math.abs(r.brkLen.widthPt - r.brkPlain.widthPt) <= 2,
    `len=${r.brkLen.widthPt.toFixed(1)} plain=${r.brkPlain.widthPt.toFixed(1)}`);
  check("\\\\* matches plain pile geometry",
    Math.abs(r.brkStar.heightPt - r.brkPlain.heightPt) <= 1
      && Math.abs(r.brkStar.widthPt - r.brkPlain.widthPt) <= 1,
    `star=${r.brkStar.heightPt.toFixed(1)}x${r.brkStar.widthPt.toFixed(1)}`
      + ` plain=${r.brkPlain.heightPt.toFixed(1)}x${r.brkPlain.widthPt.toFixed(1)}`);

  // #3: single-line stability anchor (12pt fit font).
  check("single fraction height is stable (~31pt +/- 2pt)",
    Math.abs(r.single.heightPt - 31.2) <= 2,
    `single=${r.single.heightPt.toFixed(1)}`);

  console.log(failures === 0 ? "ALL PASS" : `${failures} FAILURES`);
  process.exit(failures === 0 ? 0 : 1);
}).catch(err => {
  console.error("FAIL", err.message);
  process.exit(1);
});
