#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..', '..');
const corpusPath = path.join(root, 'docs', 'reference', 'mathtype', 'latex-coverage-corpus.json');
const outputPath = path.join(root, 'src', 'main', 'resources', 'mtef', 'official-latex-symbols.tsv');

require(path.join(root, 'node_modules', 'mathjax-full', 'js', 'input', 'tex', 'base', 'BaseMappings.js'));
require(path.join(root, 'node_modules', 'mathjax-full', 'js', 'input', 'tex', 'ams', 'AmsMappings.js'));
const {MapHandler} = require(path.join(
  root, 'node_modules', 'mathjax-full', 'js', 'input', 'tex', 'MapHandler.js'
));

const mapNames = [
  'mathchar0mi',
  'mathchar0mo',
  'mathchar7',
  'AMSsymbols-mathchar0mi',
  'AMSsymbols-mathchar0mo',
  'delimiter',
  'AMSsymbols-delimiter',
];
const maps = mapNames.map((name) => MapHandler.getMap(name)).filter(Boolean);
const corpus = JSON.parse(fs.readFileSync(corpusPath, 'utf8'));

function typeface(codePoint) {
  if (codePoint >= 0x03b1 && codePoint <= 0x03ff) return 4;
  if (codePoint >= 0x0391 && codePoint <= 0x03ab) return 5;
  if ((codePoint >= 0x41 && codePoint <= 0x5a) || (codePoint >= 0x61 && codePoint <= 0x7a)) return 3;
  return 6;
}

const rows = [];
for (const entry of corpus.entries) {
  if (!entry.command.startsWith('\\') || entry.command.startsWith('\\begin{')) continue;
  const name = entry.command.slice(1);
  let symbol;
  for (const map of maps) {
    symbol = map.lookup(name) || map.lookup(entry.command);
    if (symbol) break;
  }
  if (!symbol || typeof symbol.char !== 'string') continue;
  const points = Array.from(symbol.char);
  if (points.length !== 1) continue;
  const codePoint = points[0].codePointAt(0);
  rows.push(`${entry.command}\t${typeface(codePoint)}\t${codePoint.toString(16).toUpperCase()}`);
}

fs.mkdirSync(path.dirname(outputPath), {recursive: true});
fs.writeFileSync(outputPath, [
  '# Generated from MathJax 3.2.2 mappings for the pinned WIRIS coverage corpus.',
  '# latex-command<TAB>mtef-typeface<TAB>unicode-code-point',
  ...rows,
  '',
].join('\n'));
console.log(`wrote ${rows.length} symbol mappings to ${outputPath}`);
