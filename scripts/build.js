// Builds docs/index.html — a self-contained copy of sop.html with all
// images embedded as base64 data URIs so the file can be shared standalone
// and hosted on GitHub Pages. Referenced PDFs are copied into docs/ so
// data-pdf links work when served from the pages URL.
//
// Usage: node scripts/build.js

const fs   = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SRC  = path.join(ROOT, 'sop.html');
const DOCS = path.join(ROOT, 'docs');
const OUT  = path.join(DOCS, 'index.html');

let html = fs.readFileSync(SRC, 'utf8');

fs.mkdirSync(DOCS, { recursive: true });

// Embed images as base64 data URIs
const seen = new Set();
const srcRe = /src="([^"]+)"/g;
const replacements = [];
let m;

while ((m = srcRe.exec(html)) !== null) {
  const src = m[1];
  if (src.startsWith('http') || src.startsWith('data:') || seen.has(src)) continue;
  seen.add(src);

  const imgPath = path.join(ROOT, src);
  if (!fs.existsSync(imgPath)) {
    console.warn(`  Warning: not found — ${src}`);
    continue;
  }

  const ext  = path.extname(src).slice(1).toLowerCase();
  const mime = (ext === 'jpg' || ext === 'jpeg') ? 'image/jpeg' : 'image/png';
  const b64  = fs.readFileSync(imgPath).toString('base64');
  replacements.push([`src="${src}"`, `src="data:${mime};base64,${b64}"`]);
  console.log(`  Embedded: ${src}`);
}

for (const [from, to] of replacements) {
  html = html.split(from).join(to);
}

fs.writeFileSync(OUT, html, 'utf8');
console.log(`\nOK: docs/index.html`);

// Copy PDFs referenced by data-pdf attributes into docs/
const pdfRe = /data-pdf="([^"]+)"/g;
const pdfsSeen = new Set();
while ((m = pdfRe.exec(html)) !== null) {
  const rel = m[1];
  if (pdfsSeen.has(rel)) continue;
  pdfsSeen.add(rel);

  const src = path.join(ROOT, rel);
  const dst = path.join(DOCS, path.basename(rel));
  if (!fs.existsSync(src)) {
    console.warn(`  Warning: PDF not found — ${rel}`);
    continue;
  }
  fs.copyFileSync(src, dst);
  console.log(`  Copied PDF: ${rel}`);
}
