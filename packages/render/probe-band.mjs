import { readFileSync } from "node:fs";

const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
const [file, x0, y0, x1, y1] = process.argv.slice(2);
const [L, T, R, B] = [x0, y0, x1, y1].map(Number);

const doc = await pdfjs.getDocument({
  data: new Uint8Array(readFileSync(file)),
  useSystemFonts: false,
}).promise;

const page = await doc.getPage(1);
const [, , , H] = page.view;
for (const it of (await page.getTextContent()).items) {
  if (!it.str.trim()) continue;
  const [a, , , d, e, f] = it.transform;
  const top = H - f;
  if (e < L || e > R || top < T || top > B) continue;
  console.log(
    `TEXT ${JSON.stringify(it.str)} left=${e.toFixed(2)} baseline=${top.toFixed(2)} size=${Math.abs(d).toFixed(2)} width=${it.width.toFixed(2)} right=${(e + it.width).toFixed(2)}`,
  );
}
