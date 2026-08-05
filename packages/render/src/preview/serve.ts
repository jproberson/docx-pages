import { createReadStream, existsSync, statSync } from "node:fs";
import { createServer } from "node:http";
import { extname, join, normalize, resolve } from "node:path";

const ROOT = resolve("samples");
const PORT = Number(process.env["DOCX_PAGES_PREVIEW_PORT"] ?? 8787);
const START = "/preview/index.html";

const TYPES: Readonly<Record<string, string>> = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".pdf": "application/pdf",
  ".ttf": "font/ttf",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
};

// A preview is rewritten every time it is looked at, so a page may never be held
// on to: a cached one is the last change silently missing. Everything else has to
// stay storable, since Chrome's own pdf viewer cannot read a document it is not
// allowed to keep.
const NO_STORE = "no-store, must-revalidate";
const REVALIDATE = "no-cache";

const cacheControl = (path: string): string =>
  extname(path) === ".html" || extname(path) === ".css" ? NO_STORE : REVALIDATE;

// The pdf viewer asks for the document a range at a time.
function rangeOf(header: string | undefined, size: number): readonly [number, number] | null {
  const found = /^bytes=(\d*)-(\d*)$/.exec(header ?? "");
  if (found === null) return null;

  const [, from = "", to = ""] = found;
  const start = from === "" ? size - Number(to) : Number(from);
  const end = from === "" || to === "" ? size - 1 : Number(to);
  if (!Number.isFinite(start) || start < 0 || start > end || end >= size) return null;
  return [start, end];
}

const pathOf = (url: string): string => {
  const asked = decodeURIComponent(new URL(url, "http://localhost").pathname);
  const path = normalize(join(ROOT, asked === "/" ? START : asked));
  return path.startsWith(ROOT) ? path : ROOT;
};

createServer((request, response) => {
  const path = pathOf(request.url ?? "/");

  if (!existsSync(path) || statSync(path).isDirectory()) {
    response.writeHead(404, { "cache-control": NO_STORE });
    response.end("not found\n");
    return;
  }

  const size = statSync(path).size;
  const headers = {
    "content-type": TYPES[extname(path)] ?? "application/octet-stream",
    "cache-control": cacheControl(path),
    "accept-ranges": "bytes",
  };

  const range = rangeOf(request.headers.range, size);
  if (range === null) {
    response.writeHead(200, { ...headers, "content-length": size });
    createReadStream(path).pipe(response);
    return;
  }

  const [start, end] = range;
  response.writeHead(206, {
    ...headers,
    "content-length": end - start + 1,
    "content-range": `bytes ${String(start)}-${String(end)}/${String(size)}`,
  });
  createReadStream(path, { start, end }).pipe(response);
}).listen(PORT, () => {
  process.stdout.write(`http://localhost:${String(PORT)}${START}\n`);
});
