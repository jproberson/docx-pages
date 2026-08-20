import { execFileSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, statSync } from "node:fs";
import { basename, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { CORPUS_DIRECTORY, documentsIn, identityOf } from "./sweep.js";

// Asking Word for its own pdf of every document in a corpus, which is what turns
// one from a pile that can only say a document changed into one that can say a
// document is **wrong**.
//
// Until this, the only pages pinned against Word were the eight reference
// documents, and every one of them is a single page. So nothing about page
// breaking had ever been checked against a real document: not widow control, not a
// row torn by a break, not a paragraph held to the one after it. A corpus of real
// documents with Word's own answer beside each is the first evidence of that kind
// there has been.
//
// **Nothing this writes names a document**, exactly as the sweep beside it does
// not: a pdf is named by the first twelve characters of the hash of the document's
// bytes, which is what the sweep calls the same document. The pdfs live under
// `samples/`, which is gitignored whole, and no number out of them may be
// committed.

const RENDERED_DIRECTORY = process.env["DOCX_PAGES_CORPUS_PDF"] ?? "samples/corpus/pdf";

const SCRIPT = resolve("packages/render/src/corpus/render.applescript");

// Word is handed a batch at a time rather than the whole corpus. A batch is a unit
// of loss: one that wedges Word past saving costs the documents in it and nothing
// more, and the run picks up from the next one.
const BATCH = 40;

export type Rendered = {
  readonly id: string;
  readonly outcome: "ok" | "fail" | "lost";
  readonly detail: string;
  readonly seconds: number;
};

// Word exports beside the document it opened and only into a directory it has been
// granted, so a document is staged under `samples/` under its own name and the pdf
// falls out beside it. The staged copy is thrown away and the pdf kept, since the
// document itself is already in the corpus and copying the corpus is not the point.
const stagedPath = (id: string): string => resolve(RENDERED_DIRECTORY, `${id}.docx`);
export const renderedPath = (id: string): string => resolve(RENDERED_DIRECTORY, `${id}.pdf`);

export function parseRenderings(output: string): readonly Rendered[] {
  const rendered: Rendered[] = [];
  for (const line of output.split("\n")) {
    const [outcome, detail, path] = line.trim().split("|");
    if (outcome !== "ok" && outcome !== "fail") continue;
    rendered.push({
      id: basename(path ?? "", ".docx"),
      outcome,
      detail: outcome === "ok" ? "" : (detail ?? ""),
      seconds: outcome === "ok" ? Number(detail ?? 0) : 0,
    });
  }
  return rendered;
}

function ask(paths: readonly string[]): string {
  try {
    return execFileSync("osascript", [SCRIPT, ...paths], {
      encoding: "utf8",
      maxBuffer: 64 * 1024 * 1024,
    });
  } catch {
    // A batch that never answered took Word down with it. Whatever it managed
    // before that is on disk and counted by the caller; the rest are lost.
    return "";
  }
}

const quitWord = (): void => {
  try {
    execFileSync("osascript", ["-e", 'tell application "Microsoft Word" to quit saving no'], {
      encoding: "utf8",
      timeout: 60_000,
    });
  } catch {
    // Word holding a dialog will not be asked; the next batch is given a clean one.
  }
};

const forceQuitWord = (): void => {
  quitWord();
  try {
    execFileSync("pkill", ["-x", "Microsoft Word"], { encoding: "utf8" });
  } catch {
    // Nothing of that name was running, which is the state being asked for.
  }
};

export type Wanted = { readonly id: string; readonly path: string };

// **The same document is filed under several names.** Two copies in one batch stage
// to the one path and hand Word the same document twice, which is what leaves it
// holding a file nothing later can open, and the second copy asks for a pdf the run
// has already made.
export function documentsWanted(directory: string): readonly Wanted[] {
  const already = new Set<string>();
  const wanted: Wanted[] = [];
  for (const path of documentsIn(directory)) {
    const id = identityOf(new Uint8Array(readFileSync(path)));
    if (already.has(id)) continue;
    already.add(id);
    wanted.push({ id, path });
  }
  return wanted;
}

export type RenderProgress = (done: number, total: number, failures: number) => void;

// Every document in the corpus that has no pdf yet, since a run of three quarters
// of an hour is one worth being able to stop and start again.
export function renderCorpus(directory: string, onProgress?: RenderProgress): readonly Rendered[] {
  mkdirSync(resolve(RENDERED_DIRECTORY), { recursive: true });

  const wanted = documentsWanted(directory).filter((each) => !existsSync(renderedPath(each.id)));

  const rendered: Rendered[] = [];
  let failures = 0;

  for (let at = 0; at < wanted.length; at += BATCH) {
    const batch = wanted.slice(at, at + BATCH);
    for (const each of batch) copyFileSync(each.path, stagedPath(each.id));

    const answered = new Map(
      parseRenderings(ask(batch.map((each) => stagedPath(each.id)))).map((each) => [each.id, each]),
    );

    let lost = false;
    for (const each of batch) {
      // Word's own answer is believed only so far: what settles it is whether a
      // pdf is there, since a batch cut short answers for nothing it never reached.
      const drawn = existsSync(renderedPath(each.id)) && statSync(renderedPath(each.id)).size > 0;
      const said = answered.get(each.id);
      if (drawn) {
        rendered.push(said ?? { id: each.id, outcome: "ok", detail: "", seconds: 0 });
      } else {
        lost = true;
        failures += 1;
        rendered.push(said ?? { id: each.id, outcome: "lost", detail: "no answer", seconds: 0 });
      }
      rmSync(stagedPath(each.id), { force: true });
    }

    // A lock file is left beside a document Word still holds, and it stops the
    // same document opening ever again. Nothing staged survives a batch.
    for (const each of batch)
      rmSync(resolve(RENDERED_DIRECTORY, `~$${each.id}.docx`), { force: true });

    if (lost) forceQuitWord();
    onProgress?.(Math.min(at + BATCH, wanted.length), wanted.length, failures);
  }

  quitWord();
  return rendered;
}

function main(): void {
  if (CORPUS_DIRECTORY === null) {
    process.stdout.write(
      "No corpus configured: set DOCX_PAGES_CORPUS to a directory of .docx files.\n",
    );
    return;
  }

  // A document Word still holds open stops every later one, so the run starts by
  // asking it to let go. Asked rather than killed: a Word that was killed comes
  // back offering to recover what it lost, and that is a window no `display alerts`
  // setting suppresses and no script can get past.
  quitWord();

  process.stdout.write(`Asking Word for its pdf of every document in ${CORPUS_DIRECTORY}\n`);
  const rendered = renderCorpus(CORPUS_DIRECTORY, (done, total, failures) => {
    process.stdout.write(`  ${String(done)}/${String(total)}, ${String(failures)} without a pdf\n`);
  });

  const drawn = rendered.filter((each) => each.outcome === "ok");
  const seconds = drawn.reduce((total, each) => total + each.seconds, 0);
  process.stdout.write(
    `\n${String(rendered.length)} asked for\n` +
      `  drawn     ${String(drawn.length)}\n` +
      `  refused   ${String(rendered.filter((each) => each.outcome === "fail").length)}\n` +
      `  lost      ${String(rendered.filter((each) => each.outcome === "lost").length)}\n` +
      `\nWord spent ${String(Math.round(seconds))}s drawing them, into ${RENDERED_DIRECTORY}\n`,
  );

  const refused = rendered.filter((each) => each.outcome !== "ok");
  if (refused.length > 0) {
    process.stdout.write(`\nwhat Word said about the ones it would not draw:\n`);
    const byDetail = new Map<string, number>();
    for (const each of refused) byDetail.set(each.detail, (byDetail.get(each.detail) ?? 0) + 1);
    for (const [detail, count] of [...byDetail].sort((one, other) => other[1] - one[1])) {
      process.stdout.write(`  ${String(count).padStart(4)}  ${detail}\n`);
    }
  }
}

// Compared against this module's own path: a guard naming the built `.js` never
// fires under tsx, which is how these are run, and the run then does nothing at all
// and says nothing about it.
if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
