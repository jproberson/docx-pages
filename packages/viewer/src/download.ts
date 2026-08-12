// Handing a written pdf to the browser as a file to keep.
//
// The bytes are core's and nothing here touches them. What a browser needs on top
// of them is an address to hang them on and a link to click, neither of which
// exists anywhere else in this package, and both of which are fiddly enough to be
// worth writing once.

export type DownloadOptions = {
  // What the file is called once it is saved. A name with no suffix is left as it
  // is: the caller may mean it, and appending one to a name that already ends in
  // something else would rename the file behind them.
  readonly fileName: string;
};

/**
 * Offers bytes to the browser as a pdf to download.
 *
 * Does nothing at all outside a browser, rather than throwing: server rendering
 * reaches every branch a component has, and a page that cannot offer a download
 * still has to draw.
 */
export function downloadPdf(bytes: Uint8Array, options: DownloadOptions): void {
  if (typeof document === "undefined" || typeof URL.createObjectURL !== "function") return;

  // A copy, because a Blob holds the memory it is given and a caller who writes
  // the page again would otherwise be handing the same buffer to two files.
  const blob = new Blob([bytes.slice()], { type: "application/pdf" });
  const address = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = address;
  link.download = options.fileName;
  link.rel = "noopener";
  link.click();

  // **Released on the next turn of the loop rather than now.** A click starts the
  // download without waiting for it, and a browser handed back the address it was
  // still reading from saves an empty file.
  setTimeout(() => {
    URL.revokeObjectURL(address);
  }, 0);
}
