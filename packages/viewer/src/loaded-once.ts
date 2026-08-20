/**
 * Runs `load` when it is first asked and hands every later caller the promise it
 * answered with, so that what belongs to the whole page is fetched once however
 * many components ask for it.
 *
 * **A failure is not kept.** A promise that rejected would otherwise answer every
 * later mount with the same stale error, and nothing short of reloading the page
 * would ask again, so one lost request would disable the component for good.
 */
export function loadedOnce<T>(load: () => Promise<T>): () => Promise<T> {
  let kept: Promise<T> | null = null;

  return () => {
    if (kept !== null) return kept;

    const loading = load();
    kept = loading;
    loading.catch(() => {
      if (kept === loading) kept = null;
    });
    return loading;
  };
}
