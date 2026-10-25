'use strict';

// Preserve input order while bounding independent asynchronous work.
async function mapWithConcurrencyLimit(
  items,
  concurrencyLimit,
  mapper,
  onProgress = null
) {
  const normalizedItems = Array.isArray(items) ? items : [];

  if (normalizedItems.length === 0) {
    return [];
  }

  const results = new Array(normalizedItems.length);
  const workerCount = Math.max(
    1,
    Math.min(concurrencyLimit || 1, normalizedItems.length)
  );
  let nextIndex = 0;
  let completedCount = 0;

  async function worker() {
    while (true) {
      const currentIndex = nextIndex;
      nextIndex += 1;

      if (currentIndex >= normalizedItems.length) {
        return;
      }

      results[currentIndex] = await mapper(normalizedItems[currentIndex], currentIndex);
      completedCount += 1;

      if (typeof onProgress === 'function') {
        onProgress({
          completedCount,
          totalCount: normalizedItems.length,
          item: normalizedItems[currentIndex],
          index: currentIndex,
          result: results[currentIndex],
        });
      }
    }
  }

  await Promise.all(Array.from({ length: workerCount }, () => worker()));

  return results;
}

module.exports = { mapWithConcurrencyLimit };
