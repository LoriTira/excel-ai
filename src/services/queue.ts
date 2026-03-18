const MAX_CONCURRENT = 3;
const MAX_QUEUE_DEPTH = 100;

let activeCount = 0;
const pending: Array<() => void> = [];

export async function enqueue<T>(fn: () => Promise<T>): Promise<T> {
  if (pending.length >= MAX_QUEUE_DEPTH) {
    throw new Error("Too many pending requests");
  }

  if (activeCount >= MAX_CONCURRENT) {
    await new Promise<void>((resolve) => {
      pending.push(resolve);
    });
  }

  activeCount++;
  try {
    return await fn();
  } finally {
    activeCount--;
    const next = pending.shift();
    if (next) next();
  }
}
