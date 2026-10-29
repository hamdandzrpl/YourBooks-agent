export function createSerialQueue() {
  const pending = [];
  let running = false;

  async function drain() {
    if (running) return;
    running = true;
    try {
      while (pending.length > 0) {
        const item = pending.shift();
        try {
          item.resolve(await item.task());
        } catch (error) {
          item.reject(error);
        }
      }
    } finally {
      running = false;
      if (pending.length > 0) void drain();
    }
  }

  return {
    enqueue(task) {
      if (typeof task !== "function") return Promise.reject(new Error("Queue item must be a function"));
      const result = new Promise((resolve, reject) => pending.push({ task, resolve, reject }));
      void drain();
      return result;
    },
    get pendingCount() {
      return pending.length;
    },
  };
}

export function createOutboundRateLimiter({ minIntervalMs = 5000, send, now = Date.now, sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms)) }) {
  if (!Number.isInteger(minIntervalMs) || minIntervalMs < 0) {
    throw new Error("minIntervalMs must be a non-negative integer");
  }
  if (typeof send !== "function") throw new Error("send must be a function");

  let tail = Promise.resolve();
  let lastSentAt = 0;

  return (...args) => {
    const current = tail
      .catch(() => {})
      .then(async () => {
        const remaining = minIntervalMs - (now() - lastSentAt);
        if (lastSentAt > 0 && remaining > 0) await sleep(remaining);
        try {
          return await send(...args);
        } finally {
          lastSentAt = now();
        }
      });
    tail = current.catch(() => {});
    return current;
  };
}