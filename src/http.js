export async function fetchJsonWithTimeout(fetchImpl, url, options, timeoutMs) {
  const controller = new AbortController();
  let timeoutId;
  const request = (async () => {
    const response = await fetchImpl(url, { ...options, signal: controller.signal });
    const raw = await response.text();
    let result;
    try {
      result = JSON.parse(raw);
    } catch {
      const cleaned = raw.replace(/data:\s*\[DONE\]\s*$/i, "").trim();
      const event = cleaned.split(/\r?\n/).reverse().find((line) => line.startsWith("data: ") && line.slice(6).trim() !== "[DONE]");
      result = event ? JSON.parse(event.slice(6)) : JSON.parse(cleaned);
    }
    return { response, result };
  })();
  const timeout = new Promise((_, reject) => {
    timeoutId = setTimeout(() => {
      const error = new Error(`request timed out after ${timeoutMs} ms`);
      controller.abort(error);
      reject(error);
    }, timeoutMs);
  });

  try {
    return await Promise.race([request, timeout]);
  } finally {
    clearTimeout(timeoutId);
  }
}