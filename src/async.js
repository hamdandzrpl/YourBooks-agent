export function withTimeout(operation, timeoutMs, message) {
  if (typeof operation !== "function") throw new Error("Timed operation must be a function");
  if (!Number.isInteger(timeoutMs) || timeoutMs < 1) throw new Error("Timeout must be a positive integer");

  let timeoutId;
  const operationPromise = Promise.resolve().then(operation);
  const timeoutPromise = new Promise((_, reject) => {
    timeoutId = setTimeout(() => reject(new Error(message)), timeoutMs);
  });

  return Promise.race([operationPromise, timeoutPromise]).finally(() => clearTimeout(timeoutId));
}