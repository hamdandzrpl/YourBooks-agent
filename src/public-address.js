import { isIP } from "node:net";

/** Host header values that can never be a public address. */
const LOOPBACK = /^(127\.0\.0\.1|localhost|\[::1\])(:|$)/i;
const SAFE_HOST = /^[a-z0-9.\-:[\]]+$/i;

/**
 * Adopt the address the owner used to open the dashboard.
 *
 * Pterodactyl often exposes only a port: PUBLIC_URL is empty, and the IP lookup
 * either fails or returns an address the allocation does not route. The owner
 * reaches the dashboard somehow, and that Host header is the one address known
 * to work, so workspace links built from it are clickable. An explicitly
 * configured origin always wins, and loopback is ignored.
 *
 * Call this only from a request that proved the caller is the owner (a
 * successful login, or an authenticated dashboard load): workspace links carry
 * a bearer token, so an anonymous caller must never choose where they point.
 *
 * Mutates and returns `config` for convenience.
 */
export function adoptOriginFromRequest(config, request) {
  if (config.publicOriginExplicit) return config;
  const host = String(request.headers["x-forwarded-host"] ?? request.headers.host ?? "").split(",")[0].trim();
  if (!host || !SAFE_HOST.test(host) || LOOPBACK.test(host)) return config;
  const forwardedProto = String(request.headers["x-forwarded-proto"] ?? "").split(",")[0].trim().toLowerCase();
  const proto = forwardedProto === "https" ? "https" : "http";
  const origin = `${proto}://${host}`;
  if (origin === config.publicOrigin) return config;
  config.publicOrigin = origin;
  config.dashboardUrl = `${origin}/dashboard`;
  config.healthUrl = `${origin}/healthz`;
  return config;
}

/** Gated wrapper for the dashboard route: adopts the origin only when the
 *  request is authenticated. */
export function adoptOwnerOrigin(config, request, pathname, isAuthorized) {
  if (pathname !== "/dashboard") return config;
  if (typeof isAuthorized !== "function" || !isAuthorized(request, config)) return config;
  return adoptOriginFromRequest(config, request);
}

export async function resolveConsoleAddresses(config, { fetchImpl = fetch, timeoutMs = 5000 } = {}) {
  if (config.dashboardUrl) {
    return { dashboardUrl: config.dashboardUrl, healthUrl: config.healthUrl, discovered: false };
  }

  try {
    const response = await fetchImpl("https://api.ipify.org", {
      signal: AbortSignal.timeout(timeoutMs),
    });
    if (!response.ok) throw new Error(`IP lookup failed (${response.status})`);
    const ip = (await response.text()).trim();
    const version = isIP(ip);
    if (!version) throw new Error("IP lookup returned an invalid address");
    const host = version === 6 ? `[${ip}]` : ip;
    const origin = `http://${host}:${config.port}`;
    return {
      dashboardUrl: `${origin}/dashboard`,
      healthUrl: `${origin}/healthz`,
      discovered: true,
    };
  } catch (error) {
    return {
      dashboardUrl: null,
      healthUrl: null,
      discovered: false,
      error: error.message,
    };
  }
}