import { existsSync } from "node:fs";
import { isIP } from "node:net";

if (existsSync(".env")) {
  process.loadEnvFile(".env");
}

function required(env, name) {
  const value = env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

/**
 * Accept whatever base URL shape the provider documents and produce a full
 * chat-completions endpoint.
 *
 *   https://api.openai.com/v1                    -> .../v1/chat/completions
 *   https://generativelanguage.googleapis.com/v1beta/openai -> .../chat/completions
 *   https://openrouter.ai/api/v1/                -> .../api/v1/chat/completions
 *   https://api.anthropic.com/v1/chat/completions -> unchanged
 */
export function resolveChatCompletionsUrl(baseUrl) {
  let parsed;
  try {
    parsed = new URL(baseUrl);
  } catch {
    throw new Error("AI_BASE_URL must be a valid absolute http(s) URL");
  }
  if (!["http:", "https:"].includes(parsed.protocol)) {
    throw new Error("AI_BASE_URL must use http or https");
  }
  const trimmed = baseUrl.replace(/\/+$/, "");
  if (/\/chat\/completions$/.test(trimmed)) return trimmed;
  return `${trimmed}/chat/completions`;
}

export function loadConfig(env = process.env) {
  const port = Number(env.SERVER_PORT ?? env.PORT ?? "3000");
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error("PORT must be a valid TCP port");
  }

  let dashboardOrigin = null;
  // True only when the operator told us the public address. Anything else is a
  // guess, and a guess may be replaced by the address the owner actually used
  // in the browser (see serveRequest), so workspace links stay clickable.
  let originExplicit = false;
  if (env.PUBLIC_URL?.trim()) {
    let publicUrl;
    try {
      publicUrl = new URL(env.PUBLIC_URL.trim());
    } catch {
      throw new Error("PUBLIC_URL must be a valid public HTTP or HTTPS origin");
    }
    if (!["http:", "https:"].includes(publicUrl.protocol)
      || publicUrl.username
      || publicUrl.password
      || publicUrl.pathname !== "/"
      || publicUrl.search
      || publicUrl.hash) {
      throw new Error("PUBLIC_URL must be a public HTTP or HTTPS origin without a path");
    }
    dashboardOrigin = publicUrl.origin;
    originExplicit = true;
  } else if (env.PUBLIC_IP?.trim()) {
    const publicIp = env.PUBLIC_IP.trim();
    const ipVersion = isIP(publicIp);
    if (!ipVersion) {
      throw new Error("PUBLIC_IP must be a valid IPv4 or IPv6 address");
    }
    const formattedIp = ipVersion === 6 ? `[${publicIp}]` : publicIp;
    dashboardOrigin = `http://${formattedIp}:${port}`;
    originExplicit = true;
  } else {
    const configuredHost = env.PUBLIC_HOST?.trim() || env.PTERODACTYL_SERVER_IP?.trim() || env.SERVER_IP?.trim();
    const usableHost = configuredHost && !["0.0.0.0", "::", "::0", "localhost"].includes(configuredHost.toLowerCase());
    if (usableHost) {
      const formattedHost = configuredHost.includes(":") && !configuredHost.startsWith("[") ? `[${configuredHost}]` : configuredHost;
      dashboardOrigin = `http://${formattedHost}:${port}`;
      originExplicit = true;
    } else if (!env.SERVER_PORT?.trim() && !configuredHost) {
      dashboardOrigin = `http://127.0.0.1:${port}`;
    }
  }

  const defaultCurrency = env.DEFAULT_CURRENCY?.trim().toUpperCase() || "IDR";
  try {
    new Intl.NumberFormat("en", { style: "currency", currency: defaultCurrency });
  } catch {
    throw new Error("DEFAULT_CURRENCY must be a valid ISO 4217 currency code");
  }
  const timeZone = env.TIME_ZONE?.trim() || "UTC";
  try {
    new Intl.DateTimeFormat("en", { timeZone });
  } catch {
    throw new Error("TIME_ZONE must be a valid IANA time zone");
  }
  const xkiroModel = env.AI_MODEL?.trim() || env.XKIRO_MODEL?.trim() || "sensenova/sensenova-6.8-flash-lite";
  if (!xkiroModel || /\s/.test(xkiroModel) || xkiroModel.length > 200) {
    throw new Error("AI_MODEL must be a single non-empty model identifier");
  }
  // Optional: many providers (Anthropic, Gemini, most [OI]-compatible gateways)
  // reject the request outright when they receive an unknown reasoning field,
  // so it is only sent when the operator opts in.
  const xkiroReasoningEffort = (env.AI_REASONING_EFFORT ?? env.XKIRO_REASONING_EFFORT)?.trim().toLowerCase() || "";
  if (xkiroReasoningEffort && !new Set(["none", "low", "medium", "high", "xhigh", "max"]).has(xkiroReasoningEffort)) {
    throw new Error("AI_REASONING_EFFORT must be none, low, medium, high, xhigh, or max");
  }
  const chatCompletionsUrl = resolveChatCompletionsUrl(env.AI_BASE_URL?.trim() || env.XKIRO_BASE_URL?.trim() || "https://api.xkiro.com/v1");
  const maxOutputTokens = Number(env.AI_MAX_TOKENS ?? "0");
  if (!Number.isInteger(maxOutputTokens) || maxOutputTokens < 0) {
    throw new Error("AI_MAX_TOKENS must be a non-negative integer");
  }
  const searchUrl = env.AI_SEARCH_URL?.trim() || "";
  const searchModel = env.AI_SEARCH_MODEL?.trim() || "xkiro/web-search";
  const panelApiToken = env.PANEL_API_TOKEN?.trim() || "";
  if (panelApiToken && panelApiToken.length < 32) {
    throw new Error("PANEL_API_TOKEN must contain at least 32 characters when configured");
  }
  const panelPassword = env.PANEL_PASSWORD?.trim() || "";
  if (panelPassword.length > 256) {
    throw new Error("PANEL_PASSWORD must contain no more than 256 characters");
  }
  const panelOwnerNumber = env.PANEL_OWNER_NUMBER?.trim() || "";
  if (panelOwnerNumber && !/^\d{8,15}$/.test(panelOwnerNumber)) {
    throw new Error("PANEL_OWNER_NUMBER must be a WhatsApp number containing 8-15 digits");
  }
  const maxFileMegabytes = Number(env.MAX_FILE_MB ?? "20");
  if (!Number.isInteger(maxFileMegabytes) || maxFileMegabytes < 1 || maxFileMegabytes > 20) {
    throw new Error("MAX_FILE_MB must be an integer between 1 and 20");
  }
  const routerUrl = env.ROUTER_URL?.trim() || "";
  if (routerUrl) {
    try {
      const parsed = new URL(routerUrl);
      if (!["http:", "https:"].includes(parsed.protocol) || parsed.pathname !== "/" || parsed.search || parsed.hash) {
        throw new Error("bad");
      }
    } catch {
      throw new Error("ROUTER_URL must be a valid http(s) origin without a path");
    }
  }
  const routerPort = Number(env.ROUTER_PORT ?? "20128");
  if (!Number.isInteger(routerPort) || routerPort < 1 || routerPort > 65535) {
    throw new Error("ROUTER_PORT must be an integer between 1 and 65535");
  }
  // Default ON: /9router/ works out of the box (binary ships via package.json).
  // Set ROUTER_PROXY=0/off to disable the proxy and auto-start.
  const routerProxyEnabled = routerPort > 0 && !/^(0|false|no|off)$/i.test((env.ROUTER_PROXY ?? "").trim());

  return {
    port,
    aiApiKey: env.AI_API_KEY?.trim() || required(env, "XKIRO_API_KEY"),
    chatCompletionsUrl,
    maxOutputTokens,
    searchUrl,
    searchModel,
    panelApiToken,
    panelPassword,
    panelOwnerNumber,
    aiModel: xkiroModel,
    aiReasoningEffort: xkiroReasoningEffort,
    timeZone,
    defaultCurrency,
    dashboardUrl: dashboardOrigin ? `${dashboardOrigin}/dashboard` : null,
    healthUrl: dashboardOrigin ? `${dashboardOrigin}/healthz` : null,
    publicOrigin: dashboardOrigin,
    publicOriginExplicit: originExplicit,
    routerUrl,
    routerPort,
    routerProxyEnabled,
    dataDirectory: env.DATA_DIR?.trim() || "data",
    maxFileBytes: maxFileMegabytes * 1024 * 1024,
  };
}