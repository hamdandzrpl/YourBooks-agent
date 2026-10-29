import assert from "node:assert/strict";
import test from "node:test";
import { loadConfig } from "../src/config.js";

const validEnvironment = {
  PORT: "3000",
  AI_API_KEY: "test-key",
  AI_BASE_URL: "https://api.vikey.ai/v1",
  AI_MODEL: "stealth/space-bunny-alpha:free",
  AI_REASONING_EFFORT: "low",
  TIME_ZONE: "Asia/Jakarta",
  DEFAULT_CURRENCY: "IDR",
  DATA_DIR: "data",
  MAX_FILE_MB: "20",
};

test("configuration loads finance defaults without sender restrictions", () => {
  const config = loadConfig(validEnvironment);
  assert.equal(config.aiModel, "stealth/space-bunny-alpha:free");
  assert.equal(config.aiReasoningEffort, "low");
  assert.equal(config.defaultCurrency, "IDR");
  assert.equal(config.timeZone, "Asia/Jakarta");
  assert.equal(config.maxFileBytes, 20 * 1024 * 1024);
  assert.equal(config.panelOwnerNumber, "");
  assert.equal("whatsappAllowedNumbers" in config, false);
});

test("configuration prefers the Pterodactyl allocation and builds dashboard links when its host is known", () => {
  const config = loadConfig({
    ...validEnvironment,
    SERVER_IP: "203.0.113.10",
    SERVER_PORT: "25565",
  });
  assert.equal(config.port, 25565);
  assert.equal(config.dashboardUrl, "http://203.0.113.10:25565/dashboard");
  assert.equal(config.healthUrl, "http://203.0.113.10:25565/healthz");

  const publicHostConfig = loadConfig({
    ...validEnvironment,
    SERVER_PORT: "25565",
    PUBLIC_HOST: "panel.example.test",
  });
  assert.equal(publicHostConfig.dashboardUrl, "http://panel.example.test:25565/dashboard");

  const missingHostConfig = loadConfig({
    ...validEnvironment,
    SERVER_PORT: "25565",
  });
  assert.equal(missingHostConfig.port, 25565);
  assert.equal(missingHostConfig.dashboardUrl, null);
  assert.equal(missingHostConfig.healthUrl, null);

  const wildcardHostConfig = loadConfig({
    ...validEnvironment,
    SERVER_IP: "0.0.0.0",
    SERVER_PORT: "25565",
  });
  assert.equal(wildcardHostConfig.port, 25565);
  assert.equal(wildcardHostConfig.dashboardUrl, null);

  const publicIpConfig = loadConfig({
    ...validEnvironment,
    SERVER_IP: "0.0.0.0",
    SERVER_PORT: "25565",
    PUBLIC_IP: "203.0.113.10",
  });
  assert.equal(publicIpConfig.dashboardUrl, "http://203.0.113.10:25565/dashboard");

  const publicConfig = loadConfig({
    ...validEnvironment,
    SERVER_PORT: "25565",
    PUBLIC_URL: "https://agent.example.com",
  });
  assert.equal(publicConfig.dashboardUrl, "https://agent.example.com/dashboard");

  assert.throws(() => loadConfig({
    ...validEnvironment,
    SERVER_PORT: "25565",
    PUBLIC_IP: "not-an-ip",
  }), /PUBLIC_IP/);

  const localConfig = loadConfig(validEnvironment);
  assert.equal(localConfig.dashboardUrl, "http://127.0.0.1:3000/dashboard");
});

test("configuration rejects unsupported port, timezone, currency, and model formats", () => {
  assert.throws(() => loadConfig({ ...validEnvironment, PORT: "3000xyz" }), /PORT/);
  assert.throws(() => loadConfig({ ...validEnvironment, TIME_ZONE: "Mars/Olympus" }), /TIME_ZONE/);
  assert.throws(() => loadConfig({ ...validEnvironment, DEFAULT_CURRENCY: "ID" }), /DEFAULT_CURRENCY/);
  assert.equal(loadConfig({ ...validEnvironment, AI_MODEL: "bare-model-name" }).aiModel, "bare-model-name");
  assert.throws(() => loadConfig({ ...validEnvironment, AI_REASONING_EFFORT: "ultra" }), /AI_REASONING_EFFORT/);
  assert.throws(() => loadConfig({ ...validEnvironment, PANEL_API_TOKEN: "short" }), /PANEL_API_TOKEN/);
  assert.equal(loadConfig({ ...validEnvironment, PANEL_API_TOKEN: "p".repeat(32) }).panelApiToken, "p".repeat(32));
  assert.equal(loadConfig({ ...validEnvironment, PANEL_PASSWORD: "1122334455" }).panelPassword, "1122334455");
  assert.throws(() => loadConfig({ ...validEnvironment, PANEL_PASSWORD: "p".repeat(257) }), /PANEL_PASSWORD/);
  assert.equal(loadConfig({ ...validEnvironment, PANEL_OWNER_NUMBER: "6281234567890" }).panelOwnerNumber, "6281234567890");
  assert.throws(() => loadConfig({ ...validEnvironment, PANEL_OWNER_NUMBER: "owner@example.com" }), /PANEL_OWNER_NUMBER/);
  assert.equal("ms365AllowedNumbers" in loadConfig(validEnvironment), false);
  assert.throws(() => loadConfig({ ...validEnvironment, MAX_FILE_MB: "25" }), /MAX_FILE_MB/);
  assert.throws(() => loadConfig({ ...validEnvironment, PUBLIC_URL: "javascript:alert(1)" }), /PUBLIC_URL/);
});
test("chat completions URL is derived from any documented base URL shape", () => {
  const cases = [
    ["https://api.vikey.ai/v1", "https://api.vikey.ai/v1/chat/completions"],
    ["https://api.vikey.ai/v1/", "https://api.vikey.ai/v1/chat/completions"],
    ["https://api.openai.com/v1", "https://api.openai.com/v1/chat/completions"],
    ["https://openrouter.ai/api/v1", "https://openrouter.ai/api/v1/chat/completions"],
    ["https://generativelanguage.googleapis.com/v1beta/openai", "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"],
    ["https://api.deepseek.com/v1", "https://api.deepseek.com/v1/chat/completions"],
    ["https://api.groq.com/openai/v1", "https://api.groq.com/openai/v1/chat/completions"],
    ["http://127.0.0.1:11434/v1", "http://127.0.0.1:11434/v1/chat/completions"],
    ["https://api.anthropic.com/v1/chat/completions", "https://api.anthropic.com/v1/chat/completions"],
  ];
  for (const [base, expected] of cases) {
    assert.equal(loadConfig({ ...validEnvironment, AI_BASE_URL: base }).chatCompletionsUrl, expected, base);
  }
  assert.throws(() => loadConfig({ ...validEnvironment, AI_BASE_URL: "not-a-url" }), /AI_BASE_URL/);
  assert.throws(() => loadConfig({ ...validEnvironment, AI_BASE_URL: "ftp://example.com/v1" }), /AI_BASE_URL/);
});

test("provider search stays disabled unless a search endpoint is configured", () => {
  assert.equal(loadConfig(validEnvironment).searchUrl, "");
  assert.equal(loadConfig({ ...validEnvironment, AI_SEARCH_URL: "https://api.example.com/v1/search" }).searchUrl, "https://api.example.com/v1/search");
});

test("legacy XKIRO environment names still load for backwards compatibility", () => {
  const config = loadConfig({
    PORT: "3000",
    XKIRO_API_KEY: "legacy-key",
    XKIRO_MODEL: "sensenova/sensenova-6.8-flash-lite",
    XKIRO_BASE_URL: "https://api.xkiro.com/v1",
    TIME_ZONE: "Asia/Jakarta",
    DEFAULT_CURRENCY: "IDR",
  });
  assert.equal(config.aiApiKey, "legacy-key");
  assert.equal(config.aiModel, "sensenova/sensenova-6.8-flash-lite");
  assert.equal(config.chatCompletionsUrl, "https://api.xkiro.com/v1/chat/completions");
});
