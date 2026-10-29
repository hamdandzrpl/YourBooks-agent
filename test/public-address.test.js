import assert from "node:assert/strict";
import test from "node:test";
import { adoptOwnerOrigin, resolveConsoleAddresses } from "../src/public-address.js";

test("public URL configuration is used without an external lookup", async () => {
  let lookups = 0;
  const result = await resolveConsoleAddresses({
    port: 25565,
    dashboardUrl: "https://agent.example.com/dashboard",
    healthUrl: "https://agent.example.com/healthz",
  }, {
    fetchImpl: async () => { lookups += 1; throw new Error("should not run"); },
  });

  assert.equal(result.dashboardUrl, "https://agent.example.com/dashboard");
  assert.equal(result.discovered, false);
  assert.equal(lookups, 0);
});

test("console address lookup combines a discovered IPv4 with the allocated port", async () => {
  const result = await resolveConsoleAddresses({ port: 25565, dashboardUrl: null }, {
    fetchImpl: async () => ({ ok: true, text: async () => "198.51.100.42\n" }),
  });

  assert.equal(result.discovered, true);
  assert.equal(result.dashboardUrl, "http://198.51.100.42:25565/dashboard");
  assert.equal(result.healthUrl, "http://198.51.100.42:25565/healthz");
});

test("console address lookup brackets discovered IPv6 and falls back safely on failures", async () => {
  const ipv6 = await resolveConsoleAddresses({ port: 25565, dashboardUrl: null }, {
    fetchImpl: async () => ({ ok: true, text: async () => "2001:db8::1" }),
  });
  assert.equal(ipv6.dashboardUrl, "http://[2001:db8::1]:25565/dashboard");

  const failed = await resolveConsoleAddresses({ port: 25565, dashboardUrl: null }, {
    fetchImpl: async () => { throw new Error("network unavailable"); },
  });
  assert.equal(failed.dashboardUrl, null);
  assert.match(failed.error, /network unavailable/);

  const invalid = await resolveConsoleAddresses({ port: 25565, dashboardUrl: null }, {
    fetchImpl: async () => ({ ok: true, text: async () => "not-an-ip" }),
  });
  assert.equal(invalid.dashboardUrl, null);
  assert.match(invalid.error, /invalid address/);
});

const allow = () => true;

test("the owner's dashboard address becomes the origin when nothing is pinned", () => {
  const config = { publicOrigin: null, publicOriginExplicit: false, dashboardUrl: null, healthUrl: null };
  adoptOwnerOrigin(config, { headers: { host: "203.0.113.10:25565" } }, "/dashboard", allow);
  assert.equal(config.publicOrigin, "http://203.0.113.10:25565");
  assert.equal(config.dashboardUrl, "http://203.0.113.10:25565/dashboard");
  assert.equal(config.healthUrl, "http://203.0.113.10:25565/healthz");
});

test("a proxy's https scheme and forwarded host are honoured", () => {
  const config = { publicOrigin: null, publicOriginExplicit: false };
  adoptOwnerOrigin(config, {
    headers: { host: "internal:3000", "x-forwarded-host": "panel.example.test", "x-forwarded-proto": "https" },
  }, "/dashboard", allow);
  assert.equal(config.publicOrigin, "https://panel.example.test");
});

test("an explicit origin is never replaced, and only /dashboard teaches one", () => {
  const pinned = { publicOrigin: "https://pinned.example.test", publicOriginExplicit: true };
  adoptOwnerOrigin(pinned, { headers: { host: "203.0.113.10:25565" } }, "/dashboard", allow);
  assert.equal(pinned.publicOrigin, "https://pinned.example.test");

  const other = { publicOrigin: null, publicOriginExplicit: false };
  adoptOwnerOrigin(other, { headers: { host: "203.0.113.10:25565" } }, "/healthz", allow);
  assert.equal(other.publicOrigin, null);
});

test("an anonymous caller cannot choose where workspace links point", () => {
  // The links carry a bearer token, so a spoofed Host must never be adopted.
  const config = { publicOrigin: null, publicOriginExplicit: false };
  adoptOwnerOrigin(config, { headers: { host: "attacker.example" } }, "/dashboard", () => false);
  assert.equal(config.publicOrigin, null);
  // Nor can a caller who never logged in, when no checker is supplied at all.
  adoptOwnerOrigin(config, { headers: { host: "attacker.example" } }, "/dashboard");
  assert.equal(config.publicOrigin, null);
});

test("loopback and junk hosts are ignored", () => {
  for (const host of ["127.0.0.1:3000", "localhost:3000", "[::1]:3000", "", "evil host", "a<b"]) {
    const config = { publicOrigin: null, publicOriginExplicit: false };
    adoptOwnerOrigin(config, { headers: { host } }, "/dashboard", allow);
    assert.equal(config.publicOrigin, null, host);
  }
});