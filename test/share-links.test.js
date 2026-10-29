import { ShareLinkStore } from "../src/share-links.js";
import assert from "node:assert";
import { rmSync, mkdirSync } from "node:fs";
import { join } from "node:path";

const dir = join("data-test", "share-links");
rmSync(dir, { recursive: true, force: true });
mkdirSync(dir, { recursive: true });
const userA = "6281234567890";
const userB = "6289998887770";
let store;

try {
  store = new ShareLinkStore({ dataDirectory: dir });

  // Issue + resolve round-trip
  const share = store.issueToken(userA, "documents/dokumen-a", { name: "Dokumen A", title: "Dokumen A" });
  assert.equal(share.name, "dokumen-a");
  assert.ok(/^[A-Za-z0-9_-]{20,}$/.test(share.token));
  const resolved = store.resolve("dokumen-a", share.token);
  assert.ok(resolved);
  assert.equal(resolved.folder, "documents/dokumen-a");
  assert.equal(resolved.title, "Dokumen A");

  // Wrong token rejected
  assert.equal(store.resolve("dokumen-a", "wrong-token-value-12345678"), null);
  assert.equal(store.resolve("unknown-workspace", share.token), null);

  // Name normalization: spaces/underscores -> dashes, uppercase -> lowercase
  const spaced = store.issueToken(userA, "x", { name: "Dokumen A" });
  assert.equal(spaced.name, "dokumen-a"); // same as line 17's share, overwrites
  assert.throws(() => store.issueToken(userA, "x", { name: "!!!" }), /Workspace name/);
  const share2 = store.issueToken(userA, "x", { name: "My_Workspace2" });
  assert.equal(share2.name, "my-workspace2");

  // Re-issue overwrites folder + token (same name)
  const rotated = store.issueToken(userA, "documents/other", { name: "dokumen-a" });
  assert.equal(store.resolve("dokumen-a", share.token), null); // old token dead
  assert.equal(store.resolve("dokumen-a", rotated.token).folder, "documents/other");

  // Rotate token
  const next = store.rotateToken(userA, "dokumen-a");
  assert.equal(store.resolve("dokumen-a", rotated.token), null);
  assert.ok(store.resolve("dokumen-a", next.token));

  // Ownership checks
  assert.throws(() => store.rotateToken(userB, "dokumen-a"), /not found/);
  assert.throws(() => store.revoke(userB, "dokumen-a"), /not found/);

  // List + revoke
  assert.equal(store.list(userA).length, 2); // dokumen-a + my-workspace2
  const revoked = store.revoke(userA, "dokumen-a");
  assert.equal(revoked.revoked, "dokumen-a");
  assert.equal(store.resolve("dokumen-a", next.token), null);

  console.log("share-links self-check passed");
} finally {
  store?.close();
  try { rmSync(dir, { recursive: true, force: true }); } catch {}
}
