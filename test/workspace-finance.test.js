import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { FinanceStore } from "../src/finance/store.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { ShareLinkStore } from "../src/share-links.js";
import { runTool } from "../src/tools.js";

test("create_book provisions editable web workspace before recording", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "finance-web-workspace-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const shareLinks = new ShareLinkStore({ dataDirectory });
  const userId = "6281234567890";
  const context = { financeStore, fileWorkspace, shareLinks, publicOrigin: "https://example.test", userId, defaultCurrency: "IDR" };
  t.after(() => {
    shareLinks.close();
    fileWorkspace.close();
    financeStore.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  assert.equal(financeStore.getActiveBook(userId), null);
  await assert.rejects(runTool("record_transaction", "Asia/Jakarta", { kind: "expense", amount: 20_000_000 }, context), /No finance workspace exists/);
  const workspace = await runTool("create_book", "Asia/Jakarta", { name: "Perkebunan" }, context);
  assert.equal(workspace.name, "Perkebunan");
  assert.match(workspace.workspace_url, /^https:\/\/example\.test\/w\/finance-b-[a-f0-9]{8}\?k=/);
  assert.equal(financeStore.getActiveBook(userId).id, workspace.id);

  const transaction = await runTool("record_transaction", "Asia/Jakarta", {
    kind: "expense", amount: 20_000_000, category: "Operasional", description: "Pupuk",
  }, { ...context, sourceMessageId: "sandbox.perkebunan" });
  assert.equal(transaction.book_id, workspace.id);
  assert.equal(shareLinks.list(userId).length, 1);
  const file = (await fileWorkspace.listFiles(userId)).find((item) => item.path === `finance-${workspace.id.toLowerCase()}/dataset.json`);
  assert.ok(file);
  const dataset = JSON.parse(await fileWorkspace.readTextFile(userId, file.id));
  assert.equal(dataset.rows.length, 1);
  assert.equal(dataset.rows[0][1], "Pupuk");
});
