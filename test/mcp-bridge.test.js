import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { FinanceStore } from "../src/finance/store.js";
import { OfficeFiles } from "../src/files/office.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { connectFileWorkspaceMcp } from "../src/files/mcp-server.js";

test("filesystem MCP exposes only the authenticated user's workspace tools", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "filesystem-mcp-"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const officeFiles = new OfficeFiles(fileWorkspace);
  const bridge = await connectFileWorkspaceMcp({
    fileWorkspace,
    officeFiles,
    userId: "6281234567890",
  });
  t.after(async () => {
    await bridge.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  await fileWorkspace.createTextFile("6281234567890", "private/notes.txt", "owner file");
  await fileWorkspace.createTextFile("6289999999999", "private/secret.txt", "other user file");

  assert.ok(bridge.definitions.some((tool) => tool.function.name === "mcp_filesystem_workspace_files"));
  const listing = await bridge.call("mcp_filesystem_workspace_files", { action: "list" });
  assert.match(listing.content, /notes\.txt/);
  assert.doesNotMatch(listing.content, /secret\.txt/);

  const created = await bridge.call("mcp_filesystem_workspace_files", {
    action: "create_text",
    path: "exports/result.txt",
    content: "created through MCP",
  });
  assert.equal(created.isError, false);
  assert.match(created.content, /exports\/result\.txt/);

  const traversal = await bridge.call("mcp_filesystem_workspace_files", {
    action: "create_text",
    path: "../outside.txt",
    content: "blocked",
  });
  assert.equal(traversal.isError, true);
  assert.equal((await fileWorkspace.listFiles("6289999999999")).length, 1);
});