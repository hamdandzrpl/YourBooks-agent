import assert from "node:assert/strict";
import test from "node:test";
import { WorkflowLiveState, workflowToolCategory } from "../src/workflow-live.js";

test("live workflow tracks the configured owner without exposing other senders", () => {
  let time = 0;
  const monitor = new WorkflowLiveState({
    ownerId: "6281234567890",
    now: () => new Date(1_800_000_000_000 + time++),
  });

  assert.equal(monitor.begin("6289999999999"), false);
  assert.equal(monitor.snapshot().status, "idle");
  assert.equal(monitor.begin("6281234567890"), true);
  assert.equal(monitor.update("6289999999999", { type: "tool_start", label: "private action" }), false);
  monitor.update("6281234567890", { type: "history_loaded" });
  monitor.update("6281234567890", { type: "memory_loaded" });
  monitor.update("6281234567890", { type: "skills_loaded" });
  monitor.update("6281234567890", { type: "context_ready" });
  monitor.update("6281234567890", { type: "model_start" });
  monitor.update("6281234567890", { type: "model_end", success: true });
  monitor.update("6281234567890", { type: "model_route", route: "tools", categories: ["finance"] });
  monitor.update("6281234567890", { type: "tool_start", category: "finance", label: "Mencatat transaksi" });

  const active = monitor.snapshot();
  assert.equal(active.status, "running");
  assert.equal(active.activeNode, "finance");
  assert.equal(active.activeLabel, "Mencatat transaksi");
  assert.equal(active.nodes.find((node) => node.id === "finance").status, "active");
  assert.equal(active.edges.find((edge) => edge.id === "decision-finance").status, "active");
  assert.equal(JSON.stringify(active).includes("private action"), false);

  monitor.update("6281234567890", { type: "tool_end", category: "finance", success: true });
  monitor.update("6281234567890", { type: "model_start" });
  monitor.update("6281234567890", { type: "model_end", success: true });
  monitor.update("6281234567890", { type: "model_route", route: "reply" });
  monitor.update("6281234567890", { type: "reply_start" });
  monitor.update("6281234567890", { type: "reply_sent" });
  monitor.update("6281234567890", { type: "message_complete" });
  assert.equal(monitor.snapshot().status, "completed");
  assert.ok(monitor.snapshot().nodes.every((node) => ["done", "skipped"].includes(node.status)));
  assert.equal(monitor.snapshot().edges.find((edge) => edge.id === "finance-agent").status, "done");
  assert.equal(monitor.snapshot().edges.find((edge) => edge.id === "decision-reply").status, "done");
  assert.equal(monitor.snapshot().edges.find((edge) => edge.id === "reply-send").status, "done");
});

test("live workflow routes an optional attachment through workspace storage", () => {
  const monitor = new WorkflowLiveState({ ownerId: "6281234567890" });
  monitor.begin("6281234567890");
  monitor.update("6281234567890", { type: "attachment_start" });
  monitor.update("6281234567890", { type: "workspace_save_start" });
  monitor.update("6281234567890", { type: "attachment_end", success: true });
  monitor.update("6281234567890", { type: "history_loaded" });
  monitor.update("6281234567890", { type: "memory_loaded" });
  monitor.update("6281234567890", { type: "skills_loaded" });
  monitor.update("6281234567890", { type: "context_ready" });
  const snapshot = monitor.snapshot();
  assert.equal(snapshot.nodes.find((node) => node.id === "attachment").status, "done");
  assert.equal(snapshot.edges.find((edge) => edge.id === "message-attachment").status, "done");
  assert.equal(snapshot.nodes.find((node) => node.id === "workspace").status, "done");
  assert.equal(snapshot.edges.find((edge) => edge.id === "attachment-workspace").status, "done");
  assert.equal(snapshot.edges.find((edge) => edge.id === "workspace-context").status, "done");
  assert.equal(snapshot.nodes.find((node) => node.id === "history").status, "done");
  assert.equal(snapshot.nodes.find((node) => node.id === "memory").status, "done");
  assert.equal(snapshot.nodes.find((node) => node.id === "skills").status, "done");
});

test("live workflow marks the active stage failed without retaining error text", () => {
  const monitor = new WorkflowLiveState({ ownerId: "6281234567890" });
  monitor.begin("6281234567890");
  monitor.update("6281234567890", { type: "history_loaded" });
  monitor.update("6281234567890", { type: "memory_loaded" });
  monitor.update("6281234567890", { type: "skills_loaded" });
  monitor.update("6281234567890", { type: "context_ready" });
  monitor.update("6281234567890", { type: "model_start" });
  monitor.update("6281234567890", { type: "message_failed", error: "private message text" });

  const snapshot = monitor.snapshot();
  assert.equal(snapshot.status, "failed");
  assert.equal(snapshot.nodes.find((node) => node.id === "model").status, "failed");
  assert.equal(JSON.stringify(snapshot).includes("private message text"), false);
});

test("live workflow preserves a failed tool while the model prepares a follow-up", () => {
  const monitor = new WorkflowLiveState({ ownerId: "6281234567890" });
  monitor.begin("6281234567890");
  monitor.update("6281234567890", { type: "history_loaded" });
  monitor.update("6281234567890", { type: "memory_loaded" });
  monitor.update("6281234567890", { type: "skills_loaded" });
  monitor.update("6281234567890", { type: "context_ready" });
  monitor.update("6281234567890", { type: "model_start" });
  monitor.update("6281234567890", { type: "model_end", success: true });
  monitor.update("6281234567890", { type: "model_route", route: "tools", categories: ["files"] });
  monitor.update("6281234567890", { type: "tool_start", category: "files", label: "Mencari file" });
  monitor.update("6281234567890", { type: "tool_end", category: "files", success: false });
  monitor.update("6281234567890", { type: "model_start" });
  assert.equal(monitor.snapshot().nodes.find((node) => node.id === "files").status, "failed");
});

test("tool names route to distinct workflow services", () => {
  assert.equal(workflowToolCategory("record_transaction"), "finance");
  assert.equal(workflowToolCategory("mcp_filesystem_workspace_files"), "files");
  assert.equal(workflowToolCategory("mcp_filesystem_spreadsheet"), "office");
  assert.equal(workflowToolCategory("mcp_excel_excel_write_to_sheet"), "excel");
  assert.equal(workflowToolCategory("memory"), "memoryTool");
  assert.equal(workflowToolCategory("search_conversation_history"), "historyTool");
  assert.equal(workflowToolCategory("calculate"), "utility");
  assert.equal(workflowToolCategory("search_web"), "web");
});
