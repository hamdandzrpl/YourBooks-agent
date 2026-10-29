const toolCategories = new Set(["finance", "files", "office", "excel", "memoryTool", "historyTool", "utility", "web"]);
const financeToolNames = new Set([
  "list_books", "create_book", "select_book", "move_transaction_to_book",
  "record_transaction", "record_transactions", "list_recent_transactions",
  "update_transaction", "void_transaction", "list_pending_drafts",
  "approve_transaction_draft", "cancel_transaction_draft", "get_finance_summary",
  "export_transactions",
]);

export function workflowToolCategory(name) {
  if (financeToolNames.has(name)) return "finance";
  if (name === "mcp_filesystem_workspace_files") return "files";
  if (["spreadsheet", "word_document", "mcp_filesystem_spreadsheet", "mcp_filesystem_word_document"].includes(name)) return "office";
  if (typeof name === "string" && name.startsWith("mcp_excel_")) return "excel";
  if (name === "memory") return "memoryTool";
  if (name === "search_conversation_history") return "historyTool";
  if (name === "search_web") return "web";
  return "utility";
}

const serviceCategoryLabels = {
  finance: "Finance Ledger",
  files: "File Explorer MCP",
  office: "Office documents",
  excel: "Excel MCP",
  memoryTool: "Memory Tool",
  historyTool: "History Search Tool",
  utility: "Calculator & time",
  web: "Web Search",
};

const nodeIds = ["message", "attachment", "workspace", "history", "memory", "skills", "context", "agent", "model", "decision", ...toolCategories, "reply", "send"];
const edgeIds = [
  "message-history", "message-attachment", "attachment-workspace", "workspace-context", "message-context", "history-context", "memory-context", "skills-agent",
  "context-agent", "agent-model", "model-agent", "agent-decision",
  ...[...toolCategories].flatMap((category) => [`decision-${category}`, `${category}-agent`]),
  "decision-reply", "reply-send",
];

function createNodes() {
  return [
    { id: "message", label: "Pesan masuk", status: "done" },
    { id: "attachment", label: "Lampiran opsional", status: "pending" },
    { id: "workspace", label: "File Explorer Workspace", status: "pending" },
    { id: "history", label: "Conversation History", status: "active" },
    { id: "memory", label: "Memory Store", status: "pending" },
    { id: "skills", label: "System Skills", status: "pending" },
    { id: "context", label: "Context Builder", status: "pending" },
    { id: "agent", label: "AI Agent", status: "pending" },
    { id: "model", label: "AI Chat Model", status: "pending" },
    { id: "decision", label: "Tool Router", status: "pending" },
    ...[...toolCategories].map((id) => ({ id, label: serviceCategoryLabels[id], status: "pending" })),
    { id: "reply", label: "Balasan", status: "pending" },
    { id: "send", label: "Kirim WhatsApp", status: "pending" },
  ];
}

function createEdges() {
  return edgeIds.map((id) => ({ id, status: "pending" }));
}

export class WorkflowLiveState {
  constructor({ ownerId = "", now = () => new Date() } = {}) {
    this.ownerId = ownerId;
    this.now = now;
    this.run = null;
  }

  begin(sender) {
    if (!this.ownerId || sender !== this.ownerId) return false;
    this.run = {
      status: "running",
      nodes: createNodes(),
      edges: createEdges(),
      activeNode: "history",
      activeEdge: "message-history",
      activeLabel: "Memuat conversation history",
      queuedCategories: [],
      startedAt: this.now().toISOString(),
      updatedAt: this.now().toISOString(),
    };
    return true;
  }

  update(sender, event) {
    if (!this.run || sender !== this.ownerId || !event || typeof event.type !== "string") return false;
    const updateNode = (id, status) => {
      const node = this.run.nodes.find((item) => item.id === id);
      if (node) node.status = status;
    };
    const updateEdge = (id, status) => {
      const edge = this.run.edges.find((item) => item.id === id);
      if (edge) edge.status = status;
    };
    const clearActiveEdge = (status = "done") => {
      const active = this.run.edges.find((item) => item.id === this.run.activeEdge);
      if (active && active.status === "active") active.status = status;
      this.run.activeEdge = null;
    };
    if (event.type === "attachment_start") {
      updateNode("attachment", "active");
      this.activate("attachment", "Menyimpan lampiran", "message-attachment");
    } else if (event.type === "workspace_save_start") {
      updateNode("attachment", "done");
      updateEdge("message-attachment", "done");
      updateNode("workspace", "active");
      this.activate("workspace", "Menyimpan ke File Explorer", "attachment-workspace");
    } else if (event.type === "attachment_end") {
      updateNode("attachment", event.success ? "done" : "failed");
      if (event.success) {
        updateNode("workspace", "done");
        updateEdge("attachment-workspace", "done");
        clearActiveEdge();
      } else {
        updateEdge("message-attachment", "failed");
        updateNode("workspace", "failed");
        updateEdge("attachment-workspace", "failed");
        clearActiveEdge("failed");
      }
    } else if (event.type === "history_loaded") {
      updateNode("history", "done");
      updateEdge("message-history", "done");
      this.activate("memory", "Memuat Memory Store");
    } else if (event.type === "memory_loaded") {
      updateNode("memory", "done");
      updateEdge("memory-context", "done");
      if (this.run.nodes.find((item) => item.id === "workspace")?.status === "done") updateEdge("workspace-context", "done");
      this.activate("skills", "Memuat System Skills");
    } else if (event.type === "skills_loaded") {
      updateNode("skills", "done");
      updateEdge("skills-agent", "done");
      this.activate("context", "Menyusun konteks agent");
    } else if (event.type === "context_ready") {
      updateNode("context", "done");
      updateEdge("message-context", "done");
      if (this.run.nodes.find((item) => item.id === "attachment")?.status === "pending") {
        updateNode("attachment", "skipped");
        updateEdge("message-attachment", "skipped");
        updateNode("workspace", "skipped");
      }
      this.activate("agent", "Menyiapkan AI Agent", "context-agent");
    } else if (event.type === "model_start") {
      clearActiveEdge();
      updateEdge("agent-model", "active");
      updateNode("agent", "done");
      updateNode("model", "active");
      this.run.activeNode = "model";
      this.run.activeEdge = "agent-model";
      this.run.activeLabel = "Menunggu AI Chat Model";
    } else if (event.type === "model_route") {
      updateNode("model", "done");
      clearActiveEdge();
      updateEdge("model-agent", "done");
      this.activate("decision", "AI Agent mengevaluasi respons", "agent-decision");
      if (event.route === "tools") {
        updateNode("decision", "done");
        updateEdge("agent-decision", "done");
        this.run.queuedCategories = [...new Set((event.categories ?? []).filter((category) => toolCategories.has(category)))];
        for (const category of this.run.queuedCategories) {
          updateNode(category, "queued");
          updateEdge(`decision-${category}`, "queued");
        }
        this.run.activeNode = null;
        this.run.activeEdge = null;
        this.run.activeLabel = "Mengarahkan ke tool yang dipilih";
      } else if (event.route === "reply") {
        updateNode("decision", "done");
        updateEdge("agent-decision", "done");
        for (const category of toolCategories) {
          const serviceNode = this.run.nodes.find((item) => item.id === category);
          const serviceEdge = this.run.edges.find((item) => item.id === `decision-${category}`);
          if (serviceNode && ["pending", "queued"].includes(serviceNode.status)) serviceNode.status = "skipped";
          if (serviceEdge && ["pending", "queued"].includes(serviceEdge.status)) serviceEdge.status = "skipped";
        }
        this.activate("reply", "Model memilih balasan langsung", "decision-reply");
      } else {
        return false;
      }
    } else if (event.type === "model_end") {
      updateNode("model", event.success ? "done" : "failed");
      clearActiveEdge(event.success ? "done" : "failed");
      this.run.activeNode = null;
      this.run.activeEdge = null;
      this.run.activeLabel = event.success ? "Memeriksa keputusan model" : "Model AI mengalami kendala";
    } else if (event.type === "tool_start") {
      const category = toolCategories.has(event.category) ? event.category : "utility";
      clearActiveEdge();
      updateNode("decision", "done");
      for (const queued of this.run.queuedCategories) if (queued !== category) updateNode(queued, "queued");
      updateNode(category, "active");
      this.run.currentToolCategory = category;
      this.activate(category, typeof event.label === "string" ? event.label.slice(0, 80) : serviceCategoryLabels[category], `decision-${category}`);
    } else if (event.type === "tool_end") {
      const category = toolCategories.has(event.category) ? event.category : this.run.currentToolCategory ?? "utility";
      updateNode(category, event.success ? "done" : "failed");
      updateEdge(`decision-${category}`, event.success ? "done" : "failed");
      this.activate("agent", event.success ? "AI Agent menerima hasil tool" : "AI Agent menerima error tool", `${category}-agent`);
    } else if (event.type === "reply_start") {
      updateNode("model", "done");
      updateNode("decision", "done");
      this.activate("reply", "Menyusun balasan", "decision-reply");
    } else if (event.type === "reply_sent") {
      updateNode("reply", "done");
      updateEdge("decision-reply", "done");
      this.activate("send", "Mengirim pesan ke WhatsApp", "reply-send");
    } else if (event.type === "message_complete") {
      for (const id of nodeIds) {
        const node = this.run.nodes.find((item) => item.id === id);
        if (node.status === "active") node.status = "done";
        else if (node.status === "pending" || node.status === "queued") node.status = "skipped";
      }
      for (const edge of this.run.edges) {
        if (edge.status === "active") edge.status = "done";
        else if (edge.status === "pending" || edge.status === "queued") edge.status = "skipped";
      }
      this.run.status = "completed";
      this.run.activeNode = null;
      this.run.activeEdge = null;
      this.run.activeLabel = "Selesai";
    } else if (event.type === "message_failed") {
      const node = this.run.nodes.find((item) => item.id === this.run.activeNode);
      if (node) node.status = "failed";
      const edge = this.run.edges.find((item) => item.id === this.run.activeEdge);
      if (edge) edge.status = "failed";
      this.run.status = "failed";
      this.run.activeNode = null;
      this.run.activeEdge = null;
      this.run.activeLabel = "Proses gagal";
    } else {
      return false;
    }
    this.run.updatedAt = this.now().toISOString();
    return true;
  }

  activate(id, label, edgeId = null) {
    this.run.activeNode = id;
    this.run.activeEdge = edgeId;
    this.run.activeLabel = label;
    const node = this.run.nodes.find((item) => item.id === id);
    if (node) node.status = "active";
    if (edgeId) {
      const edge = this.run.edges.find((item) => item.id === edgeId);
      if (edge) edge.status = "active";
    }
  }

  snapshot() {
    if (!this.run) return {
      status: "idle",
      nodes: createNodes().map((node) => ({ ...node, status: "pending" })),
      edges: createEdges(),
      activeNode: null,
      activeEdge: null,
      activeLabel: "Menunggu aktivitas",
      startedAt: null,
      updatedAt: null,
    };
    return structuredClone(this.run);
  }
}
