import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { runAgent } from "../src/agent.js";
import { connectFileWorkspaceMcp } from "../src/files/mcp-server.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { formatSystemSkills, loadSystemSkills, setSystemSkillEnabled } from "../src/skills.js";
import { getCurrentDateTime, runTool } from "../src/tools.js";

test("system skill registry loads versioned accounting skills and honors enabled state", () => {
  const skills = loadSystemSkills();
  assert.deepEqual(skills.map((skill) => skill.id), [
    "dashboard-workbooks",
    "evidence-intake",
    "exception-escalation",
    "period-reporting",
    "professional-spreadsheets",
    "reconciliation-review",
    "transaction-capture",
    "workspace-documents",
  ]);
  // Skills that have been revised since 1.0.0 are listed explicitly; everything
  // else must still be the original version.
  const revised = new Set(["workspace-documents", "dashboard-workbooks"]);
  assert.ok(skills.filter((skill) => !revised.has(skill.id)).every((skill) => skill.version === "1.0.0"));
  assert.ok(skills.find((skill) => skill.id === "workspace-documents"));

  const prompt = formatSystemSkills(skills);
  assert.match(prompt, /Pencatatan Transaksi/);
  assert.match(prompt, /Workbook Keuangan Profesional/);
  assert.match(prompt, /Struktur Workbook Dashboard/);
  assert.match(prompt, /belum menyediakan impor mutasi bank/i);
  assert.match(prompt, /cannot authorize new tools/i);
  assert.equal(formatSystemSkills(skills.map((skill) => ({ ...skill, enabled: false }))), "");
});

test("system skill toggle persists enabled state in its manifest", () => {
  const directory = mkdtempSync(join(tmpdir(), "agent-skills-"));
  try {
    writeFileSync(join(directory, "sample.json"), JSON.stringify({
      id: "sample-skill",
      name: "Sample Skill",
      description: "A test skill",
      version: "1.0.0",
      enabled: true,
      instructions: ["Follow the test instruction."],
    }));
    const updated = setSystemSkillEnabled("sample-skill", false, directory);
    assert.equal(updated.enabled, false);
    assert.equal(loadSystemSkills(directory)[0].enabled, false);
    assert.throws(() => setSystemSkillEnabled("missing-skill", true, directory), /not found/i);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});

test("date-time tool formats the configured timezone", () => {
  const result = getCurrentDateTime("Asia/Jakarta", new Date("2026-10-14T00:00:00.000Z"));
  assert.equal(result.date, "2026-10-14");
  assert.equal(result.time, "07:00:00");
  assert.equal(result.timeZone, "Asia/Jakarta");
});

test("calculator tool computes arithmetic and rejects division by zero", async () => {
  assert.deepEqual(await runTool("calculate", "UTC", { operation: "multiply", a: 6, b: 7 }), {
    operation: "multiply",
    a: 6,
    b: 7,
    result: 42,
  });
  await assert.rejects(
    () => runTool("calculate", "UTC", { operation: "divide", a: 1, b: 0 }),
    /Cannot divide by zero/,
  );
});

test("agent uses XKIRO Chat Completions to run a function tool", async () => {
  const requests = [];
  const responses = [
    {
      choices: [{ message: {
        role: "assistant",
        content: null,
        tool_calls: [{
          id: "call-1",
          type: "function",
          function: { name: "get_current_datetime", arguments: "{}" },
        }],
      } }],
    },
    { choices: [{ message: { role: "assistant", content: "Sekarang pukul 07.00 WIB." } }] },
  ];

  const answer = await runAgent({
    apiKey: "test-key",
    model: "test-model",
    reasoningEffort: "low",
    timeZone: "Asia/Jakarta",
    text: "Jam berapa sekarang?",
    runToolImpl: () => ({
      timeZone: "Asia/Jakarta",
      date: "2026-10-14",
      time: "07:00:00",
      iso: "2026-10-14T00:00:00.000Z",
    }),
    fetchImpl: async (url, options) => {
      requests.push({ url, body: JSON.parse(options.body) });
      return { ok: true, json: async () => responses.shift() };
    },
  });

  assert.equal(answer, "Sekarang pukul 07.00 WIB.");
  assert.equal(requests.length, 2);
  assert.equal(requests[0].url, "https://api.xkiro.com/v1/chat/completions");
  assert.equal(requests[0].body.reasoning_effort, "low");
  assert.equal(requests[0].body.tools.length, 19);
  assert.equal(requests[0].body.tools.some((tool) => tool.function.name === "create_finance_report"), false);
  assert.equal(requests[0].body.tools.some((tool) => tool.function.name === "approve_transaction_draft"), true);
  assert.equal(requests[0].body.tools.some((tool) => tool.function.name === "cancel_transaction_draft"), true);
  assert.equal(requests[0].body.tools.some((tool) => ["workspace_files", "spreadsheet", "word_document"].includes(tool.function.name)), false);
  assert.doesNotMatch(requests[0].body.messages[0].content, /reply exactly SIMPAN/);
  assert.doesNotMatch(requests[0].body.messages[0].content, /post-save review|ask them to check whether the details are correct/i);
  assert.equal(requests[0].body.tools.some((tool) => tool.function.name === "search_web"), true);
  assert.equal(requests[1].body.messages.at(-1).role, "tool");
  assert.deepEqual(JSON.parse(requests[1].body.messages.at(-1).content), {
    timeZone: "Asia/Jakarta",
    date: "2026-10-14",
    time: "07:00:00",
    iso: "2026-10-14T00:00:00.000Z",
  });
});

test("agent returns a normal answer in one model turn when no tool is needed", async () => {
  let requestCount = 0;
  let requestBody;
  const progressEvents = [];
  const answer = await runAgent({
    apiKey: "test-key",
    model: "test-model",
    timeZone: "Asia/Jakarta",
    text: "halo",
    onProgress: (event) => progressEvents.push(event),
    fetchImpl: async (_url, options) => {
      requestCount += 1;
      requestBody = JSON.parse(options.body);
      return { ok: true, json: async () => ({ choices: [{ message: { role: "assistant", content: "Halo! Ada yang bisa saya bantu?" } }] }) };
    },
  });

  assert.equal(answer, "Halo! Ada yang bisa saya bantu?");
  assert.equal(requestCount, 1);
  assert.equal(requestBody.tool_choice, "auto");
  assert.match(requestBody.messages[0].content, /greetings.*directly in one response without calling tools/i);
  assert.deepEqual(progressEvents, []);
});

test("agent retries one upstream capacity response before failing", async () => {
  let requestCount = 0;
  const answer = await runAgent({
    apiKey: "test-key",
    model: "test-model",
    timeZone: "Asia/Jakarta",
    text: "halo",
    fetchImpl: async () => {
      requestCount += 1;
      if (requestCount === 1) {
        return { ok: false, status: 503, json: async () => ({ error: { message: "upstream at capacity" } }) };
      }
      return { ok: true, status: 200, json: async () => ({ choices: [{ message: { role: "assistant", content: "Halo!" } }] }) };
    },
  });

  assert.equal(answer, "Halo!");
  assert.equal(requestCount, 2);
});

test("agent retries a timed-out model request before failing", async () => {
  let requestCount = 0;
  const answer = await runAgent({
    apiKey: "test-key",
    model: "test-model",
    timeZone: "Asia/Jakarta",
    text: "halo",
    requestTimeoutMs: 60,
    fetchImpl: async () => {
      requestCount += 1;
      if (requestCount === 1) {
        // Simulate an upstream hang that the timeout aborts.
        return await new Promise((_, reject) => {
          setTimeout(() => reject(new Error("aborted by caller")), 500);
        });
      }
      return { ok: true, status: 200, json: async () => ({ choices: [{ message: { role: "assistant", content: "Halo!" } }] }) };
    },
  });

  assert.equal(answer, "Halo!");
  assert.equal(requestCount, 2);
});

test("natural-language file requests execute through the workspace function tool", async (t) => {
  const directory = mkdtempSync(join(tmpdir(), "agent-workspace-call-"));
  const fileWorkspace = new FileWorkspace({ dataDirectory: directory });
  t.after(() => {
    fileWorkspace.close();
    rmSync(directory, { recursive: true, force: true });
  });
  const saved = await fileWorkspace.createTextFile("6281234567890", "Catatan/anggaran.md", "Anggaran bulanan");
  const mcpSource = await connectFileWorkspaceMcp({ fileWorkspace, officeFiles: {}, userId: "6281234567890" });
  const requests = [];
  const responses = [
    { choices: [{ message: {
      role: "assistant",
      content: null,
      tool_calls: [{
        id: "call-files-1",
        type: "function",
          function: { name: "mcp_filesystem_workspace_files", arguments: JSON.stringify({ action: "list", folder: "Catatan" }) },
      }],
    } }] },
    { choices: [{ message: { role: "assistant", content: "Saya menemukan Catatan/anggaran.md." } }] },
  ];

  const answer = await runAgent({
    apiKey: "test-key",
    model: "test-model",
    timeZone: "Asia/Jakarta",
    text: "Tolong lihat catatan anggaran yang kusimpan di folder Catatan",
    fileWorkspace,
    officeFiles: {},
    mcpSources: [mcpSource],
    userId: "6281234567890",
    fetchImpl: async (_url, options) => {
      requests.push(JSON.parse(options.body));
      return { ok: true, json: async () => responses.shift() };
    },
  });

  assert.equal(answer, "Saya menemukan Catatan/anggaran.md.");
  assert.equal(requests[0].tools.some((tool) => tool.function.name === "mcp_filesystem_workspace_files"), true);
  const toolResult = JSON.parse(requests[1].messages.at(-1).content);
  const files = JSON.parse(toolResult.content).content ? JSON.parse(JSON.parse(toolResult.content).content) : JSON.parse(toolResult.content);
  assert.equal(files[0].id, saved.id);
  assert.equal(files[0].path, "Catatan/anggaran.md");
  await mcpSource.close();
});

test("agent exposes namespaced MCP tools and routes calls to their source", async () => {
  const requests = [];
  const calls = [];
  const mcpSource = {
    definitions: [{
      type: "function",
      function: {
        name: "mcp_example_list_workbooks",
        description: "List workbooks",
        parameters: { type: "object", properties: {}, required: [], additionalProperties: false },
      },
    }],
    async call(name, args) {
      calls.push({ name, args });
      return { content: "[]", isError: false };
    },
  };
  const responses = [
    { choices: [{ message: {
      role: "assistant",
      content: null,
      tool_calls: [{
        id: "call-example-1",
        type: "function",
        function: { name: "mcp_example_list_workbooks", arguments: "{}" },
      }],
    } }] },
    { choices: [{ message: { role: "assistant", content: "Workbook terbaca." } }] },
  ];

  const answer = await runAgent({
    apiKey: "test-key",
    model: "test-model",
    timeZone: "Asia/Jakarta",
    text: "Tampilkan workbook saya",
    mcpSources: [mcpSource],
    fetchImpl: async (_url, options) => {
      requests.push(JSON.parse(options.body));
      return { ok: true, json: async () => responses.shift() };
    },
  });

  assert.equal(answer, "Workbook terbaca.");
  assert.equal(requests[0].tools.some((tool) => tool.function.name === "mcp_example_list_workbooks"), true);
  assert.deepEqual(calls, [{ name: "mcp_example_list_workbooks", args: {} }]);
});

test("agent times out a stalled provider request", async () => {
  await assert.rejects(() => runAgent({
    apiKey: "test-key",
    model: "test-model",
    timeZone: "UTC",
    text: "Halo",
    requestTimeoutMs: 10,
    fetchImpl: (_url, { signal }) => new Promise((resolve, reject) => {
      signal.addEventListener("abort", () => reject(signal.reason), { once: true });
    }),
  }), /request timed out after 10 ms/);
});

test("web search uses the configured search endpoint", async () => {
  let request;
  const results = await runTool("search_web", "UTC", { query: "weather Jakarta", max_results: 3 }, {
    apiKey: "test-key",
    searchUrl: "https://api.example.com/v1/search",
    searchModel: "example/web-search",
    fetchImpl: async (url, options) => {
      request = { url, options };
      return {
        ok: true,
        json: async () => ({ results: [{ title: "Weather", url: "https://example.com", snippet: "Sunny" }] }),
      };
    },
  });
  assert.equal(request.url, "https://api.example.com/v1/search");
  assert.equal(JSON.parse(request.options.body).model, "example/web-search");
  assert.deepEqual(results, [{
    title: "Weather",
    url: "https://example.com",
    snippet: "Sunny",
    source: undefined,
    publishedDate: undefined,
  }]);
});

test("web search degrades clearly when the provider has no search endpoint", async () => {
  await assert.rejects(
    () => runTool("search_web", "UTC", { query: "weather Jakarta" }, { apiKey: "test-key" }),
    /not configured for this provider/,
  );
});

test("web search times out a stalled provider request", async () => {
  await assert.rejects(() => runTool("search_web", "UTC", { query: "weather Jakarta" }, {
    apiKey: "test-key",
    searchUrl: "https://api.example.com/v1/search",
    requestTimeoutMs: 10,
    fetchImpl: (_url, { signal }) => new Promise((resolve, reject) => {
      signal.addEventListener("abort", () => reject(signal.reason), { once: true });
    }),
  }), /request timed out after 10 ms/);
});


test("agent never sends the same file twice in one turn", async () => {
  const sent = [];
  const responses = [
    { choices: [{ message: { role: "assistant", content: null, tool_calls: [
      { id: "c1", type: "function", function: { name: "create_finance_report", arguments: "{\"send\":true}" } },
    ] } }] },
    { choices: [{ message: { role: "assistant", content: null, tool_calls: [
      { id: "c2", type: "function", function: { name: "create_finance_report", arguments: "{\"send\":true}" } },
    ] } }] },
    { choices: [{ message: { role: "assistant", content: "Laporan dikirim." } }] },
  ];
  const answer = await runAgent({
    apiKey: "test-key",
    model: "test-model",
    timeZone: "Asia/Jakarta",
    text: "kirim laporan",
    sendFile: async (payload) => { sent.push(payload.filename); return { ok: true }; },
    runToolImpl: async (name, _tz, _args, context) => {
      await context.sendFile({ filename: "laporan.xlsx", mimeType: "application/x", bytes: Buffer.alloc(10) });
      return { filename: "laporan.xlsx", sent: true };
    },
    fetchImpl: async () => ({ ok: true, json: async () => responses.shift() }),
  });
  assert.equal(answer, "Laporan dikirim.");
  assert.deepEqual(sent, ["laporan.xlsx"]);
});

test("agent wraps up instead of failing when the tool budget runs out", async () => {
  let round = 0;
  const requests = [];
  const answer = await runAgent({
    apiKey: "test-key",
    model: "test-model",
    timeZone: "Asia/Jakarta",
    text: "rapikan file",
    runToolImpl: () => ({ ok: true }),
    fetchImpl: async (url, options) => {
      const body = JSON.parse(options.body);
      requests.push(body);
      round += 1;
      if (round > 24) return { ok: true, json: async () => ({ choices: [{ message: { role: "assistant", content: "Sebagian selesai, sisanya belum." } }] }) };
      return { ok: true, json: async () => ({ choices: [{ message: { role: "assistant", content: null, tool_calls: [
        { id: `c${round}`, type: "function", function: { name: "list_recent_transactions", arguments: "{}" } },
      ] } }] }) };
    },
  });
  assert.match(answer, /Sebagian selesai/);
  // The last request must forbid further tool calls and carry the wrap-up nudge.
  const last = requests.at(-1);
  assert.equal(last.tool_choice, "none");
  assert.ok(last.messages.some((m) => m.role === "system" && /BUDGET LANGKAH HABIS/.test(m.content)));
});

test("agent speaks to any [OI]-compatible provider base URL", async () => {
  const seen = [];
  const answer = await runAgent({
    apiKey: "vikey-key",
    model: "claude-sonnet-4",
    chatCompletionsUrl: "https://api.vikey.ai/v1/chat/completions",
    reasoningEffort: "high",
    maxOutputTokens: 800,
    timeZone: "Asia/Jakarta",
    text: "halo",
    fetchImpl: async (url, options) => {
      seen.push({ url, headers: options.headers, body: JSON.parse(options.body) });
      return { ok: true, json: async () => ({ choices: [{ message: { role: "assistant", content: "Hai!" } }] }) };
    },
  });
  assert.equal(answer, "Hai!");
  assert.equal(seen[0].url, "https://api.vikey.ai/v1/chat/completions");
  assert.equal(seen[0].headers.authorization, "Bearer vikey-key");
  assert.equal(seen[0].body.model, "claude-sonnet-4");
  assert.equal(seen[0].body.reasoning_effort, "high");
  assert.equal(seen[0].body.max_tokens, 800);
  assert.equal(seen[0].body.tool_choice, "auto");
});

test("agent uses provider-specific auth headers and omits unsupported reasoning fields", async () => {
  const seen = [];
  await runAgent({
    apiKey: "sk-ant",
    model: "claude-sonnet-4",
    chatCompletionsUrl: "https://api.anthropic.com/v1/chat/completions",
    reasoningEffort: "high",
    timeZone: "UTC",
    text: "halo",
    fetchImpl: async (url, options) => {
      seen.push({ headers: options.headers, body: JSON.parse(options.body) });
      return { ok: true, json: async () => ({ choices: [{ message: { role: "assistant", content: "ok" } }] }) };
    },
  });
  assert.equal(seen[0].headers["x-api-key"], "sk-ant");
  assert.equal(seen[0].headers.authorization, undefined);
  assert.equal(seen[0].headers["anthropic-version"], "2023-06-01");
  assert.equal("reasoning_effort" in seen[0].body, false);
});

test("agent understands provider responses that use function_call or content parts", async () => {
  let call = 0;
  const answer = await runAgent({
    apiKey: "k",
    model: "gemini-2.5-pro",
    chatCompletionsUrl: "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions",
    timeZone: "UTC",
    text: "jam berapa",
    runToolImpl: () => ({ date: "2026-10-20" }),
    fetchImpl: async () => {
      call += 1;
      if (call === 1) {
        return { ok: true, json: async () => ({ choices: [{ message: { content: [{ functionCall: { name: "get_current_datetime", args: {} } }] } }] }) };
      }
      return { ok: true, json: async () => ({ choices: [{ message: { content: [{ type: "text", text: "Tanggal 20." }] } }] }) };
    },
  });
  assert.equal(answer, "Tanggal 20.");
});
