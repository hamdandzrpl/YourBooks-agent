import assert from "node:assert/strict";
import test from "node:test";
import { ProgressTimeline } from "../src/progress.js";

test("progress timeline animates one message and marks completed steps", async () => {
  const sent = [];
  const edits = [];
  let tick;
  let cleared = false;
  const messageKey = { remoteJid: "628123@s.whatsapp.net", fromMe: true, id: "progress-1" };
  const timeline = new ProgressTimeline({
    to: "628123",
    editIntervalMs: 0,
    sendText: async (message) => {
      sent.push(message);
      return { key: messageKey };
    },
    editText: async (message) => edits.push(message),
    setIntervalImpl: (callback) => { tick = callback; return "timer"; },
    clearIntervalImpl: (timer) => { cleared = timer === "timer"; },
  });

  await timeline.start();
  assert.equal(sent.length, 1);
  assert.match(sent[0].text, /Memproses permintaan \./);
  tick();
  await timeline.requestEdit();
  assert.match(edits.at(-1).text, /Memproses permintaan \. \./);
  await timeline.update({ type: "step_start", label: "Mencatat 3 transaksi" });
  assert.match(edits.at(-1).text, /> Memproses permintaan/);
  assert.match(edits.at(-1).text, /Mencatat 3 transaksi \./);
  assert.equal(edits.at(-1).messageKey, messageKey);

  await timeline.update({ type: "step_end", success: true });
  await timeline.update({ type: "step_start", label: "Menyiapkan laporan" });
  assert.match(edits.at(-1).text, /> Mencatat 3 transaksi/);
  await timeline.complete();

  assert.equal(cleared, true);
  assert.match(edits.at(-1).text, /> Menyiapkan laporan/);
  assert.match(edits.at(-1).text, /> Selesai/);
  assert.doesNotMatch(edits.at(-1).text, /reasoning|chain.of.thought|record_transaction\(/i);
});

test("progress timeline marks a failed step without exposing internal details", async () => {
  const edits = [];
  const timeline = new ProgressTimeline({
    to: "628123",
    editIntervalMs: 0,
    sendText: async () => ({ key: { remoteJid: "628123@s.whatsapp.net", fromMe: true, id: "progress-2" } }),
    editText: async (message) => edits.push(message.text),
    setIntervalImpl: () => 1,
    clearIntervalImpl: () => {},
  });
  await timeline.start();
  await timeline.update({ type: "step_start", label: "Mengekspor data keuangan" });
  await timeline.update({ type: "step_end", success: false });
  await timeline.fail();

  assert.match(edits.at(-1), /! Mengekspor data keuangan/);
  assert.match(edits.at(-1), /! Terjadi kendala/);
  assert.doesNotMatch(edits.at(-1), /secret|token|api.?key/i);
});

test("progress timeline appears only when a real step starts", async () => {
  const sent = [];
  const timeline = new ProgressTimeline({
    to: "628123",
    editIntervalMs: 0,
    sendText: async (message) => { sent.push(message); return { key: { id: "lazy-progress" } }; },
    editText: async () => {},
    setIntervalImpl: () => 1,
    clearIntervalImpl: () => {},
  });

  await timeline.complete();
  assert.equal(sent.length, 0);

  const activeTimeline = new ProgressTimeline({
    to: "628123",
    sendText: async (message) => { sent.push(message); return { key: { id: "lazy-progress" } }; },
    editText: async () => {},
    setIntervalImpl: () => 1,
    clearIntervalImpl: () => {},
  });
  await activeTimeline.update({ type: "step_start", label: "Mencatat transaksi" });
  assert.equal(sent.length, 1);
  assert.match(sent[0].text, /Mencatat transaksi/);
  assert.doesNotMatch(sent[0].text, /Memproses permintaan/);
  await activeTimeline.complete();
});