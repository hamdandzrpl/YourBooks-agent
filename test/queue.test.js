import assert from "node:assert/strict";
import test from "node:test";
import { createOutboundRateLimiter, createSerialQueue } from "../src/queue.js";

test("one serial queue processes messages from different users in FIFO order", async () => {
  const queue = createSerialQueue();
  const order = [];
  const first = queue.enqueue(async () => {
    order.push("first-start");
    await new Promise((resolve) => setTimeout(resolve, 10));
    order.push("first-end");
  });
  const second = queue.enqueue(async () => order.push("second"));
  const third = queue.enqueue(async () => order.push("third"));

  await Promise.all([first, second, third]);
  assert.deepEqual(order, ["first-start", "first-end", "second", "third"]);
  assert.equal(queue.pendingCount, 0);
});

test("global sender waits only for the remaining five-second minimum", async () => {
  let clock = 1000;
  const sentAt = [];
  const limiter = createOutboundRateLimiter({
    minIntervalMs: 5000,
    now: () => clock,
    sleep: async (ms) => { clock += ms; },
    send: async (value) => {
      sentAt.push({ value, at: clock });
      return value;
    },
  });

  await limiter("a");
  await limiter("b");
  clock += 9000;
  await limiter("c");
  assert.deepEqual(sentAt, [
    { value: "a", at: 1000 },
    { value: "b", at: 6000 },
    { value: "c", at: 15000 },
  ]);
});

test("outbound limiter serializes concurrent sends and throttles failures too", async () => {
  let clock = 1000;
  const sentAt = [];
  const limiter = createOutboundRateLimiter({
    minIntervalMs: 5000,
    now: () => clock,
    sleep: async (ms) => { clock += ms; },
    send: async (value) => {
      sentAt.push({ value, at: clock });
      if (value === "fail") throw new Error("send failed");
      return value;
    },
  });
  const results = await Promise.allSettled([limiter("fail"), limiter("next")]);

  assert.equal(results[0].status, "rejected");
  assert.equal(results[1].status, "fulfilled");
  assert.equal(sentAt[1].at - sentAt[0].at, 5000);
});