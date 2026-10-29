import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { loadSystemSkills } from "../src/skills.js";

test("missing skills directory does not block agent startup or replies", (t) => {
  const parent = mkdtempSync(join(tmpdir(), "agent-missing-skills-"));
  t.after(() => rmSync(parent, { recursive: true, force: true }));
  assert.deepEqual(loadSystemSkills(join(parent, "skills")), []);
});
